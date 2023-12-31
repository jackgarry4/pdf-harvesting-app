from xml.etree.ElementTree import TreeBuilder
from bs4 import BeautifulSoup
from classes.Company import Company
from classes.PDF import PDF
import requests 
import re
import logging
import time


MAX_RETRIES = 3
stop_flag = False


def fetchDataFromURL(url, session, max_retries = MAX_RETRIES):

    """
    Make an HTTP request to the provided URL and return the HTML response as a string.

    Parameters:
    - url (str): The URL to make the HTTP request.
    - session (requests.Session): The requests session to use for the request.
    - max_retries (int): Maximum number of retries in case of failure.

    Returns:
    - dict: A dictionary with 'data' containing the HTML response text (if successful),
            and 'error' containing the error message (if an error occurs).
    """

    if stop_flag:
        return None
    #Client server has rare breaks in remote end connection, so need to retry in these instances
    for attempt in range(max_retries):
        waitTime = min(2**attempt, 30)
        try:
            pageData = session.get(url, timeout = (30,30))
            pageData.raise_for_status()
            return {'data': pageData.text, 'error': None}
        except requests.exceptions.HTTPError as errh:
            logging.warning(f'HTTP Error on attempt {attempt+1} : {errh}. Failed to fetch data from {url}. ')
        except requests.exceptions.ConnectionError as errc:
            logging.warning(f'Error Connecting on attempt {attempt+1}: {errc}. Failed to fetch data from {url}.')
        except requests.exceptions.Timeout as errt:
            logging.warning(f'Timeout Error on attempt {attempt+1}: {errt}. Failed to fetch data from {url}.')
        except requests.exceptions.RequestException as e: 
            logging.warning(f'Oops on attempt {attempt+1}: Something else {e}. Failed to fetch data from {url}.')
        time.sleep(waitTime)
    
    #Log the final error message when maximum retries are reached
    logging.error(f'Max retries reached.  Failed to fetch data from {url}')
    return {'data': None, 'error': f'Max retries reached.  Failed to fetch data from {url}'}
    

def extractCompanyName(doc):
    """
    Extract and return company name from HTML document if present.

    Parameters:
    - doc (BeautifulSoup): The BeautifulSoup object representing the HTML document.

    Returns:
        Either doc.h2.b.string or None:
            - doc.h2.b.string - Returned if doc is an active TransAmerica document.  Represents the name of the company that the doc HTML points to.
            - None - Returned if doc is not an active TransAmerica document and brings up attribute error when trying to find doc.h2.b.string
    """
    try:
        return doc.h2.b.string
    except AttributeError as e:
        logging.debug(f'Attribute error in extractCompanyName: {e}')
        return None
    

def extractUrlFromExpression(pdfUrlJS):
    """
    Extract URL using regex from JavaScript openWindow method call.

    Parameters:
    - pdfUrlJS (str): The JavaScript expression containing the openWindow method call.

    Returns:
    - str or None: The extracted URL or None if no match is found.
    """
    #Define regex to match the url
    urlRegex = r"openWindow\('([^']+)"

    #Search for the URL in the expression
    match = re.search(urlRegex, pdfUrlJS)

    #Check if a match is found
    if match and match.group(1):
        extractedUrl = match.group(1)
        logging.info(f'URL extracted successfully: {extractedUrl}')
        return extractedUrl
    else:
        logging.warning(f'No URL match is found in the expression {pdfUrlJS}')
        return None


def extractPDFs(company, doc):
    """
    Extract PDF information from the provided HTML document and add PDF objects to the given company.

    Parameters:
    - company (Company): The Company object to which the extracted PDFs will be added.
    - doc (BeautifulSoup): The BeautifulSoup object representing the HTML document.

    Returns:
    - tuple: A tuple containing a Company object (if successful) and an error message (if an error occurs).
      - If the extraction is successful, the first element is a Company object.
      - If there's an error during the extraction, the first element is None, and the second element is an error message.
    """
    try: 
        planDocsHTML = doc.find(id='planDocuments')
        if planDocsHTML:
            pdfAnchors = planDocsHTML.find_all('a')
            for a in pdfAnchors:
                try:
                    pdfUrlJS = a['href']
                    pdfUrl = extractUrlFromExpression(pdfUrlJS)
                    pdfTitleTag = a.find('li')
                    pdfTitle = pdfTitleTag.text if pdfTitleTag else "Untitled PDF"

                    company.add_pdf(pdfUrl, pdfTitle)
                except (AttributeError, KeyError, IndexError) as e:
                    logging.error(f"Error extracting PDF information {company}: {e}")
                    return None, f"Error extracting PDF information {company}: {e}"

            return company, None
    except Exception as e:
        logging.error(f"Error extracting PDFs {company}: {e}")
        return None, f"Error extracting PDFs: {e}"


def findValidDoc(homePageUrl, session, recursionDepth = 0):
    """
    Fetches and parses HTML content from a given URL, searching for a valid document with an account number.

    Args:
        homePageUrl (str): The URL of the webpage to be fetched and parsed.
        session: The session object used for making HTTP requests.
        recursion_depth (int, optional): The current depth of recursion (default is 0).

    Returns:
        BeautifulSoup object or None: 
            - Returns a BeautifulSoup object if a valid document is found.
            - Returns None if a valid document is not found after reaching the maximum recursion depth.
    """
    if stop_flag:
        return None
    try:
        if recursionDepth >= 3:
            logging.warning(f"Max recursion depth reached for {homePageUrl}. Aborting")
            return None
        
        pageHTML = fetchDataFromURL(homePageUrl, session)

        if pageHTML['error'] is not None:
            logging.error(pageHTML['error'])
            return None
            
        doc = BeautifulSoup(pageHTML["data"], 'html.parser')
        #Check to see if the account number exists and if not run fetchData again until it does
        accountNumberTable = doc.find('table', {'style': 'background-color:#F8F8F8;border-width: thin;border-collapse:collapse;border-color:#DCDCDC'})
        # Find the cell with the label "Account #:"
        accountLabelCell = accountNumberTable.find('td', text='Account #:')
        
        try:
            if accountLabelCell and accountLabelCell.find_next('td'):
                accountValue = accountLabelCell.find_next('td').text.strip()
                if accountValue:
                    logging.info("Valid doc found")
                    return doc
                else:
                    logging.warning(f"Error: Invalid doc found {homePageUrl} (No Account Value).  Rerunning for correct output...")
            else: 
                logging.warning(f"Error: Invalid doc found {homePageUrl} (Error is not none) Rerunning for correct output...")
        except Exception as e:
            logging.error(f"Error extracting account value {homePageUrl}: {e}")
        
        return findValidDoc(homePageUrl, session, recursionDepth + 1)  
            
    except Exception as e:
        logging.error(f"Error fetching data from URL {homePageUrl}: {e}")
        return findValidDoc(homePageUrl, session, recursionDepth + 1)



def scrape_pdf_links(homePageUrl, session):
    """
    Scrape PDF links from a TransAmerica (TA) page.

    Parameters:
    - homePageUrl (str): The URL of the TransAmerica home page.

    Returns:
    - tuple: A tuple containing a Company object (if successful) and an error message (if an error occurs).
      - If the URL is a valid TA page, the first element is a Company object representing the scraped information.
      - If the URL is not a valid TA page, the first element is None,
        and the second element is an error message describing why the page is considered invalid.
      - If there's an error during the HTTP request or parsing, the first element is None, and the second
        element is an error message describing the issue.
    """

    if stop_flag:
        logging.info("Program was closed")
        return None, f"Program was closed"
    
    logging.info(f"Scraping {homePageUrl}")
    doc = findValidDoc(homePageUrl, session)
    

    if doc is not None and not stop_flag:
        docTitle = doc.head.title.string
        if docTitle == "Fund and Fee Information":
            companyName = extractCompanyName(doc)
            company = Company(companyName)
            logging.info(f"Scraped {homePageUrl}")

            company, pdfErrorMessage = extractPDFs(company, doc)

            if pdfErrorMessage:
                return None, pdfErrorMessage
            elif not company.pdfs:
                return None, "This page does not contain any Plan Documents"
            else: 
                return company, None
        else:
            logging.info(f"{homePageUrl} does not appear to be a valid TA page")
            return None, "This page does not appear to be a valid TA Page"
    else:
        return None, f"Error fetching {homePageUrl}"


def stopProcessingScraper():
    global stop_flag
    stop_flag = True