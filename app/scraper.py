from xml.etree.ElementTree import TreeBuilder
from bs4 import BeautifulSoup
from classes.Company import Company
from classes.PDF import PDF
import requests 
import re
import logging
import time




def fetchDataFromURL(url, session, max_retries = 3):

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

    #Client server has rare breaks in remote end connection, so need to retry in these instances
    for attempt in range(max_retries):
        waitTime = min(2**attempt, 30)
        try:
            pageData = session.get(url, timeout = (30,30))
            pageData.raise_for_status()
            return {'data': pageData.text, 'error': None}
        except requests.exceptions.HTTPError as errh:
            logging.warning(f'HTTP Error on attempt {attempt+1} : {errh}. Failed to fetch data from {url}. ')
            time.sleep(waitTime)
            continue
        except requests.exceptions.ConnectionError as errc:
            logging.warning(f'Error Connecting on attempt {attempt+1}: {errc}. Failed to fetch data from {url}.')
            time.sleep(waitTime)
            continue
        except requests.exceptions.Timeout as errt:
            logging.warning(f'Timeout Error on attempt {attempt+1}: {errt}. Failed to fetch data from {url}.')
            time.sleep(waitTime)
            continue
        except requests.exceptions.RequestException as e: 
            logging.warning(f'Oops on attempt {attempt+1}: Something else {e}. Failed to fetch data from {url}.')
            time.sleep(waitTime)
            continue
    return {'data': None, 'error': f'Max retries reached.  Failed to fetch data from {url}'}
    

#Extract the company name from the html and return
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
        logging.info(f'Attribute error in extractCompanyName: {e}')
        return None
    

#Extract URL using reg ex from JS openWindow method call
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
        #Extracted URL is in the first capture group 
        return match.group(1)
    else:
        #Return None if no match is found
        return None


def extractPDFs(company, doc):
    """
    Extract PDF information from the provided HTML document and add PDF objects to the given company.

    Parameters:
    - company (Company): The Company object to which the extracted PDFs will be added.
    - doc (BeautifulSoup): The BeautifulSoup object representing the HTML document.

    Returns:
    - Company: The Company object with the added PDFs.

    This method searches for the 'planDocuments' section in the HTML document and extracts PDF information
    from the anchor tags within that section. For each anchor tag, it obtains the PDF URL and title, creates
    a PDF object, and adds it to the provided Company.
    """
    planDocsHTML = doc.find(id='planDocuments')
    anchorInstances = planDocsHTML.find_all('a')
    for a in anchorInstances:
        pdfUrlJS = a['href']
        pdfUrl = extractUrlFromExpression(pdfUrlJS)
        pdfTitle = a.find('li').text
        company.add_pdf(pdfUrl, pdfTitle)
    return company


def findValidDoc(homePageUrl, session):
    pageHTML = fetchDataFromURL(homePageUrl, session)
    if pageHTML['error'] is None:
        doc = BeautifulSoup(pageHTML["data"], 'html.parser')
        #Check to see if the account number exists and if not run fetchData again until it does
        accountNumberTable = doc.find('table', {'style': 'background-color:#F8F8F8;border-width: thin;border-collapse:collapse;border-color:#DCDCDC'})
        # Find the cell with the label "Account #:"
        accountLabelCell = accountNumberTable.find('td', text='Account #:')
        if accountLabelCell and accountLabelCell.find_next('td'):
            accountValue = accountLabelCell.find_next('td').text.strip()
            if accountValue:
                logging.info("Valid doc found")
                return doc
            else:
                logging.warning(f"Error: Invalid doc found {homePageUrl} (No Account Value).  Rerunning for correct output...")
                return findValidDoc(homePageUrl, session)
        else: 
            logging.warning(f"Error: Invalid doc found {homePageUrl} (Error is not none) Rerunning for correct output...")
            return findValidDoc(homePageUrl, session) 
    else:
        logging.error(pageHTML['error'])
        return None



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
    logging.info(f"Scraping {homePageUrl}")
    doc = findValidDoc(homePageUrl, session)
    if doc is not None:
        docTitle = doc.head.title.string
        if docTitle == "Fund and Fee Information":
            companyName = extractCompanyName(doc)
            company = Company(companyName)
            logging.info(f"Scraped {homePageUrl}")
            company = extractPDFs(company, doc)
            if len(company.pdfs) == 0:
                return None, "This page does not contain any Plan Documents"
            else: 
                return company, None
        else:
            logging.info(f"{homePageUrl} does not appear to be a valid TA page")
            return None, "This page does not appear to be a valid TA Page"
    else:
        return None, f"Error fetching"


