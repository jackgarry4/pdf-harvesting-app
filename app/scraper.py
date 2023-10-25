from cgitb import text
from uu import Error
from xml.etree.ElementTree import TreeBuilder
from bs4 import BeautifulSoup
from ..classes.Company import Company
from ..classes.PDF import PDF
import requests 
import re
import logging


#Create a Session object

def fetchDataFromURL(url, session):
    """
    Make an HTTP request to the provided URL and return the HTML response as a string.

    Parameters:
    - url (str): The URL to make the HTTP request.

    Returns:
    - dict: A dictionary with 'data' containing the HTML response text (if successful),
            and 'error' containing the error message (if an error occurs).
    """
    try:
        pageData = session.get(url)
        pageData.raise_for_status()
        return {'data': pageData.text, 'error': None}
    except requests.exceptions.HTTPError as errh:
        logging.warning(f'HTTP Error: {errh}')
        return {'data': None, 'error': f'HTTP Error: {errh}'}
    except requests.exceptions.ConnectionError as errc:
        logging.warning(f'Error Connecting: {errc}')
        return {'data': None, 'error': f'Error connecting: {errc}'}
    except requests.exceptions.Timeout as errt:
        logging.warning(f'Timeout Error: {errt}')
        return {'data': None, 'error': f'Timeout Error: {errt}'}
    except requests.exceptions.RequestException as e: 
        logging.warning(f'Oops: Something else {e}')
        return {'data': None, 'error': f'Oops error: {e}'}
    

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
        logging.warning(f'Attribute error in extractCompanyName: {e}')
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
        pdf = PDF(pdfUrl, pdfTitle)
        company.add_pdf(pdf)
    return company



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
    print(f"Scraping {homePageUrl}")
    logging.info(f"Scraping {homePageUrl}")
    pageHTML = fetchDataFromURL(homePageUrl, session)
    if pageHTML["error"] is None:
        doc = BeautifulSoup(pageHTML["data"], 'html.parser')
        docTitle = doc.head.title.string
        companyName = extractCompanyName(doc)
        if docTitle == "Fund and Fee Information" and companyName is not None:
            company = Company(companyName)
            logging.info(f"Scraped {homePageUrl}")
            return extractPDFs(company, doc), None
        else:
            logging.info(f"Scraped {homePageUrl}")
            return None, "This page does not appear to be a valid TA Page"
    else:
        logging.info(f"Scraped {homePageUrl}")
        return None, "Error fetching HTML: {pageHTML['error']}"


