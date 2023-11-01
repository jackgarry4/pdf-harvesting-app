from asyncio import as_completed
from .scraper import scrape_pdf_links
import pandas as pd
import openpyxl
import requests
import logging
import concurrent.futures
import time
import re


def getTaURLs(xls):
    """
    Takes an Excel file path containing TA URLs as input and returns the urls contained in the Excel as a list of lists.
    
    Each inner list corresponds to the URLs from a specific sheet in the Excel file.

    Parameters:
    - xls (str): Path to the Excel file.

    Returns:
    List[List[str]]: A list of lists, where each inner list contains URLs for the corresponding sheet.
    """
    taUrls = []
    for xls_sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name = xls_sheet_name)
        dfUrls = df['URL'].values.tolist()
        taUrls.append(dfUrls)
    return taUrls



def processURLs(taUrls, xls, session):
    """
    Process TransAmerica URLs concurrently, scrape PDF links, and update an Excel file.

    This function takes a list of lists containing TransAmerica URLs, an Excel file,
    and a session object for making HTTP requests. It utilizes a ThreadPoolExecutor
    for concurrent processing of URLs, scrapes PDF links, updates the corresponding
    Excel file with the status of each URL, and returns a list of Company objects.

    Parameters:
    - taUrls (List[List[str]]): List of lists, where each inner list contains URLs for a sheet.
    - xls (str): Path to the Excel file.
    - session: Session object for making HTTP requests.

    Returns:
    List[Company]: A list of Company objects representing the processed companies.
    """
    start_time = time.time()
    companies = []

    for ind, sheetList in enumerate(taUrls):
        current_sheet = xls.sheet_names[ind]
        df = pd.read_excel(xls, sheet_name = current_sheet)

        try:
            with concurrent.futures.ThreadPoolExecutor() as executor:
                futureToURL = {executor.submit(scrape_pdf_links, url, session): url for url in sheetList}

            for future in concurrent.futures.as_completed(futureToURL):
                url = futureToURL[future]
                url_index = df.index[df['URL'] == url][0]
                companyTuple = future.result()


                if companyTuple[0] is not None:
                    #Find company by url in xlsPath and write status to excel file
                    df.at[url_index, 'Active'] = "True"
                    companies.append(companyTuple[0])
                else:
                    df.at[url_index, 'Active'] = companyTuple[1]
        except Exception as e:
            logging.error(f"Error scraping")
            continue
        try:
            df.to_excel(xls, sheet_name = current_sheet, index=False)
            logging.info(f"File for {current_sheet} saved successfully")
        except PermissionError as e:
            logging.error(f'PermissionError: {e}')
            logging.warning(f"The file is open in another application")
        
    logging.info("--- %s seconds ---" % (time.time() - start_time))
    return companies


def extractExcel(xlPath):
    """
    Extract data from a TransAmerica Excel file.

    This function takes the path to a TransAmerica Excel file as input, processes the contained
    URLs concurrently, scrapes PDF links, and returns a list of Company objects representing
    the processed companies. The function uses a requests.Session for making HTTP requests and
    ensures proper resource management by using context managers for both the session and Excel file.

    Parameters:
    - xlPath (str): Path to the TransAmerica Excel file.

    Returns:
    List[Company] or None: A list of Company objects if extraction is successful, or None in case of an error.
    """
    try:
        with requests.Session() as session, pd.ExcelFile(xlPath) as xls: 
            taUrls = getTaURLs(xls)
            companies = processURLs(taUrls, xls, session)
        return companies
    except Exception as e:
        logging.exception(f"Error extracting data from Excel file {xlPath}: {e}")
        return None

    
def generateOutputXlPath(inputPath):
    outputPath = re.match(r"^(.*?)/[^/]+$", inputPath).group(1)
    return outputPath+"/ScrapedPDFs.xlsx"

def saveCompanyPDFs(companies, outputPath):
    df = pd.DataFrame(columns=['Company', 'PDF Title', 'PDF URL', 'Assets', 'Number of Participants', 'Source'])
    for company in companies:
        for pdf in company.pdfs:
            d = {'Company' : company.name, 'PDF Title': pdf.title, 'PDF URL': pdf.url, 'Assets' : 0, 'Number of Participants': 0, 'Source' : pdf.source}
            dfVal = pd.DataFrame(data = [d], columns=['Company', 'PDF Title', 'PDF URL', 'Assets', 'Number of Participants', 'Source'])
            df= pd.concat([df, dfVal], ignore_index=True)
            df.reset_index(drop=True, inplace = True)
    
    df.to_excel(outputPath)
    return None


def generatePDFPage(inputPath):
    inputPath = 'C:/Users/Computer/OneDrive - The Ohio State University/Documents/Mosby Project/pdf-harvesting-app/docs/Formatted TA URLs.xlsx'
    companies = extractExcel(inputPath)
    outputPath = generateOutputXlPath(inputPath)
    saveCompanyPDFs(companies, outputPath)

def savePDFPages(inputPath):
    #Save PDF pages on excel document to local directory
    return None

def main():
    #inputPath = 'C:/Users/Computer/OneDrive - The Ohio State University/Documents/Mosby Project/pdf-harvesting-app/docs/Formatted TA URLs.xlsx'
    #generatePDFPage(inputPath)
    inputPath = 'C:\Users\Computer\OneDrive - The Ohio State University\Documents\Mosby Project\pdf-harvesting-app\docs\ScrapedPDFs.xlsx'
    savePDFPages(inputPath)

if __name__ == "__main__":
    main()