from asyncio import as_completed
from .scraper import scrape_pdf_links
from ..config.logging_config import configure_logging
from urllib.error import URLError, HTTPError
from http.client import RemoteDisconnected
import pandas as pd
import requests
import logging
import concurrent.futures
import time
import urllib.request
import os
from pathlib import Path
import time
import threading
import win32com.client


#PROBLEMS 
#1. Add Hotlinks to each companies folder path
#2. Add comments for methods

#QUESTIONS FOR DYLAN
#1. Do you want to add company name by the URLs on Formatted URL.xlsx page?
#2. 


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
            with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
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
                    df.at[url_index, 'Active'] = str(companyTuple[1])
        except Exception as e:
            logging.error(f"Error scraping: {e}")
            continue
        try:
            df.to_excel(xls, sheet_name = current_sheet, index=False)
            logging.info(f"File for {current_sheet} saved successfully")
        except PermissionError as e:
            logging.error(f'PermissionError: {e}')
            logging.warning(f"The file is open in another application")
        
    logging.info("--- %s seconds ---" % (time.time() - start_time))
    return companies


def extractTAExcel(xlPath):
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
        with  requests.Session() as session, pd.ExcelFile(xlPath) as xls: 
            taUrls = getTaURLs(xls)
            companies = processURLs(taUrls, xls, session)
        return companies
    except Exception as e:
        logging.exception(f"Error extracting data from Excel file {xlPath}: {e}")
        return None

    

def saveCompanyandPDFs(companies, outputPath):
    dfPDFData = []
    dfCompanyData = []

    for company in companies:
        dfCompanyData.append({'Company' : company.name, 'Assets' : 0, 'Plan Participants' : 0})
        for pdf in company.pdfs:
            dfPDFData.append({
                'Company' : f"=IFERROR(VLOOKUP(\"{company.name}\",Companies!A:A, 1, FALSE), \"null\")", 
                'PDF Title': pdf.title, 
                'PDF URL': pdf.url, 
                'Source' : pdf.source})
    # Create DataFrames
    dfPDF = pd.DataFrame(dfPDFData, columns=['Company', 'PDF Title', 'PDF URL', 'Source'])
    dfCompany = pd.DataFrame(dfCompanyData, columns=['Company', 'Assets', 'Plan Participants'])

    #Write DataFrames to Excel file
    with pd.ExcelWriter(outputPath, engine='openpyxl') as writer:
        dfPDF.to_excel(writer, sheet_name= 'PDFs', index=False)
        dfCompany.to_excel(writer, sheet_name='Companies', index = False)
    return None



def generateXLSheet(inputPath):
    companies = extractTAExcel(inputPath)
    outputPath = inputPath.parent / Path("ScrapedPDFs.xlsx")
    saveCompanyandPDFs(companies, outputPath)




def downloadPDF(pdfURL,filePath, maxRetries = 3, retryDelay = 1):
    retries = 0
    while retries < maxRetries:
        try:
            with urllib.request.urlopen(pdfURL) as response, open(filePath, 'wb') as out_file:
                data = response.read()
                out_file.write(data)
            return True
        except (URLError, HTTPError, RemoteDisconnected) as e:
            logging.error(e)
            retries += 1
            time.sleep(retryDelay)
    return False

def addCompanyLinks(inputPath):
    #Save PDF pages on excel document to local directory
    with pd.ExcelFile(inputPath, engine='openpyxl') as xls:
        dfCompany = pd.read_excel(xls, sheet_name='Companies')

    companyHotlinks= []
    lock = threading.Lock()

    def downloadAndSave(company):
        fileDirectory = inputPath.parent / Path(company)
        with lock:
            companyHotlinks.append(fileDirectory)

    with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
        executor.map(downloadAndSave, dfCompany['Company'])

    with pd.ExcelWriter(inputPath, if_sheet_exists='replace', mode='a') as writer:
        #Save FilePath to row
        dfCompany['Hotlink'] = companyHotlinks
        dfCompany['Hotlink'] = dfCompany['Hotlink'].apply(lambda x: f'=HYPERLINK("{x}", "CLICK FOR FILE")')
        dfCompany.to_excel(writer, sheet_name = "Companies", index = False)


def refreshExcel(inputPath):
    # Opening Excel software using the win32com 
    File = win32com.client.Dispatch("Excel.Application")    

    File.Visible = 1

    Workbook = File.Workbooks.open(str(inputPath))

    Workbook.RefreshAll()

    Workbook.Save()

    File.Quit()


def extractPDFPages(inputPath):
    
    refreshExcel(inputPath)

    #Save PDF pages on excel document to local directory
    with pd.ExcelFile(inputPath, engine='openpyxl') as xls:
        dfPDF = pd.read_excel(xls, sheet_name = 'PDFs')

    localFilePaths = []
    lock = threading.Lock()

    def downloadAndSave(pdfData):
        pdfURL, pdfTitle, company = pdfData
        filePath = "null"

        if pd.notna(company):    
            fileDirectory = inputPath.parent / Path(company)
            #Create the local company directory if it does not exist
            if not os.path.exists(fileDirectory):
                os.makedirs(fileDirectory)
            
        

            #Replace instances of / as will mess up file path
            pdfTitle = pdfTitle.replace("/","-").replace("\n", "")
            filePath = fileDirectory / Path(f"{pdfTitle}.pdf")

            

            #Download the pdf and save to the local file path
            success = downloadPDF(pdfURL, filePath)
            if success:
                logging.info(f"{pdfTitle} saved successfully")
            else:
                logging.warning(f"Failed to download {pdfTitle}")
        with lock:
                localFilePaths.append(filePath)

    
    with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
        executor.map(downloadAndSave, zip(dfPDF['PDF URL'], dfPDF['PDF Title'], dfPDF['Company']))

    with pd.ExcelWriter(inputPath, if_sheet_exists='replace', mode='a') as writer:
        #Save FilePath to row
        dfPDF['Local FilePath'] = localFilePaths
        dfPDF['Local FilePath'] = dfPDF['Local FilePath'].apply(lambda x: f'=HYPERLINK("{x}", "CLICK FOR FILE")')
        dfPDF['Company'] = dfPDF['Company'].apply(lambda x: f"=IFERROR(VLOOKUP(\"{x}\",Companies!A:A, 1, FALSE), \"null\")")
        dfPDF.to_excel(writer, sheet_name = "PDFs", index = False)

    addCompanyLinks(inputPath)
    return None

def main():

    #Configure logging 
    configure_logging(Path("pdf-harvesting-app") / Path("LogFile.log"))
    
    start = time.time()
    # inputPath = Path('C:/Users/Computer/OneDrive - The Ohio State University/Documents/Mosby Project/pdf-harvesting-app/docs/Formatted TA URLs.xlsx')
    # generateXLSheet(inputPath)


    inputPath = Path('C:/Users/Computer/OneDrive - The Ohio State University/Documents/Mosby Project/pdf-harvesting-app/docs/ScrapedPDFs.xlsx')
    extractPDFPages(inputPath)
    end = time.time()
    print(f"Time: {end-start}")

if __name__ == "__main__":
    main()