from asyncio import as_completed
from .scraper import scrape_pdf_links
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
from openpyxl import load_workbook


#TODOs 
# Add Asset and Plan Participants scraping 
# Add loading bar
# Write app use instructions




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
    try:
        for xls_sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name = xls_sheet_name)
            dfUrls = df['URL'].values.tolist()
            taUrls.append(dfUrls)
        return taUrls
    except Exception as e:
        logging.exception(f"Error getting TA URLS: {e}")
        raise e
    



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
            # with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:
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
                    df.at[url_index, 'Active'] = str(companyTuple[1])
        except Exception as e:
            logging.error(f"Error scraping: {e}")
            continue
        try:
            df.to_excel(xls, sheet_name = current_sheet, index=False)
            logging.info(f"File for {current_sheet} saved successfully")
        except PermissionError as pe:
            logging.error(f'PermissionError: {pe}')
            logging.warning(f"The file is open in another application")
            raise pe

        
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
        session = requests.Session()
        adapter = requests.adapters.HTTPAdapter(pool_connections = 100, pool_maxsize = 100)
        session.mount('http://', adapter)
        session.mount('https://', adapter)
        with pd.ExcelFile(xlPath) as xls: 
            taUrls = getTaURLs(xls)
            companies = processURLs(taUrls, xls, session)
        return companies
    except PermissionError as pe:
        logging.error(f'PermissionError: {pe}')
        raise pe
    except Exception as e:
        logging.exception(f"Error extracting data from Excel file {xlPath}: {e}")
        raise e

    

def saveCompanyandPDFs(companies, outputPath):
    """
    Save company and PDF data to an Excel file.

    Parameters:
    - companies (list): List of Company objects.
    - outputPath (str): Path to the output Excel file.

    Returns:
    - None: The function saves the data to the specified Excel file.
    """
    
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
            
    #Create DataFrames
    dfPDF = pd.DataFrame(dfPDFData, columns=['Company', 'PDF Title', 'PDF URL', 'Source'])
    dfCompany = pd.DataFrame(dfCompanyData, columns=['Company', 'Assets', 'Plan Participants'])

    #Write DataFrames to Excel file
    try:
        with pd.ExcelWriter(outputPath, engine='openpyxl') as writer:
            dfPDF.to_excel(writer, sheet_name= 'PDFs', index=False)
            dfCompany.to_excel(writer, sheet_name='Companies', index = False)
    except Exception as e:
        logging.error(e)
        raise e



def generateXLSheet(inputPath):
    """
    Generate an Excel sheet with company and PDF data.

    Parameters:
    - inputPath (Path): Path to the input Excel file containing company data.

    Returns:
    - None: The function creates an Excel sheet with company and PDF data.
    """
    try:
        companies = extractTAExcel(inputPath)
        outputPath = inputPath.parent / Path("ScrapedPDFs.xlsx")
        saveCompanyandPDFs(companies, outputPath)
    except Exception as e:
        logging.exception(f"Error generating Excel sheet: {e}")
        raise e





def downloadPDF(pdfURL,filePath, maxRetries = 3, retryDelay = 1):
    """
    Download a PDF from the given URL and save it to the specified file path.

    Parameters:
    - pdfURL (str): The URL of the PDF to download.
    - filePath (str): The local file path to save the downloaded PDF.
    - maxRetries (int, optional): Maximum number of download retries in case of failure. Default is 3.
    - retryDelay (int, optional): Delay (in seconds) between download retries. Default is 1.

    Returns:
    - bool: True if the PDF is successfully downloaded, False otherwise.
    """
    retries = 0
    while retries < maxRetries:
        try:
            with urllib.request.urlopen(pdfURL) as response, open(filePath, 'wb') as out_file:
                data = response.read()
                out_file.write(data)
            return True
        except (URLError, HTTPError, RemoteDisconnected) as e:
            logging.error(f"Download PDF Error {retries}: {e}")
            retries += 1
            time.sleep(retryDelay)
    return False

def addCompanyLinks(inputPath):
    """
    Add hyperlinks to local directories for each company in the Companies sheet of an Excel file.

    Parameters:
    - inputPath (Path): Path to the input Excel file.

    Returns:
    - None: The function adds hyperlinks to the Companies sheet and saves the modified Excel file.
    """

    with pd.ExcelFile(inputPath, engine='openpyxl') as xls:
        dfCompany = pd.read_excel(xls, sheet_name='Companies')

    companyHotlinks= {}
    lock = threading.Lock()

    def downloadAndSave(company):
        fileDirectory=""
        if pd.notna(company): 
            fileDirectory = inputPath.parent / Path(company)
        with lock:
            companyHotlinks[company]= fileDirectory

    
    # with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:
    with concurrent.futures.ThreadPoolExecutor() as executor:
        executor.map(downloadAndSave, dfCompany['Company'])
    
    dfCompany['Hotlink'] = dfCompany['Company'].map(companyHotlinks)
    dfCompany['Hotlink'] = dfCompany['Hotlink'].apply(lambda x: f'=HYPERLINK("{x}", "CLICK FOR FOLDER")')

    with pd.ExcelWriter(inputPath, if_sheet_exists='replace', mode='a') as writer:
        #Save FilePath to row
        dfCompany.to_excel(writer, sheet_name = "Companies", index = False)


def refreshExcel(inputPath):
    """
    Refresh all data connections and calculations in an Excel workbook.

    Parameters:
    - inputPath (Path): Path to the Excel file to be refreshed.

    Returns:
    - None: The function refreshes the workbook and saves the changes.
    """
    File = win32com.client.Dispatch("Excel.Application")    
    File.Visible = 1
    Workbook = File.Workbooks.open(str(inputPath))
    Workbook.RefreshAll()
    Workbook.Save()
    File.Quit()


def extractPDFPages(inputPath):
    """
    Extract PDF data from an Excel file, download PDFs, and update the Excel file.

    Parameters:
    - inputPath (Path): Path to the input Excel file containing PDF data.

    Returns:
    - None: The function downloads PDFs, updates the Excel file, and adds hyperlinks.
    """    
    refreshExcel(inputPath)

    
    #Save PDF pages on excel document to local directory
    with pd.ExcelFile(inputPath, engine='openpyxl') as xls:
        dfPDF = pd.read_excel(xls, sheet_name = 'PDFs')

    
    localFilePaths = {}
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
            
            try:
                #Download the pdf and save to the local file path
                success = downloadPDF(pdfURL, filePath)
                if success:
                    logging.info(f"{pdfTitle} saved successfully")
                else:
                    logging.warning(f"Failed to download {pdfTitle}")
            except Exception as e:
                logging.error(f"Error downloading {pdfTitle}: {e}")
                raise e
        with lock:
                localFilePaths[pdfURL] = filePath

    
    # with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:
    with concurrent.futures.ThreadPoolExecutor() as executor:
        executor.map(downloadAndSave, zip(dfPDF['PDF URL'], dfPDF['PDF Title'], dfPDF['Company']))

    # Update 'Local FilePath' column based on the corresponding pdfURL
    dfPDF['Local FilePath'] = dfPDF['PDF URL'].map(localFilePaths)
    dfPDF['Local FilePath'] = dfPDF['Local FilePath'].apply(lambda x: f'=HYPERLINK("{x}", "CLICK FOR FILE")')

    try:
        with pd.ExcelWriter(inputPath, if_sheet_exists='replace', mode='a') as writer:
            dfPDF['Company'] = dfPDF['Company'].apply(lambda x: f"=IFERROR(VLOOKUP(\"{x}\",Companies!A:A, 1, FALSE), \"null\")")
            dfPDF.to_excel(writer, sheet_name = "PDFs", index = False)
    except Exception as e:
        logging.error(f"Error saving: {e}")
        raise e

    addCompanyLinks(inputPath)
    return None

def is_excel_file_open(file_path):
    try:
        # Try to open the file with pandas
        with pd.ExcelFile(file_path):
            pass
        # If successful, the file is not open
        return False
    except PermissionError:
        # If PermissionError occurs, the file is open
        return True
    except OSError:
        return False
    



def handleScraping(inputPath, outputPath):
    try:
        if (not(is_excel_file_open(inputPath) or is_excel_file_open(outputPath))):
            generateXLSheet(inputPath)
        else:
            raise PermissionError
    except Exception as e:
        raise e

def handleDownload(inputPath):
    try:
        if (not(is_excel_file_open(inputPath))):
            extractPDFPages(inputPath)
        else: 
            raise PermissionError
    except Exception as e:
        logging.error(f"Download error: {e}")
        raise e


