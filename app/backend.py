from asyncio import as_completed
from .scraper import scrape_pdf_links
import pandas as pd
import openpyxl
import requests
import logging
import concurrent.futures
import time


#Takes excel file path containing TA urls as input and returns list of lists with each list containing urls for that indexed excel sheet
def get_TA_urls(xls):
    taUrls = []
    #Iterate through each sheet in the excel file
    for xls_sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name = xls_sheet_name)
        dfUrls = df['URL'].values.tolist()
        taUrls.append(dfUrls)
    return taUrls

#Takes in TransAmerica URL and outputs company with all pdfs saved as attribute
#WITH CONCURRENCY CONTROL AND THREADING
def process_urls(taUrls, xls, session):
    start_time = time.time()
    companies = []
    #Urls stored in list of lists with index of list corresponding to sheet index
    for ind, sheetList in enumerate(taUrls):
        current_sheet = xls.sheet_names[ind]
        df = pd.read_excel(xls, sheet_name = current_sheet)

        
        try:
            with concurrent.futures.ThreadPoolExecutor() as executor:
                futureToURL = {executor.submit(scrape_pdf_links, url, session): url for url in sheetList}
                
        except Exception as e:
            logging.error(f"Error scraping")
            continue
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
        try:
            df.to_excel(xls, sheet_name = current_sheet, index=False)
            logging.info(f"File for {current_sheet} saved successfully")
        except PermissionError as e:
            logging.error(f'PermissionError: {e}')
            logging.warning(f"The file is open in another application")
    print("--- %s seconds ---" % (time.time() - start_time))
    return companies


# #Takes in TransAmerica URL and outputs company with all pdfs saved as attribute
# #WITHOUT CONCURRENCY CONTROL AND WITHOUT THREADING 
# def process_urls(taUrls, xls, session):
#     start_time = time.time()
#     companies = []
#     #Urls stored in list of lists with index of list corresponding to sheet index
#     for ind, sheetList in enumerate(taUrls):
#         current_sheet = xls.sheet_names[ind]
#         df = pd.read_excel(xls, sheet_name = current_sheet)

#         for url in sheetList:
#             logging.info(f"Scraping {url}")
#             print(f"Scraping {url}")
#             try:
#                 companyTuple = scrape_pdf_links(url, session)
#                 logging.info(f"Scraped {url}")
#             except Exception as e:
#                 logging.error(f"Error scraping {url}: {e}")
#                 continue
#             print(f"Scraped {url}")

#             url_index = df.index[df['URL'] == url][0]

#             if companyTuple[0] is not None:
#                 #Find company by url in xlsPath and write status to excel file
#                 df.at[url_index, 'Active'] = "True"
#                 companies.append(companyTuple[0])
#             else:
#                 df.at[url_index, 'Active'] = companyTuple[1]
#         try:
#             df.to_excel(xls, sheet_name = current_sheet, index=False)
#             logging.info(f"File for {current_sheet} saved successfully")
#         except PermissionError as e:
#             logging.error(f'PermissionError: {e}')
#             logging.warning(f"The file is open in another application")
#     print("--- %s seconds ---" % (time.time() - start_time))
#     return companies

def extract_excel(xlPath):
    session = requests.Session()
    with pd.ExcelFile(xlPath) as xls: 
        taUrls = get_TA_urls(xls)
        companies = process_urls(taUrls, xls, session)
    session.close()
    return companies


def main():
    xlPath = 'C:/Users/Computer/OneDrive - The Ohio State University/Documents/Mosby Project/pdf-harvesting-app/docs/Formatted TA URLs.xlsx'
    return extract_excel(xlPath)

if __name__ == "__main__":
    main()