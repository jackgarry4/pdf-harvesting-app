from .scraper import scrape_pdf_links
import pandas as pd
import openpyxl
import requests
import logging


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
def process_urls(taUrls, xls):
    companies = []
    #Urls stored in list of lists with index of list corresponding to sheet index
    for ind, sheetList in enumerate(taUrls):
        current_sheet = xls.sheet_names[ind]
        df = pd.read_excel(xls, sheet_name = current_sheet)
        for url in sheetList:
            logging.info(f"Scraping {url}")
            companyTuple = scrape_pdf_links(url)
            logging.info(f"Scraped {url}")

            url_index = df.index[df['URL'] == url][0]

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
            print(f"PermissionError: The file is open in another application")
    return companies

def extract_excel(xlPath):
    with pd.ExcelFile(xlPath) as xls: 
        taUrls = get_TA_urls(xls)
        companies = process_urls(taUrls, xls)
        return companies


def main():
    xlPath = 'C:/Users/Computer/OneDrive - The Ohio State University/Documents/Mosby Project/pdf-harvesting-app/docs/Formatted TA URLs.xlsx'
    return extract_excel(xlPath)

if __name__ == "__main__":
    main()