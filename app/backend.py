from .scraper import scrape_pdf_links
import pandas as pd
import openpyxl




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
            company = scrape_pdf_links(url)
            #Find company by url in xlsPath and write status to excel file
            url_index = df.index[df['URL'] == url][0]
            df.at[url_index, 'Active'] = company.active
            if (company.active):
                companies.append(company)
        df.to_excel(xls, sheet_name = current_sheet, index=False)
    return companies


def main():
    xlPath = 'C:/Users/Computer/OneDrive - The Ohio State University/Documents/Mosby Project/pdf-harvesting-app/docs/Formatted TA URLs.xlsx'
    with pd.ExcelFile(xlPath) as xls: 
        taUrls = get_TA_urls(xls)
        companies = process_urls(taUrls, xls)
        return companies

if __name__ == "__main__":
    main()