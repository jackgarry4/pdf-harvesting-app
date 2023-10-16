from cgitb import text
from bs4 import BeautifulSoup
from ..classes.Company import Company
from ..classes.PDF import PDF
import requests 
import re
import webbrowser




#Make HTTP request and return HTML response as String
def fetchDataFromURL(url):
    #Make HTTP request to provided url
    try:
        pageData = requests.get(url)
        pageData.raise_for_status()
    except requests.exceptions.HTTPError as errh:
        print ("HTTP Error:", errh)
    except requests.exceptions.ConnectionError as errc:
        print("Error Connecting:", errc)
    except requests.exceptions.Timeout as errt:
        print("Timeout Error:",errt)
    except requests.exceptions.RequestException as e: 
        print("Oops: Something else", e)
    return pageData.text

#Extract the company name from the html and return
def extractCompanyName(doc):
    #MAY HAVE TO ADD TRY/EXCEPT IN CASES WHERE THE DOCUMENTS DO NOT EXIST 
    companyTitle = doc.h2.b.string
    return companyTitle

#Extract URL using reg ex from JS openWindow method call
def extractUrlFromExpression(pdfUrlJS):
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
    planDocsHTML = doc.find(id='planDocuments')
    anchorInstances = planDocsHTML.find_all('a')
    #FOREACH to run through each a instance and build PDF object and append to companies.pdf
    for a in anchorInstances:
        pdfUrlJS = a['href'] #javascript that will open the url.  Need to pass through RegEx to obtain only URL
        pdfUrl = extractUrlFromExpression(pdfUrlJS)
        print(pdfUrl)
        pdfTitle = a.li.string
        pdf = PDF("", pdfTitle)
    return company




#Create Company object from the provided TransAmerica company specific html script
def scrapeCompanyPage(html):
    doc = BeautifulSoup(html, 'html.parser')
    company = Company(extractCompanyName(doc))
    extractPDFs(company, doc)
    
    return company



#From provided url fetch the 
def scrape_pdf_links(homePageUrl):
    pageHTML = fetchDataFromURL(homePageUrl)
    company = scrapeCompanyPage(pageHTML)
    return company



def main():
    url = "https://fp.trsretire.com/PublicFP/fpClient.jsp?c=TT069026&a=00001&l=TDA&p=mhs"
    return scrape_pdf_links(url)

    

if __name__ == "__main__":
    main()