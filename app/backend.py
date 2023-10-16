from scraper import scrape_pdf_links

def process_urls(urls):
    pdf_links = scrape_pdf_links(urls)

    return pdf_links