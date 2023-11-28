import tkinter
from .backend import handleScraping, handleDownload
import logging 
from ..config.logging_config import configure_logging
from pathlib import Path
from threading import Thread
from tkinter import ttk
import pythoncom



def handlePDFScraping():
    pythoncom.CoInitialize()

    global ScrapeResultText

    topProgressBar.pack()
    topProgressBar['value']= 0

    
    textColor="black"
    inputPath = Path(TAURLFileEntry.get())
    outputPath = inputPath.parent / Path("ScrapedPDFs.xlsx")

    logging.info("Scrape PDFs button clicked")

    ScrapeResultText = "Loading..."
    topResultLabel.config(text = ScrapeResultText, fg = textColor)
    window.update()

    def updateProgress(value):
        topProgressBar['value'] = value
        window.update()

    try:
        handleScraping(Path(TAURLFileEntry.get()), outputPath, updateProgress)
        ScrapeResultText = "Successfully scraped! Check for ScrapedPDFs file in parent directory"
        textColor = "green"
    except PermissionError as pe:
        logging.error(f'Permission Error: {pe}')
        ScrapeResultText = f"Permission error: Make sure to close input ({inputPath}) and output ({outputPath}) TransAmerica URL File"
        textColor = "red"
    except OSError as ose:
        if ose.errno == 22 or ose.errno == 2:
            logging.error(f'Type Error: {ose}')
            ScrapeResultText = f"Invalid argument entered.  Please check file{inputPath}"
        else:
            logging.error(f'ERROR: {ose}')
            ScrapeResultText = f"Error: {str(ose)}"
        textColor = "red"
    except Exception as e:
        logging.error(f'ERROR: {e}')
        ScrapeResultText = f"Error: {str(e)}"
        textColor = "red"

    topResultLabel.config(text = ScrapeResultText, fg=textColor)
    topProgressBar.pack_forget()
    window.update()


def handlePDFDownloading():
    # Ensure CoInitialize is called in the thread
    pythoncom.CoInitialize()

    global DownloadResultText

    bottomProgressBar.pack()
    bottomProgressBar['value']= 0

    textColor="black"
    inputPath = Path(PDFFileEntry.get())
    parentPath = inputPath.parent

    logging.info("Save PDFs button clicked")

    DownloadResultText = "Loading..."
    bottomResultLabel.config(text = DownloadResultText, fg = textColor)
    window.update()

    def updateProgress(value):
        bottomProgressBar['value'] = value
        window.update()

    try:
        handleDownload(Path(PDFFileEntry.get()), updateProgress)
        DownloadResultText = f"Successfully saved! Check for PDFs in {parentPath}"
        textColor = "green"
    except PermissionError as pe:
        logging.error(f'Permission Error: {pe}')
        DownloadResultText = f"Permission error: Make sure to close input ({inputPath})"
        textColor = "red"
    except OSError as ose:
        if ose.errno == 22 or ose.errno == 2:
            logging.error(f'Type Error: {ose}')
            DownloadResultText = f"Invalid argument entered.  Please check file {inputPath}"
        else:
            logging.error(f'ERROR: {ose}')
            DownloadResultText = f"Error: {str(ose)}"
        textColor = "red"
    except Exception as e:
        logging.error(f'ERROR: {e}')
        print(type(e))
        DownloadResultText = f"Error: {str(e)}"
        textColor = "red"


    bottomResultLabel.config(text = DownloadResultText, fg=textColor)
    bottomProgressBar.pack_forget()
    window.update()

def start_thread(func):
    t = Thread(target = func)
    t.start()        

if __name__ == '__main__':
    
    #Configure logging 
    configure_logging(Path("pdf-harvesting-app") / Path("LogFile.log"))
    logging.info("Application Opened")

    window = tkinter.Tk()



    window.geometry("450x500")
    window.title("PDF Harvesting Application")

    top_frame = tkinter.Frame(window).pack()
    bottom_frame = tkinter.Frame(window).pack(side = "bottom")

    #TOP Frame Widgets
    topLabel = tkinter.Label(top_frame, text="Enter TransAmerica URL file location (Ex: C:...xlsx): ")
    topWarning = tkinter.Label(top_frame, text="REMEMBER TO CLOSE THE INPUT AND OUTPUT FILES")
    topProgressBar = ttk.Progressbar(top_frame, orient="horizontal", length = 200, mode = 'determinate')
    TAURLFileEntry = tkinter.Entry(top_frame, width = 50)
    processURLButton = tkinter.Button(top_frame, text = "Scrape PDFs", command=lambda: start_thread(handlePDFScraping))
    topResultLabel = tkinter.Label(top_frame, wraplength = 400)
    
    topLabel.pack(padx = 20, pady = 20)
    TAURLFileEntry.pack()
    topWarning.pack()
    processURLButton.pack(pady=20)
    topResultLabel.pack()

    bottomLabel = tkinter.Label(bottom_frame, text="Enter PDF file location (Ex: C:...xlsx)")
    bottomWarning = tkinter.Label(bottom_frame, text="REMEMBER TO CLOSE THE INPUT AND OUTPUT FILES")
    bottomProgressBar = ttk.Progressbar(top_frame, orient="horizontal", length = 200, mode = 'determinate')
    PDFFileEntry = tkinter.Entry(bottom_frame, width = 50)
    downloadPDFButton = tkinter.Button(bottom_frame, text = "Save PDFs", command=lambda: start_thread(handlePDFDownloading))
    bottomResultLabel = tkinter.Label(bottom_frame, wraplength = 400)

    bottomLabel.pack(padx = 20, pady = 20)
    PDFFileEntry.pack()
    bottomWarning.pack()
    downloadPDFButton.pack(pady=20)
    bottomResultLabel.pack()



    window.mainloop()
