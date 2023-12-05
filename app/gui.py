import tkinter
from app.backend import handleScraping, handleDownload
import logging 
from pathlib import Path
from threading import Thread
from tkinter import ttk
import pythoncom




class PDFHarvestingApp:
    def __init__(self, window):
        self.window = window
        self.window.geometry("450x500")
        self.window.title("PDF Harvesting Application")
        
        self.create_gui_elements()

    def create_gui_elements(self):
        top_frame = tkinter.Frame(self.window).pack()
        bottom_frame = tkinter.Frame(self.window).pack(side = "bottom")

        
        #TOP Frame Widgets
        self.topLabel = tkinter.Label(top_frame, text="Enter TransAmerica URL file location (Ex: C:...xlsx): ")
        self.topWarning = tkinter.Label(top_frame, text="REMEMBER TO CLOSE THE INPUT AND OUTPUT FILES")
        self.topProgressBar = ttk.Progressbar(top_frame, orient="horizontal", length = 0, mode = 'determinate')
        self.TAURLFileEntry = tkinter.Entry(top_frame, width = 50)
        self.processURLButton = tkinter.Button(top_frame, text = "Scrape PDFs", command=lambda: self.start_thread(self.handlePDFScraping))
        self.topResultLabel = tkinter.Label(top_frame, wraplength = 400)

        self.topLabel.pack(padx = 20, pady = 20)
        self.TAURLFileEntry.pack()
        self.topWarning.pack()
        self.processURLButton.pack(pady=20)
        self.topProgressBar.pack()
        self.topResultLabel.pack()

        self.bottomLabel = tkinter.Label(bottom_frame, text="Enter PDF file location (Ex: C:...xlsx)")
        self.bottomWarning = tkinter.Label(bottom_frame, text="REMEMBER TO CLOSE THE INPUT AND OUTPUT FILES")
        self.bottomProgressBar = ttk.Progressbar(top_frame, orient="horizontal", length = 0, mode = 'determinate')
        self.PDFFileEntry = tkinter.Entry(bottom_frame, width = 50)
        self.downloadPDFButton = tkinter.Button(bottom_frame, text = "Save PDFs", command=lambda: self.start_thread(self.handlePDFDownloading))
        self.bottomResultLabel = tkinter.Label(bottom_frame, wraplength = 400)

        self.bottomLabel.pack(padx = 20, pady = 20)
        self.PDFFileEntry.pack()
        self.bottomWarning.pack()
        self.downloadPDFButton.pack(pady=20)
        self.bottomProgressBar.pack()
        self.bottomResultLabel.pack()


    def run(self):
        self.window.mainloop()

        


    def handlePDFScraping(self):
        pythoncom.CoInitialize()

        global ScrapeResultText

        self.topProgressBar.configure(length=200)
        self.topProgressBar['value']= 0

        
        textColor="black"
        inputPath = Path(self.TAURLFileEntry.get())
        outputPath = inputPath.parent / Path("ScrapedPDFs.xlsx")

        logging.info("Scrape PDFs button clicked")

        ScrapeResultText = "Loading..."
        self.topResultLabel.config(text = ScrapeResultText, fg = textColor)
        self.window.update()

        def updateProgress(value):
            ScrapeResultText = f"Loading...{value}%"
            self.topResultLabel.config(text = ScrapeResultText, fg = "black")
            self.topProgressBar['value'] = value
            self.window.update()

        try:
            handleScraping(Path(self.TAURLFileEntry.get()), updateProgress)
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

        self.topResultLabel.config(text = ScrapeResultText, fg=textColor)
        self.topProgressBar.config(length=0)
        self.window.update()


    def handlePDFDownloading(self):
        # Ensure CoInitialize is called in the thread
        pythoncom.CoInitialize()

        global DownloadResultText

        self.bottomProgressBar.config(length=200)
        self.bottomProgressBar['value']= 0

        textColor="black"
        inputPath = Path(self.PDFFileEntry.get())
        parentPath = inputPath.parent

        logging.info("Save PDFs button clicked")

        DownloadResultText = "Loading..."
        self.bottomResultLabel.config(text = DownloadResultText, fg = textColor)
        self.window.update()

        def updateProgress(value):
            DownloadResultText = f"Loading...{value}%"
            self.bottomResultLabel.config(text = DownloadResultText, fg = textColor)
            self.bottomProgressBar['value'] = value
            self.window.update()

        try:
            handleDownload(Path(self.PDFFileEntry.get()), updateProgress)
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
            DownloadResultText = f"Error: {str(e)}"
            textColor = "red"


        self.bottomResultLabel.config(text = DownloadResultText, fg=textColor)
        self.bottomProgressBar.config(length = 0)
        self.window.update()

    def start_thread(self, func):
        t = Thread(target = func)
        t.start() 







