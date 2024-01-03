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

        self.window.protocol("WM_DELETE_WINDOW", self.on_close)

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

    def on_close(self):
        logging.info("Exiting Program")
        self.window.destroy()


    def handlePDFScraping(self):
        pythoncom.CoInitialize()

        
        inputPath = Path(self.TAURLFileEntry.get())
        outputPath = inputPath.parent / Path("ScrapedPDFs.xlsx")

        logging.info("Scrape PDFs button clicked")

        self.updateTopProgress("Loading...", 0)

        try:
            handleScraping(Path(self.TAURLFileEntry.get()), self.updateTopProgress)
            resultText = "Successfully scraped! Check for ScrapedPDFs file in parent directory"
            textColor = "green"
        except PermissionError as pe:
            logging.error(f'Permission Error: {pe}')
            resultText = f"Permission error: Make sure to close input ({inputPath}) and output ({outputPath}) TransAmerica URL File"
            textColor = "red"
        except OSError as ose:
            if ose.errno == 22 or ose.errno == 2:
                logging.error(f'Type Error: {ose}')
                resultText = f"Invalid argument entered.  Please check file{inputPath}"
            else:
                logging.error(f'ERROR: {ose}')
                resultText = f"Error: {str(ose)}"
            textColor = "red"
        except Exception as e:
            logging.error(f'ERROR: {e}')
            resultText = f"Error: {str(e)}"
            textColor = "red"

        self.updateTopProgress(resultText = resultText, value = 0, textColor = textColor)


    def handlePDFDownloading(self):
        # Ensure CoInitialize is called in the thread
        pythoncom.CoInitialize()

        inputPath = Path(self.PDFFileEntry.get())
        parentPath = inputPath.parent

        logging.info("Save PDFs button clicked")

        self.updateBottomProgress("Loading...", 0)

        try:
            handleDownload(Path(self.PDFFileEntry.get()), self.updateBottomProgress)
            resultText = f"Successfully saved! Check for PDFs in {parentPath}"
            textColor = "green"
        except PermissionError as pe:
            logging.error(f'Permission Error: {pe}')
            resultText = f"Permission error: Make sure to close input ({inputPath})"
            textColor = "red"
        except OSError as ose:
            if ose.errno == 22 or ose.errno == 2:
                logging.error(f'Type Error: {ose}')
                resultText = f"Invalid argument entered.  Please check file {inputPath}"
            else:
                logging.error(f'ERROR: {ose}')
                resultText = f"Error: {str(ose)}"
            textColor = "red"
        except Exception as e:
            logging.error(f'ERROR: {e}')
            resultText = f"Error: {str(e)}"
            textColor = "red"


        self.updateBottomProgress(resultText = resultText, value = 0, textColor = textColor)


    def updateBottomProgress(self, resultText, value, textColor = "black"):
        if (value == 0):
            self.bottomProgressBar.config(length = 0)
        else:
            self.bottomProgressBar.config(length = 200)
        self.bottomResultLabel.config(text = resultText, fg = textColor)
        self.bottomProgressBar['value'] = value
        self.window.update()

    def updateTopProgress(self, resultText, value, textColor = "black"):
        if (value == 0):
            self.topProgressBar.config(length = 0)
        else:
            self.topProgressBar.config(length = 200)
        self.topResultLabel.config(text = resultText, fg = textColor)
        self.topProgressBar['value'] = value
        self.window.update()

    def start_thread(self, func):
        t = Thread(target = func)
        t.start() 







