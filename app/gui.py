import tkinter
from .backend import handleScraping, handleDownload
import logging 
from ..config.logging_config import configure_logging
from pathlib import Path





if __name__ == '__main__':
    
    #Configure logging 
    configure_logging(Path("pdf-harvesting-app") / Path("LogFile.log"))
    logging.info("Application Opened")

    window = tkinter.Tk()



    def handlePDFScraping():
        global ScrapeResultText

        textColor="black"
        inputPath = Path(TAURLFileEntry.get())
        outputPath = inputPath.parent / Path("ScrapedPDFs.xlsx")

        logging.info("Scrape PDFs button clicked")

        ScrapeResultText = "Loading..."
        topResultLabel.config(text = ScrapeResultText, fg = textColor)
        window.update()


        try:
            handleScraping(Path(TAURLFileEntry.get()), outputPath)
            ScrapeResultText = "Successfully scraped! Check for ScrapedPDFs file in parent directory"
            textColor = "green"
        except PermissionError as pe:
            logging.error(f'Permission Error: {pe}')
            ScrapeResultText = f"Permission error: Make sure to close input ({inputPath}) and output ({outputPath}) TransAmerica URL File"
            textColor = "red"
        except OSError as ose:
            if ose.errno == 22 or ose.errno == 2:
                logging.error(f'Type Error: {ose}')
                ScrapeResultText = f"Invalid argument entered.  Please check {inputPath}"
            else:
                logging.error(f'ERROR: {ose}')
                ScrapeResultText = f"Error: {str(ose)}"
            textColor = "red"
        except Exception as e:
            logging.error(f'ERROR: {e}')
            print(type(e))
            ScrapeResultText = f"Error: {str(e)}"
            textColor = "red"


        topResultLabel.config(text = ScrapeResultText, fg=textColor)
        window.update()
        

        

    def handlePDFdownloading():
        global downloadResultText

        logging.info("Save PDFs button clicked")

        downloadResultText = "Button clicked"
        bottomResultLabel.config(text = downloadResultText)
        


    window.geometry("450x500")
    window.title("PDF Harvesting Application")

    top_frame = tkinter.Frame(window).pack()
    bottom_frame = tkinter.Frame(window).pack(side = "bottom")

    topLabel = tkinter.Label(top_frame, text="Enter TransAmerica URL file location (Ex: C:...xlsx): ")
    topWarning = tkinter.Label(top_frame, text="REMEMBER TO CLOSE THE INPUT AND OUTPUT FILES")
    TAURLFileEntry = tkinter.Entry(top_frame, width = 50)
    processURLButton = tkinter.Button(top_frame, text = "Scrape PDFs", command = handlePDFScraping)
    topResultLabel = tkinter.Label(top_frame, wraplength = 400)
    
    topLabel.pack(padx = 20, pady = 20)
    TAURLFileEntry.pack()
    topWarning.pack()
    processURLButton.pack(pady=20)
    topResultLabel.pack()

    bottomLabel = tkinter.Label(bottom_frame, text="Enter PDF file location (Ex: C:...xlsx)")
    PDFFileEntry = tkinter.Entry(bottom_frame, width = 50)
    downloadPDFButton = tkinter.Button(bottom_frame, text = "Save PDFs", command = handlePDFdownloading)
    bottomResultLabel = tkinter.Label(bottom_frame)

    bottomLabel.pack(padx = 20, pady = 20)
    PDFFileEntry.pack()
    downloadPDFButton.pack(pady=20)
    bottomResultLabel.pack()



    window.mainloop()
