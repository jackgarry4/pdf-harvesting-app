from app.gui import PDFHarvestingApp
from config.logging_config import configure_logging
import logging 
from pathlib import Path
import tkinter
import threading


def main():

    #Configure logging 
    configure_logging( Path("LogFile.log"))
    logging.info("Application Opened")
    logging.info(f"Main thread: {threading.current_thread()}")

    #Run your gui
    window = tkinter.Tk()
    app = PDFHarvestingApp(window)
    app.run()
    


if __name__ == "__main__":
    main()