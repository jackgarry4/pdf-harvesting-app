import logging


def configure_logging(logPath):
    # Configure logging to console
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.ERROR)  # Set the desired logging level for console output
    console_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    console_handler.setFormatter(console_formatter)

    # Configure logging to file
    file_handler = logging.FileHandler(logPath)  # Specify the file path
    file_handler.setLevel(logging.INFO)  # Set the desired logging level for the file
    file_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(file_formatter)

    # Get the root logger and add the configured handlers
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.INFO)  # Set the desired logging level for the root logger
    root_logger.addHandler(console_handler)
    root_logger.addHandler(file_handler)