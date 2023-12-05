# PDF Harvesting Application
PDF Harvesting Application is a Python application that scrapes pdfs from TransAmerica and stores the pdfs into local root directory in organized manner with an accompanying excel file containing a file index.


![Git Demo Mid Run](image.png)


## Built With
* [Python](https://www.python.org/):Python is a versatile and powerful programming language.  It is known for its readability, ease of use, and extensive library.  Given tasks involving data manipulation, GUI development, and file handling, Python was a suitable choice.
* [Tkinter](https://docs.python.org/3/library/tkinter.html): Tkinter is the standard GUI toolkit for Python.  It provides a simple way to create GUI applications and is well-suited for smaller projects.  It being included with most Python installations, makes it convenient for the end user.
* [pandas](https://pandas.pydata.org/): Pandas is used for data manipulation and analysis.  In this project, pandas simplifies tasks such as filterin, aggregating, and transforming data.  Its DataFrame structure is particularly useful for tabular data. 
* [concurrent.futures](https://docs.python.org/3/library/concurrent.futures.html):  Multithreading is employed to parallelize tasks and improve performance.  Given that the application involves time-consuming operation like downloading PDFs and looping through large amount of data, using threads allows the execution of these tasks concurrently.  It is also useful to preven the GUI from becoming unresponsive.  
* [pyinstaller](https://pyinstaller.org/en/stable/):  Allows the packaging of Python applications into standalone executables.  This simplifies the deployment for the end-users who lack technical experience.


## Install
The program is run with python installed: 
`python main.py`




## Instructions

##

