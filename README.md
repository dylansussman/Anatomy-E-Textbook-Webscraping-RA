# README

# Table of Contents
1. [Overview](#Overview)
2. [Requirements](#Requirements)
3. [Components](#Components)
4. [Running](#Running)

## Overview
The following project contains the files and code used to perform webscraping on Anatomy E-Textbooks, to write bold terms to Excel sheets, and to compare
the bold terms (using one of the textbooks as a key) to determine how many times said terms appear in the different textbooks. This README will detail
the different components in this project, and how to run those components successfully. There is additional documentation within the code files that
goes into more specific details as needed since this README is more of an overview and is intended to give one enough information to get one or more
of this project's components running.

## Requirements
This project is written entirely in Python 3.0, so ensure that at least that version of Python has been installed, in addition to the Python interpreter to run this code. Additional details and information can be seen [here](https://www.python.org).

There are four dependencies used in this project. The first is [Selenium WebDriver](https://www.selenium.dev/documentation/webdriver/) which is the webscraping framework used in this project. The documentation is available [here](https://www.selenium.dev/documentation/webdriver/). The second is [OpenPyXL](https://openpyxl.readthedocs.io/en/stable/index.html) which is a framework used for all things related to reading from and writing to Excel sheets. The documentation is available [here](https://openpyxl.readthedocs.io/en/stable/index.html). The third is [python-docx](https://python-docx.readthedocs.io/en/latest/index.html) which is a framework used for all things related to reading from and writing to Word documents. The documentation is available [here](https://python-docx.readthedocs.io/en/latest/index.html). The fourth is [striprtf](https://pypi.org/project/striprtf/) which is used to read from a .txt file when logging in the the e-textbook website. The documentation is available [here](https://pypi.org/project/striprtf/).

To ensure these are installed run the following in the terminal/command line:

``` $ pip install selenium```

``` $ pip install openpyxl ```

``` $ pip install python-docx ```

``` $ pip install striprtf```

## Components
There are two components to this project. The first is the webscraping of the online textbooks and the input of the bold terms in Excel sheets. All of the
code for this component is contained in the Webscaping directory of this project. The second component is the comparison of these terms, and this code is all
contained in the Term_Comparison directory.

### Webscraping
The code for all the webscraping is broken down between 4 files:
- scraping_main.py
  - This is the file to run for the webscraping and writting the terms to an Excel sheet. Everything except the code for the one textbook to be scraped should be
  commented out. New code is written for each textbook, with the differences between each being the link to the textbook, the path/name of the output sheet, the
  path/name of the .txt file containing scraped info in a JSON format, and the pathes used to determine what bold words to scrape from the textbook.   
- book_scraper.py
  - This file contains the bookScraper class which has functions to deal with webscraping at the textbook level such as logging in, creating an Excel Workbook
  for a given textbook, and setting up the data structures necessary to store data, then looping through the chapters of a given textbook to get bold terms. The
  bulk of the code for this class is within the get_book_data function which sets up data structures and invokes the chapterScraper class. 
- chapter_scraper.py
  - This file contains the chapterScraper class which has functions to deal with webscraping at the chapter level such as getting chapter headers/sections, getting
  bold terms from the chapter, and writing to Excel worksheets. This class should only be invoked within the bookScraper class to ensure the code stays modular and
  organized. 
- docx_scraper.py
  - This file contains code to scrape Word documents (.docx) which was only used for the McGraw Hill textbooks. All of the code needed tp scrape for bold terms
  and output them to an Excel sheet is contained here.

There is an additional directory, Textbook_Data, which contains text files for each textbook containing their bold terms in a JSON format. This was stored to be able
to redo/reformat Excel sheets without having to rescrape textbooks because the scraping took a lot of time; thus, allowing for quick changes to the Excel sheets without
wasting the time to rescrape textbooks since the actual data remained the same.

Within the code in these files, there are many comments labeled as a "NOTE" which signifies code that may need to be changed when scraping different textbooks. I have done
my best to label these areas, but there were many small tweaks made from textbook to textbook in order to get all of the correct bold terms. This is due to the fact that
the HTML was different for many of the textbooks, and these changes account for that. Code noted under these comments should only be uncommented when scraping a textbook
specified by that comment. To comment out a line(s) of code, simply place a "#" at the front of the line or highlight the line(s) and use the keyboard shortcut
command/control + /. To uncomment code, simply remove the "#" from the line(s) or use the same keyboard shortcut to comment line(s) out.

### Term_Comparison
The code for all the term comparisons is broken down between 3 files:
- compare_main.py
  - This is the file to run to initiate the term comparisons and write the result to an Excel sheet. There are comments throughout the file detailing the specifics
  of what is happening at different points in the code.
- comparison_mappings.py
  - This file contains the copmarisonMappings class which creates multiple mappings required to do the term comparisons. These mappings are denoted in an Excel sheet listing general chapter names, the titles of all the textbooks scraped, and the corresponding chapters in each textbook. The original title of this Workbook is "Textbook Chapter Comparisons," with the Worksheet titled "Comparison Key." These can certainly be changed as long as the code is updated accordingly (more details in the documentation within these files). There is extensive documentation within this file detailing the different data structures used and explaining each mapping's relationship and what it is used for.
- comparison_key.py
  - This file contains the comparisonKey class which creates a wordbank/key from the textbook that is determined to be the comparator (textbook used as the key). The comparator is determined by the textbook and its chapters that are in column B of the Excel sheet mentioned above. Thus the comparator can easily be changed by rearraging the sheet and inserting the textbook that one wishes to use as the workbank/key to the right of Column A, resulting in it becoming the new Column B and shifting all other columns to the right.

## Running

### Webscraping
To run the webscraping component of this project, open Webscraping/scraping_main.py. Ensure that the only portion uncommented is that of the textbook that needs to be scraped. Everything else in the file (except imports), should be comented out. It is also possible to scrape more than one textbook at a time if one wishes (uncomment all textbook sections that need to be scraped); however, this may take some time. Details on commenting/uncommenting that can be found [here](#webscraping). Run that file using the Python interpreter and this code will webscrape the textbook(s) and output them to a spreadsheet (one spreadsheet/textbook).

### Term Comparison
To run the term comparison component of this project, open Term_Comparison/compare_main.py and run that file using the Python interpreter. This code will compare all the bold terms from every textbook specified and write the output to an Excel spreadsheet (denoted by the constant OUTPUT_SHEET_NAME in compare_main.py).