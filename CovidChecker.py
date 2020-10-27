#! python3
# CovidChecker.py - A program to check the new coronavirus cases in Scotland
# Data is taken from - https://www.gov.scot/publications/coronavirus-covid-19-trends-in-daily-data/


# pip installs
# bs4, requests, openpyxl

# Imports
from bs4 import BeautifulSoup as bs
import requests
from datetime import date
import urllib.request
from openpyxl import *
import os
import sys
import shutil


# Functions
# Default parameter is used to return the String in the required format to get the correct URL
# Passing False is used to get a un-formatted String - e.g. 25October2020
def getFormattedDate(formatted=True):
    # Get date in DDMMYYYY format
    today = date.today()            # Create an object to get the date details
    dayNum = today.strftime("%d")   # Returns today's date in a int
    month = today.strftime("%B")    # Returns today's month in word format
    year = today.strftime("%Y")     # Returns year in 4 digits
    # Will be used to return formatted date to add to URL string later
    if formatted:
        dateStr = dayNum + "%2B" + month + "%2B" + year
    else:
        dateStr = dayNum + month + year
    return dateStr


# Downloads a file from the given URL
def getFile(url, file):
    print("Downloading the most recent data from the Scottish Gov website")

    # Check if a file with that name already exists, if not, download a fresh copy
    if file in os.listdir(os.getcwd() + '\\ExcelFiles\\'):
        print('A file with today\'s date already exists, using that.')
    else:
        try:
            print('Downloading...')     # TODO: Remove unnecessary prints
            urllib.request.urlretrieve(url, file)   # Downloads to the same directory as the python file
        except urllib.error.URLError:
            print("Error! The given URL is incorrect, please check the URL")
            print("URL Given - " + url)
            print('Ending program.')
            sys.exit()
        print("File downloaded!")
        shutil.move(newestFile, 'ExcelFiles')


# Scans the web page for all links, returning them in a list
# TODO: Add some comments on how this step works, to make it clearer
def getAllLinks():
    url = 'https://www.gov.scot/publications/coronavirus-covid-19-trends-in-daily-data/'
    soup = bs(requests.get(url).text, 'html.parser')
    links = []

    for a in soup.find_all('a'):
        links.append(a['href'])
    return links


# Take a list of links and return the file name in a readable format
# TODO: Add some comments to make this step clearer
def getFileNameFromLinks():
    links = getAllLinks()
    fileName = ''

    try:
        for link in links:
            if '/binaries/content/documents/govscot/publications/statistics/2020/04/coronavirus-covid-19-trends-in-daily-data/documents/covid-19-data-by-nhs-board/' in link:
                if '?forceDownload' in link:
                    fileName = link

        stringPos = fileName.index('COVID-19%2B')  # Finds the start of the file name and stores its character position
        fileName = fileName[stringPos:]  # Removes the first part of the URL to get a cleaner file name
        # Removes all the characters from the String that are not needed
        fileName = fileName.replace('%2B', '')
        fileName = fileName.replace('?forceDownload=true', '')
        fileName = fileName.replace('dailydata-byNHSBoard', '')
        # The file name in a readable format
        return fileName
    except:
        print('Error! Found no file to download. Please check the URL has a valid file to download.\nIt is possible the file name has changed.')


# Variables
newestFile = getFileNameFromLinks()
today = getFormattedDate()      # Gets the date in a URL format to add to the source file URL
# TODO: Find a better way to format this URL in the IDE to stop it complaining
fileURL = "http://www.gov.scot/binaries/content/documents/govscot/publications/statistics/2020/04/coronavirus-covid-19-trends-in-daily-data/documents/covid-19-data-by-nhs-board/covid-19-data-by-nhs-board/govscot%3Adocument/COVID-19%2Bdaily%2Bdata%2B-%2Bby%2BNHS%2BBoard%2B-%2B" + today + ".xlsx?forceDownload=true"

# data_only ensures we only get cell values, not formulas
excel = load_workbook('ExcelFiles//' + newestFile, data_only=True)


# Code
# Start by checking if the file is available
print("Starting!")
getFile(fileURL, newestFile)
# TODO: File management, how to handle all the older files? Should be done at this stage

# Get the correct sheet from the Excel file
for theSheet in range(len(excel.sheetnames)):
    if excel.sheetnames[theSheet] == 'Table 1 - Cumulative cases':
        excel.active = theSheet
sheet = excel.active        # Set the active sheet

# Loop through all the rows of data, finding the most recent one
newRow = ''
for row in range(1, sheet.max_row+1):
    cell = sheet.cell(row=row, column=1)
    if cell.value is None:
        newRow = cell.row
print(newRow)

# TODO: Is this needed? urllib.request.urlcleanup() - https://docs.python.org/3/library/urllib.request.html
