#! python3
# CovidChecker.py - A program to check the new coronavirus cases in Scotland
# Data is taken from - https://www.gov.scot/publications/coronavirus-covid-19-trends-in-daily-data/

# TODO: Remove unnecessary prints from WHOLE program
# TODO: Program should only need an internet connection
#  IF there is no already existing excel file to use.
#  Or to download a fresh version.
#  So it should start, check for a file, if no file exists, try to get a fresh file. If it fails, let exception handle.
#  If there is a file, but its not with today's date, try to get a fresh file, if fails, exception handles it.
#  If there is a file and you cant get a new file, post a message and just use older data
#  If there is a file and its with today's date, use that!
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
    # Check if formatting is required or not, and return the value as a String
    if formatted:
        dateStr = dayNum + "%2B" + month + "%2B" + year
    else:
        dateStr = dayNum + month + year
    return dateStr


# Downloads a file from the given URL
def getFile(url, file):
    # TODO: Find a way to handle files better

    # Check if a file with that name already exists, if not, download a fresh copy
    # Uses the os module to check the name of the file to download, against all the files in the ExcelFiles directory
    if file in os.listdir(os.getcwd() + '\\ExcelFiles\\'):
        print('A file with today\'s date already exists, using that.')
    else:
        # TODO: Other exceptions; download fails, no internet connection (this should not be needed if a file exists
        try:
            # Attempts to use the urllib module to download the given file from the Scot Gov website
            print('Downloading the new data')
            urllib.request.urlretrieve(url, file)   # Downloads to the same directory as the python file
            # TODO: If a new file is downloaded, delete the other excel files
        # If there is an issue with the given URL, display an error, the given URL and end the program
        except urllib.error.URLError:
            print("Error! The given URL is incorrect, please check the URL")
            print("URL Given - " + url)
            print('Ending program.')
            sys.exit()              # Ends the program as the URL failed # TODO: Change this once better filing is added
        print("File downloaded!")
        # Move the newly downloaded file into the correct diretory
        # TODO: Add some sort of error in case moving the file returns an error
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


def getLastRow(sheet):
    # Loop through all the rows of data, finding the most recent one
    lastRow = ''
    for row in range(4, sheet.max_row + 1):
        cell = sheet.cell(row=row, column=1)
        if not cell.is_date:
            lastRow = cell.row - 1
            break
    return lastRow


def printNewest(lastRow):
    printLocations()
    for col in range(2, 17):
        cell = sheet.cell(row=lastRow, column=col)
        if col == 1:
            print(str(cell.value)[:10] + ' | ', end='')
        else:
            print(str(cell.value) + ' | ', end='')


def getNewest(lastRow):
    newCasesList = []
    for col in range(2, 17):
        cell = sheet.cell(row=lastRow, column=col)
        newCasesList.append(str(cell.value))
    return newCasesList


def printLocations():
    for col in range(2, 17):
        cell = sheet.cell(row=3, column=col)
        print(str(cell.value) + ' | ', end='')
    print()


def getLocations():
    locationsList = []
    for col in range(2, 17):
        cell = sheet.cell(row=3, column=col)
        locationsList.append(str(cell.value))
    return locationsList


def printNewCases(locations, cases, newestDate):
    print('Cases from - ' + newestDate)
    for x in range(len(locations)):
        print(locations[x] + ' | ' + cases[x])


def getLocationName(location, locationsList):
    name = ''
    for x in range(len(locationsList)):
        if location in locationsList[x]:
            name = locationsList[x]
    if name == '':
        print('Error - Invalid name given, please enter a name that matches one of the following:')
        print(getLocations())
        print('Ending program')
        sys.exit()

    elif name in locationsList:
        return name


# Variables
newestFile = getFileNameFromLinks()
today = getFormattedDate()      # Gets the date in a URL format to add to the source file URL
# TODO: Find a better way to format this URL in the IDE to stop it complaining
fileURL = "http://www.gov.scot/binaries/content/documents/govscot/publications/statistics/2020/04/coronavirus-covid" \
          "-19-trends-in-daily-data/documents/covid-19-data-by-nhs-board/covid-19-data-by-nhs-board/govscot" \
          "%3Adocument/COVID-19%2Bdaily%2Bdata%2B-%2Bby%2BNHS%2BBoard%2B-%2B" + today + ".xlsx?forceDownload=true "
columnLtr = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P']
lastColLtr = columnLtr[-1]
lastColNum = 16
# Shorter versions of each of possible locations - Saves typing in the full name
quickNames = ['Ayrshire', 'Arran', 'Ayrshire & Arran', 'Borders', 'Dumfries', 'Galloway', 'Fife', 'Forth', 'Valley',
              'Forth Valley', 'Grampian', 'Greater Glasgow', 'Greater Glasgow & Clyde', 'Glasgow', 'Clyde', 'Highland',
              'Lanarkshire', 'Lothian', 'Orkney', 'Shetland', 'Tayside', 'Western Isles', 'Scotland']


# Start by checking if the file is available
print("Starting!")
getFile(fileURL, newestFile)
# TODO: File management, how to handle all the older files? Should be done at this stage
excel = load_workbook('ExcelFiles//' + newestFile, data_only=True)

# Get the correct sheet from the Excel file
for theSheet in range(len(excel.sheetnames)):
    if excel.sheetnames[theSheet] == 'Table 1 - Cumulative cases':
        excel.active = theSheet
sheet = excel.active        # Set the active sheet

# Store the last row number
lastRow = getLastRow(sheet)
newestDate = sheet.cell(row=lastRow, column=1)
newestDate = str(newestDate.value)[:10]

# Prints all of the new cases from the most recent value
# printNewCases(getLocations(), getNewest(lastRow), newestDate)
# print(getLocations())







# TODO: Is this needed? urllib.request.urlcleanup() - https://docs.python.org/3/library/urllib.request.html
