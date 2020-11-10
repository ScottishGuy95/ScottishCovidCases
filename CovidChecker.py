#! python3
# CovidChecker.py - A program to check the new coronavirus cases in Scotland
# Data is taken from - https://www.gov.scot/publications/coronavirus-covid-19-trends-in-daily-data/

# TODO: Setup a -h or help option to get details on how to use the program
# TODO: Remove unnecessary prints from WHOLE program
# TODO: Program needs an internet connection (for now, change in future version if I can)
#  IF there is no already existing excel file to use - download a fresh version.
#  So it should start, check for a file, if no file exists, try to get a fresh file. If it fails, let exception handle.
#  If there is a file, but its not with today's date, try to get a fresh file, if fails, exception handles it.
#  If there is a file and you cant get a new file, post a message and just use older data
#  If there is a file and its with today's date, use that!
# pip installs
# bs4, requests, openpyxl, send2trash

# Imports
from bs4 import BeautifulSoup as bs
import requests
from datetime import date
import urllib.request
from openpyxl import *
import os
import sys
import shutil
from send2trash import send2trash


# Functions
# Default parameter is used to return the String in the required format to get the correct URL
# Passing False is used to get a un-formatted String - e.g. 25October2020
def getFormattedDate(formatted=True):
    # Get date in DDMMYYYY format
    today = date.today()  # Create an object to get the date details
    dayNum = today.strftime("%d")  # Returns today's date in a int
    month = today.strftime("%B")  # Returns today's month in word format
    year = today.strftime("%Y")  # Returns year in 4 digits
    # Check if formatting is required or not, and return the value as a String
    if formatted:
        dateStr = dayNum + "%2B" + month + "%2B" + year
    else:
        dateStr = dayNum + month + year
    return dateStr


# Downloads a file from the given URL
def getFile(url, file):

    # Check if a file with that name already exists, if not, download a fresh copy
    # Uses the os module to check the name of the file to download, against all the files in the ExcelFiles directory
    if file in os.listdir(os.getcwd() + '\\ExcelFiles\\'):
        print('A file with today\'s date already exists, using that.')
    else:
        try:
            # Attempts to use the urllib module to download the given file from the Scot Gov website
            print('Downloading the new data')
            urllib.request.urlretrieve(url, file)  # Downloads to the same directory as the python file
        # If there is an issue with the given URL, display an error, the given URL and end the program
        except urllib.error.URLError:
            print("Error! The given URL is incorrect, please check the URL")
            print("URL Given: " + url)
            print('Ending program as no data available to analyse')
            sys.exit()  # Ends the program as the URL failed, so no data available
        print("File downloaded!")
        # Move the newly downloaded file into the correct directory
        # TODO: Add some sort of error in case moving the file returns an error
        shutil.move(newestFile, 'ExcelFiles')


# Scans a web page for all links, returning them in a list
def getAllLinks():
    # The website that hosts the files
    url = 'https://www.gov.scot/publications/coronavirus-covid-19-trends-in-daily-data/'
    # Create a soup object, that reads the HTML content and create a list to later store the results
    soup = bs(requests.get(url).text, 'html.parser')
    links = []

    # Loop through all of the content in the soup object, adding each hyperlink to the list and return the list
    # This allows the hyperlinks to be searched for the correct file later, by 'etFileNameFromLinks()
    for a in soup.find_all('a'):
        links.append(a['href'])
    return links


# Take a list of ULRs and return the file name in a readable format
def getFileNameFromLinks():
    links = getAllLinks()  # Get a list of URLs to search through for the Excel file
    fileName = ''  # The intended filename for the download object

    try:
        # Loop through all of the URLs in the list, for the specific file name and store in the fileName variable
        for link in links:
            if '/binaries/content/documents/govscot/publications/statistics/2020/04/coronavirus-covid-19-trends-in' \
               '-daily-data/documents/covid-19-data-by-nhs-board/' in link:
                if '?forceDownload' in link:
                    fileName = link

        stringPos = fileName.index('COVID-19%2B')  # Finds the start of the file name and stores its character position
        # Removes the first part of the URL to get a cleaner file name, using the above index as the cut off point
        fileName = fileName[stringPos:]
        # Removes all the characters from the String that are not needed (cleaning it up)
        fileName = fileName.replace('%2B', '')
        fileName = fileName.replace('?forceDownload=true', '')
        fileName = fileName.replace('dailydata-byNHSBoard', '')
        # The file name in a readable format
        return fileName
    except:
        # Print warning message to user
        print('Error! Found no file to download. Please check the URL has a valid file to download.\nIt is possible '
              'the file name has changed.')


# ------------------------------------------------------------------------------------------------------------------
# ------------------------------------------------------------------------------------------------------------------

# Returns a list of each Health Boards name, according to the data available in the sheet
def getHealthBoardList():
    healthBoardsList = []                           # Used to store all the health board names
    # Loops from the first available health board name (col 2) to the last health board name (col 16)
    for row in sheet.iter_rows(min_row=3, min_col=2, max_row=3, max_col=16):
        for cell in row:                            # Get the cell of the iteration
            healthBoardsList.append(cell.value)     # Add the health board name to the list
    return healthBoardsList


def getNewDataFromAllBoards():
    newDataList = []
    for col in range(2, 17):
        cell = sheet.cell(row=sheet.max_row, column=col)
        newDataList.append(cell.value)
    return newDataList


def getNewScotlandTotal():
    return sheet.cell(row=sheet.max_row, column=16).value


def getSpecificHealthBoardNewCases(healthBoard):
    # Get the column number of that health board
    for col in range(2, 17):
        if sheet.cell(row=3, column=col).value == healthBoard:
            column = col
    newData = getNewDataFromAllBoards()
    return newData[column-2]

# ------------------------------------------------------------------------------------------------------------------
# ------------------------------------------------------------------------------------------------------------------

# TODO: If using this, clean up format so its clear
# Prints the newest row of data, alongside all of their health board titles
def printNewestCasesOnly(lastRow):
    for col in range(2, 17):  # Loop from column 2 (first health board) to 17 (Scottish total)
        cell = sheet.cell(row=lastRow, column=col)  # Grabs the specific cell each increment of the loop
        # If it is column A, only show the date, not the time
        # Otherwise just print the data for each column
        if col == 1:
            print(str(cell.value)[:10] + ' | ', end='')
        else:
            print(str(cell.value) + ' | ', end='')


# Returns a list of the most recent case data
# Requires the row number of the most recent data
def getNewestCases(lastRow):
    newCasesList = []  # Stores the row of data
    for col in range(2, 17):  # Loops from column A - P
        cell = sheet.cell(row=lastRow, column=col)  # Each increment, get the cell data from that column
        newCasesList.append(str(cell.value))  # Add the cells value to the list
    return newCasesList  # Return the list


# TODO: Can this be merged with the above function - pass it how many rows to go back
def getSecondNewestCases(lastRow):
    secondNewList = []  # Stores the row of data
    for col in range(2, 17):  # Loops from column A - P
        cell = sheet.cell(row=lastRow-1, column=col)  # Each increment, get the cell data from that column
        secondNewList.append(str(cell.value))  # Add the cells value to the list
    return secondNewList  # Return the list


# Reads the list of health boards and returns them in a list or prints them
# Depending if printIt is True or not
def getLocations(printIt=False):
    locationsList = []  # Used to store the health board names
    for col in range(2, 17):  # Loop columns A - P
        cell = sheet.cell(row=3, column=col)  # Grab the cell from the current column per increment
        if not printIt:  # Not printing to console, so append the cells value to the list
            locationsList.append(str(cell.value))
        else:
            print(str(cell.value) + ' | ', end='')  # Print each cell value as a String, without a new line
    if not printIt:  # Not printing to console, so return the list
        return locationsList
    else:
        print()  # Print to get to a fresh line after using end=''


# Takes a health board name, and returns the exact spelling/format of that health board
# Used to allow more user friendly input of names - e.g. Grampian becomes NHS Grampian
# Expects a name to check for and a list of locations to compare against
def getLocationName(location):
    locationsList = getHealthBoardList()
    healthBoardNameFull = ''                        # Stores the health board name
    for x in range(len(locationsList)):             # Loops the entire list of health boards
        if location.lower() in locationsList[x].lower():    # If the given finds a match in the list of locations
            healthBoardNameFull = locationsList[x]  # Store that full health board name in a variable
    if healthBoardNameFull == '':                   # If there is no match, print error, print valid names and end
        print('Error - Invalid name given, please enter a name that matches one of the following:')
        print(getLocations(printIt=True))           # Prints the health board names
        print('Ending program')
        sys.exit()                                  # Uses the sys package to end the program, as no valid name given
    elif healthBoardNameFull in locationsList:      # If there is a matching name in the location list, return as String
        return healthBoardNameFull


def printAllTotals(healthBoards, todaysCases):
    for x in range(len(healthBoards)):
        print(healthBoards[x] + ' | ' + todaysCases[x])


def getNewTodayAll():
    today = getNewestCases(getLastRowNumber(sheet))
    yesterday = getSecondNewestCases(getLastRowNumber(sheet))
    new = []
    for values in range(len(today)):
        difference = int(today[values]) - int(yesterday[values])
        new.append(str(difference))
    return new


def printLocationsAndValues(caseList):
    locations = getLocations(False)
    for x in range(len(locations)):
        print(locations[x] + ' | ' + caseList[x])


# Code
# TODO: Set up error checking for all possible arguments
# if len(sys.argv) < 1:
#     print('Usage: CovidChecker.py [Location]')
#     print('Example: CovidChecker.py Highland')
#     sys.exit()
# else:
#     requestedLocation = sys.argv[1]
#     print('User asking for data on - ' + requestedLocation)


newestFile = getFileNameFromLinks()     # Stores the expected file name from the website
today = getFormattedDate()              # Gets the date in a URL format to add to the source file URL
# The URL of where the most recent file will be
fileURL = "http://www.gov.scot/binaries/content/documents/govscot/publications/statistics/2020/04/coronavirus-covid" \
          "-19-trends-in-daily-data/documents/covid-19-data-by-nhs-board/covid-19-data-by-nhs-board/govscot" \
          "%3Adocument/COVID-19%2Bdaily%2Bdata%2B-%2Bby%2BNHS%2BBoard%2B-%2B" + today + ".xlsx?forceDownload=true "

# A list of letters representing the columns used in the Excel file
columnLtr = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P']
lastColLtr = columnLtr[-1]              # The last column letter stored as a String
lastColNum = 16                         # The column number of the last letter

# Shorter versions of each of possible locations - Saves typing in the full name for users convience
# TODO - What was the purpose of this again - Getting the full name from the sheet and interpreting the users input,
#  is the better way to do it
# quickNames = ['ayrshire', 'arran', 'ayrshire & arran', 'borders', 'dumfries', 'galloway', 'gife', 'forth', 'valley',
#              'forth valley', 'grampian', 'greater glasgow', 'greater glasgow & clyde', 'glasgow', 'clyde', 'highland',
#               'lanarkshire', 'lothian', 'orkney', 'shetland', 'tayside', 'western isles', 'scotland']


# Start by checking if the file is available
print("Starting!")
getFile(fileURL, newestFile)    # Downloads newest file is available or uses older file

# File management - Clear out any other older files
print('Checking if there are any old files to clear.')
count = 0
for theFile in os.listdir(os.getcwd() + '\\ExcelFiles\\'):
    if theFile != newestFile:
        count += 1
        # print('need to bin this file: ' + str(os.getcwd() + '\\ExcelFiles\\' + theFile))
        send2trash(os.getcwd() + '\\ExcelFiles\\' + theFile)
if count >= 1:
    print('Deleted ' + str(count) + ' files.')

# Loads the most recent excel file, data_only used to ignore any formulas
excel = load_workbook('ExcelFiles//' + newestFile, data_only=True)

# Get the correct sheet from the Excel file
for theSheet in range(len(excel.sheetnames)):
    if excel.sheetnames[theSheet] == 'Table 1 - Cumulative cases':
        excel.active = theSheet             # The active sheet is used to access the sheets data
sheet = excel.active                        # Stores the active sheet in a usable variable

lastRowNum = sheet.max_row
# lastRow = getLastRowNumber(sheet)                     # Store the last row number
newestDate = sheet.cell(row=lastRowNum, column=1)  # Get the data of the most available row of data
newestDate = str(newestDate.value)[:10]         # Remove the end of the String, as it stores an unusable time value


def testALL():
    print(lastRowNum)
    print()
    print('--------------------')
    print()
    printNewestCasesOnly(lastRowNum)
    print()
    print('--------------------')
    print()
    print(getNewestCases(lastRowNum))
    print()
    print('--------------------')
    print()
    print(getLocations())
    print()
    print('--------------------')
    print()
    getLocations(True)
    print()
    print('--------------------')
    print()
    print(getLocationName('Highland', getLocations()))
    print()
    print('--------------------')
    print()
    printAllTotals(getLocations(), getNewestCases(lastRowNum))


print('\nEnding before checking CMD input')
sys.exit()

if str(sys.argv[1]).lower() == 'today':
    # if str(sys.argv[2]).lower() in quickNames:
    #     print('checking todays values for ' + getLocationName(str(sys.argv[2])))
    printLocationsAndValues(getNewTodayAll())

# TODO: Is this needed? urllib.request.urlcleanup() - https://docs.python.org/3/library/urllib.request.html
