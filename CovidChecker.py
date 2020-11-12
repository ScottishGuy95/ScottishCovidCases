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
def getFormattedDate(formatted=True):
    """
    Returns the data in one of 2 formatted options.
    Parameters:
        formatted (Boolean): True returns the data as expected for a URL, otherwise returns DDMMYY
    Return:
        dateStr (str): A str representing the formatted date - e.g. 25October2020
    """
    # Get date in DDMMYYYY format
    today = date.today()  # Create an object to get the date details
    dayNum = today.strftime("%d")  # Returns today's date in a int
    month = today.strftime("%B")  # Returns today's month in word format
    year = today.strftime("%Y")  # Returns year in 4 digits
    # Check if formatting is required or not, and return the value as a String
    if formatted:
        # Returns data, formatted to be used as part of a URL
        dateStr = dayNum + "%2B" + month + "%2B" + year
    else:
        # Returns today's date in format - DDMonthYear, e.g. 25October2020
        dateStr = dayNum + month + year
    return dateStr


def getFile(url, file):
    """
    Downloads a file from the given URL
    Parameters:
        url (str): The URL of the Scottish Govs covid data website
        file (str): The name of the file to look for on the Scottish Gov website
    """
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


def getAllLinks():
    """
    Scans a web page for all URLs, returning them in a list
    Return:
        links (list): A list of all the URLs found on the Scot Gov COVID-19 data website
    """
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


def getFileNameFromLinks():
    """
    Returns a String to be used as the file name, for the downloaded file
    Return:
        filename (str): A name based of the currently available COVID-19 file on the Scot Gob website
    """
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
        sys.exit()


# ------------------------------------------------------------------------------------------------------------------
# ------------------------------------------------------------------------------------------------------------------

# A function to handle outputting error messages for each function
# Did this to avoid having similar ending statements in each function
# TODO: Use this for now, but I do need to add proper error handling throughout the program
#  Meaning this WILL be deleted at the end!
def endIt(errorMsg):
    print(errorMsg)
    print('Ending the program')
    sys.exit()


def getHealthBoardList():
    """
    Returns a list of each Health Boards names, according to the column data in the Excel sheet
    Return:
        healthBoardsList (list): A list of each Health Boards name
    """
    healthBoardsList = []                           # Used to store all the health board names
    # Loops from the first available health board name (col 2) to the last health board name (col 16)
    for row in sheet.iter_rows(min_row=3, min_col=2, max_row=3, max_col=16):
        for cell in row:                            # Get the cell of the iteration
            healthBoardsList.append(cell.value)     # Add the health board name to the list
    return healthBoardsList


def getNewDataFromAllBoards():
    """
    Returns the most recent case totals for each health board
    Returns:
        newDataList (list): A list of each health boards case totals
    """
    newDataList = []                                        # The list to store the case values
    for col in range(2, 17):                                # Loop each health board area & Scottish total
        cell = sheet.cell(row=sheet.max_row, column=col)    # Store the cell for the current health board on the row
        newDataList.append(cell.value)                      # Get the data for that cell  (stores case numbers)
    return newDataList                                      # Returns all the case values in the list


def getNewScotlandTotal():
    """
    Returns the Scottish total cases from the Excel data
    Return:
        (int): Total cases in Scotland
    """
    # Reads the last column of data, on the last row of data, and returns that cells value
    return sheet.cell(row=sheet.max_row, column=16).value


def getSpecificHealthBoardNewCases(healthBoard):
    """
    Takes a Health Boards full name, and returns its total cases
    Parameters:
        healthBoard (str): The name of the health board (Should match that of the Excel sheet)
    Return:
        (int): The data from the column for the given health board
    """
    columnNum = getHealthBoardColumnNum(healthBoard)    # Gets the column number, for the given health board
    newData = getNewDataFromAllBoards()                 # Stores all of the newest available total cases in a list
    return newData[columnNum]                   # From the list of data, select the element from the given column num


def getHealthBoardColumnNum(healthBoard):
    """
    Takes a health boards full name, and returns the column number from the Excel sheet
    Parameters:
        healthBoard (str): The name of the health board (Should match that of the Excel sheet)
    Return:
        (int): The column number of the health board
    """
    for col in range(2, 17):            # Loops all of the columns of health boards
        if sheet.cell(row=3, column=col).value == healthBoard:  # Check if the current health board matches the request
            # Remove 2 from the column number, as the column in excel, starts at 2
            # so removing 2 the first element back to 1, so it becomes the 'first' column again
            return col - 2              # Returns the column number


def getHealthBoardCasesOverPeriod(timePeriod, healthBoard='all'):
    """
    Returns case numbers over a given period, for either all health boards or the specifically requested health board
    Parameters:
        timePeriod (int): A number representing how many days of data to look back over
        healthBoard (str): The name of health board, or 'all' or blank, to get every health board
    Return:
        healthBoardList (list): A list for each health boards cases over the requested period
        (int): The value for the cases over the requested time period
    """

    # 39 was selected due to formatting in the Gov file
    # When cases were less than 5, as '*' was displayed, for disclosure reasons
    # So, to avoid a ValueError, 39 was the first row of data to not have an *
    # TODO: Find a way to accommodate the asterix value
    max = sheet.max_row - 39
    try:
        length = int(timePeriod)
    except:
        print('ERROR: You must give a number between 1 and ' + str(max) + '\nEnding Program.')
        sys.exit()
    if int(length) < 1 or int(length) > max:
        errorMsg = 'ERROR: Given value is not valid for data available. Please enter a value between 1 and ' + str(max)
        endIt(errorMsg)

    if healthBoard == 'all':
        print('Getting all health boards cases over ' + str(length) + ' days')  # TODO Remove this after list handling
        healthBoardList = []
        for col in range(2, 17):
            newCell = sheet.cell(row=sheet.max_row, column=col).value
            olderCell = sheet.cell(row=sheet.max_row - length, column=col).value
            data = newCell - olderCell
            healthBoardList.append(data)
        return healthBoardList
    else:
        healthBoard = getHealthBoardFullName(healthBoard)
        print('Getting the last ' + str(length) + ' days of cases for ' + healthBoard)  # TODO Remove after CMD handling
        healthBoardColNum = getHealthBoardColumnNum(healthBoard)
        newCell = sheet.cell(row=sheet.max_row, column=healthBoardColNum).value
        olderCell = sheet.cell(row=sheet.max_row - length, column=healthBoardColNum).value
        return newCell - olderCell


def outputHandler(locations, values):
    if type(locations) == list and type(values) == list:
        print('Received a list, so it should be ALL health boards')
        maxLenElement = max(locations, key=len)
        for x in range(len(locations)):
            spacing = ' ' * (len(maxLenElement) - len(locations[x]) + 2)
            print(locations[x] + spacing + ' | ' + str(values[x]))

    elif type(locations) == str and type(values) == int:
        print('Received a str, so it will just be a single health board')
        print(locations + '\t|\t' + str(values))


# ------------------------------------------------------------------------------------------------------------------
# ------------------------------------------------------------------------------------------------------------------

# Reads the list of health boards and returns them in a list or prints them
# Depending if printIt is True or not
def getHealthBoardsList(printIt=False):
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
def getHealthBoardFullName(location):
    locationsList = getHealthBoardList()
    healthBoardNameFull = ''                        # Stores the health board name
    for x in range(len(locationsList)):             # Loops the entire list of health boards
        if location.lower() in locationsList[x].lower():    # If the given finds a match in the list of locations
            healthBoardNameFull = locationsList[x]  # Store that full health board name in a variable
    if healthBoardNameFull == '':                   # If there is no match, print error, print valid names and end
        print('Error - Invalid name given, please enter a name that matches one of the following:')
        print(getHealthBoardsList(printIt=True))           # Prints the health board names
        print('Ending program')
        sys.exit()                                  # Uses the sys package to end the program, as no valid name given
    elif healthBoardNameFull in locationsList:      # If there is a matching name in the location list, return as String
        return healthBoardNameFull


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

# Start by checking if the file is available
print("Checking for new data...")
getFile(fileURL, newestFile)    # Downloads newest file is available or uses older file

# File management - Clear out any other older files
print('Checking if there are any old files to clear...')
count = 0
for theFile in os.listdir(os.getcwd() + '\\ExcelFiles\\'):
    if theFile != newestFile:
        count += 1
        # print('need to bin this file: ' + str(os.getcwd() + '\\ExcelFiles\\' + theFile))
        send2trash(os.getcwd() + '\\ExcelFiles\\' + theFile)
if count >= 1:
    print('Deleted ' + str(count) + ' files.')
else:
    print('No files to clear.')

# Loads the most recent excel file, data_only used to ignore any formulas, as we only need the actual values
excel = load_workbook('ExcelFiles//' + newestFile, data_only=True)

# Get the correct sheet from the Excel file
for theSheet in range(len(excel.sheetnames)):
    if excel.sheetnames[theSheet] == 'Table 1 - Cumulative cases':
        excel.active = theSheet             # The active sheet is used to access the sheets data
sheet = excel.active                        # Stores the active sheet in a usable variable

lastRowNum = sheet.max_row
newestDate = sheet.cell(row=lastRowNum, column=1)  # Get the data of the most available row of data
newestDate = str(newestDate.value)[:10]         # Remove the end of the String, as it stores an unusable time value


def testALL():
    return None


# outputHandler(getHealthBoardFullName('Grampian'), getSpecificHealthBoardNewCases(getHealthBoardFullName('Grampian')))
# print()
# outputHandler(getHealthBoardList(), getHealthBoardCasesOverPeriod(19))

print(getHealthBoardCasesOverPeriod.__doc__)

print('\nEnding before checking CMD input')
sys.exit()

if str(sys.argv[1]).lower() == 'today':
    # if str(sys.argv[2]).lower() in quickNames:
    #     print('checking todays values for ' + getHealthBoardFullName(str(sys.argv[2])))
    printLocationsAndValues(getNewTodayAll())

# TODO: Is this needed? urllib.request.urlcleanup() - https://docs.python.org/3/library/urllib.request.html
