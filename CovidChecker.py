#! python3
# CovidChecker.py - A program to check the new coronavirus cases in Scotland
# Data is taken from - https://www.gov.scot/publications/coronavirus-covid-19-trends-in-daily-data/
# Example file name - COVID-19+daily+data+-+by+NHS+Board+-+21+October+2020.xlsx
# Expected file URL:
# https://www.gov.scot/binaries/content/documents/govscot/publications/statistics/2020/04/coronavirus-covid-19-trends-in-daily-data/documents/covid-19-data-by-nhs-board/covid-19-data-by-nhs-board/govscot%3Adocument/COVID-19%2Bdaily%2Bdata%2B-%2Bby%2BNHS%2BBoard%2B-%2B21%2BOctober%2B2020.xlsx?forceDownload=true

# pip installs
# pip install openpyxl

# Imports
from datetime import date
import urllib.request
import openpyxl
# import os # TODO is this needed? Could be used to check if file exists in directory


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
def getFile(url):
    # TODO - Download the file and set the name that comes from the file online
    # If the file is not available, post a msg saying that and end
    # If a file is available, check its filename - if it already exists in the dir, post a msg saying that and use the already existing file
    # If the file available has a more recent date, download that and use it
    # If the file online already exists in the dl dir, msg saying - already got recent data, and use that data
    
    # Only shows starting URL for clarity, as URL is rather large
    print("Downloading the most recent data from the " + fileURL[:20] + " website")
    # TODO - Change download location to a temp folder? To allow the file to be deleted after use
    try:
        urllib.request.urlretrieve(url, 'COVID-' + getFormattedDate(False) + '.xlsx')   # Downloads to the same directory as the python file
    except urllib.error.URLError:
        print("Error! The given URL is incorrect, please check the URL")
        print("URL Given - " + url)
    print("File downloaded!")


# Variables
today = getFormattedDate()      # Gets the date in a URL format to add to the source file URL
# TODO - Find a better way to format this URL in the IDE to stop it complaining
fileURL = "http://www.gov.scot/binaries/content/documents/govscot/publications/statistics/2020/04/coronavirus-covid-19-trends-in-daily-data/documents/covid-19-data-by-nhs-board/covid-19-data-by-nhs-board/govscot%3Adocument/COVID-19%2Bdaily%2Bdata%2B-%2Bby%2BNHS%2BBoard%2B-%2B" + today + ".xlsx?forceDownload=true"


# Code
# Start by checking if the file is available
print("Starting!")
print(fileURL)
getFile(fileURL)

# urllib.request.urlcleanup() - https://docs.python.org/3/library/urllib.request.html
