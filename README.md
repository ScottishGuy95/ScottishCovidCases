# ScottishCovidCases.py
Analyses Scottish Covid-19 case numbers and returns specific case numbers


Data is taken from the Scottish Governments website here (COVID-19 data file)- [Data](https://www.gov.scot/publications/coronavirus-covid-19-trends-in-daily-data/)

Health Boards
* Ayrshire Arran
* Borders
* Dumfries Galloway
* Fife
* Forth Valley
* Grampian
* Greater Glasgow Clyde
* Highland
* Lanarkshire
* Lothian
* Orkney
* Shetland
* Tayside
* Western Isles
* Scotland

## Usage
Requires python3.8 installed
To download and use this script in Windows:
* Clone this repository
    * At the top of this page, select Green code button
    * If you have `git` installed, select the https option
    * Otherwise, download the Zip and extract
* Once downloaded into your chosen directory
    * In CMD go to the chosen directory
    * Install the virtual environment, type: `pip install virtualenv`
    * To create your virtual environment, type: `python -m venv venv`
    * Activate the virtual environment, type: `.\venv\Scripts\activate.bat`
    * Install required modules, type: `pip install -r requirements.txt`
    * Run the script, type: `python ScottishCovidCases.py -h`
    
## Arguments
`ScottishCovidCases.py [argument]`


```
* -h                    Returns the help message and exits
* -n                    Returns today's newest case numbers for each health board
* -s                    Returns the total number of cases for Scotland
* -a HEALTHBOARD        Takes a health board name, returns that health boards total cases
* -c DAYS HEALTHBOARD   Takes a number of days & a health board or 'all', returns the case numbers over that period
* -t                    Returns the total number of cases for every health board
* -hb                   Returns all health boards available
```

## Examples
Get a list of all available commands: `ScottishCovidCases.py -h`

Get a list of all of the Health Boards you can use: `ScottishCovidCases.py -hb`

Get today's newest case numbers: `ScottishCovidCases.py -n`

Get the last 7 days of case numbers in NHS Highland: `ScottishCovidCases -c 7 Highland`

## License
[MIT](https://choosealicense.com/licenses/mit/)