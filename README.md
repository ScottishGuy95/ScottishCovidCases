# ScottishCovidCases.py
Analyses Scottish Covid-19 case numbers and returns specific case numbers


Data is taken from the Scottish Governments website here - [Data](https://www.gov.scot/publications/coronavirus-covid-19-trends-in-daily-data/)
## Usage
To download and use this script:
* Clone this repository
* Once downloaded into your chosen directory
    * In terminal go to the chosen directory
    * Install the virtual environment, type: `pip install virtualenv`
    * To create your virtual environment, type: `python -m venv venv`
    * Activate the virtual environment, type: `.\venv\Scripts\activate.bat`
    * Install required modules, type: `pip install -r requirements.txt`
    * Run the script, type: `ScottishCovidCases.py -h`
    

`ScottishCovidCases.py *argument*`

Optional Arguments:
```
* -h                    Returns the help message and exits
* -n                    Returns today's newest case numbers for each health board
* -s                    Returns the total number of cases for Scotland
* -a HEALTHBOARD        Takes a health board name, returning the case numbers for that health board
* -c DAYS HEALTHBOARD   Takes a number of days & a health board, returns the case numbers over that period
* -t                    Returns the total number of cases for every health board
```

## License
[MIT](https://choosealicense.com/licenses/mit/)