# journey-finder
## Description
A python script that allows the user to scrape 150+ Points of Interest and Restaurant data from [Tripadvisor.com](https://www.tripadvisor.com/) through the use of Chrome Driver. Once the location is confirmed, the automated driver will go to Trip Advisor and scrape and compile data, outputting an Excel file in the files directory. 
## Requirements
Python 3, Google Chrome Browser, Chrome Driver, Selenium and Openpyxl libraries for Python
## Installation
Use <code>pip install selenium</code> and <code>pip install openpyxl</code> to download packages. </br>
</br>
ChromeDriver is also required. For installation, click [HERE](https://sites.google.com/a/chromium.org/chromedriver/) for instructions. </br>
After ChromeDriver installation, navigate to journeyFinder.py and change the PATH variable to the filepath where ChromeDriver is installed on your local machine
## Execution
### Webscraper
To run, input <code>python3 journeyFinder.py</code> in terminal. </br>
The script will then prompt you to input a city name followed by a country or state. Following this input format is optional, but will provide more accurate results.
### CSV Converter
To run, input <code>python3 excel_to_csv.py</code> in terminal. You will need to provide the name of the file to convert, which will be in files.
