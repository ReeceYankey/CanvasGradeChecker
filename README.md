# Overview
CanvasWebScraper.py launches a chrome emulator where you just have to log in to canvas and it will automatically grab existing grades. It then stores these into csv files, updates the Galipatia Academic Success Database.xlsx accordingly, and saves it into a file called updated.xlsx. If you run UpdateFromCSV.py directly, it uses the existing csv files from a previous run to update the spreadsheet.

# Dependencies
use pip install on the following:
1. selenium
2. pandas
3. beautifulsoup4
4. decouple

# Setup
1. download the [chrome driver](https://chromedriver.chromium.org/downloads) for your version of chrome
2. run the program and follow the instructions in the console
