# Overview
CanvasWebScraper.py launches a chrome emulator where you just have to log in to canvas and it will automatically grab existing grades. It then stores these into csv files, updates the Galipatia Academic Success Database.xlsx accordingly, and saves it into a file called updated.xlsx. If you run UpdateFromCSV.py directly, it uses the existing csv files from a previous run to update the spreadsheet.

# Requirements
Note: Must have Chrome installed on a Windows machine.

Use ```pip install requirements.txt``` to download all dependencies


# Setup
1. download the [chrome driver](https://chromedriver.chromium.org/downloads) for your version of chrome (or let it auto-install)
2. run the program and follow the instructions in the terminal
