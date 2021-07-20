# Overview

CanvasWebScraper.py launches a chrome emulator where you just have to log in to canvas and it will automatically grab existing grades. It then stores these into csv files, updates the Galipatia Database Template.xlsx accordingly, and saves it into a file called updated.xlsx. If you run UpdateFromCSV.py directly, it uses the existing csv files from a previous run to update the spreadsheet.

# Requirements

Must have Chrome installed on a Windows machine.

Tested only on Python 3.9.2

Currently only supports Virginia Tech users

Use ``pip install requirements.txt`` to download all dependencies

# Setup

1. download the [chrome driver](https://chromedriver.chromium.org/downloads) for your version of chrome (or let it auto-install)
2. run main.py and follow the instructions in the terminal

If your chrome driver breaks, run Setup.py

If you have class files under /ClassData and want to skip the process of recollecting the grade data from Canvas, run UpdateFromCSV.py
