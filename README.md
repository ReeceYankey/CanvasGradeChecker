# Overview

CanvasWebScraper.py launches a chrome emulator where you just have to log in to canvas and it will automatically grab existing grades. It then stores these into csv files, updates the Galipatia Database Template.xlsx accordingly, and saves it into a file called updated.xlsx. If you run UpdateFromCSV.py directly, it uses the existing csv files from a previous run to update the spreadsheet.

# Requirements

Tested Chrome and Firefox on Windows 10

Tested Chromium and Firefox on Manjaro Linux

Tested only on Python 3.9.2

Currently only supports Virginia Tech users

Use ``pip install requirements.txt`` to download all dependencies (or use your preferred method)

# Setup

1. Download Selenium Drivers

If using chrome (win10) or chromium (linux), either let it auto install or manually download the [chrome driver](https://chromedriver.chromium.org/downloads) for your version of chrome

If using Firefox, download the [gecko driver](https://github.com/mozilla/geckodriver/releases/tag/latest)

2. Run main.py and follow the instructions in the terminal

## Tips

If your chrome/gecko driver breaks, run Setup.py or edit settings.ini manually

If you have class files under /ClassData and want to skip the process of recollecting the grade data from Canvas, run UpdateFromCSV.py
