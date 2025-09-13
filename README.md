# Nifty Options Tracker 📈
Developed a Python-powered Excel tool using xlwings, requests, and pandas to fetch live option chain data, analyze trends, and update a macro-enabled workbook with one-click refresh functionality.


This project automates the tracking of Nifty option chain data using Python and Excel.

## Features
- Fetches live CE/PE data using `requests`
- Analyzes trends with `pandas`
- Updates Excel workbook via `xlwings`
- One-click refresh using Excel macro

## Requirements
- Python 3.13+
- Modules: `requests`, `pandas`, `xlwings`, `openpyxl`, `schedule`

## Setup
1. Clone the repo or download files
2. Install dependencies:
pip install requests pandas xlwings openpyxl schedule

3. Open `NiftyTrendMulti.xlsm`
4. Press the refresh button to run the macro

## Author
Shubhodeep Kundu
