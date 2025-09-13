import requests
import pandas as pd
import xlwings as xw
import time

def fetch_data():
    url = "https://www.nseindia.com/api/option-chain-indices?symbol=NIFTY"
    headers = {"User-Agent": "Mozilla/5.0"}
    session = requests.Session()
    session.get("https://www.nseindia.com", headers=headers)
    data = session.get(url, headers=headers).json()
    return data['records']['data'], data['records']['underlyingValue']

def analyze_trend(prev_price, prev_oi, curr_price, curr_oi):
    if curr_price > prev_price and curr_oi > prev_oi:
        return "Bullish"
    elif curr_price < prev_price and curr_oi > prev_oi:
        return "Bearish"
    elif curr_price > prev_price and curr_oi < prev_oi:
        return "Short Covering"
    elif curr_price < prev_price and curr_oi < prev_oi:
        return "Longs Unwinding"
    else:
        return "Neutral"

def update_excel():
    data, spot_price = fetch_data()
    atm_strike = round(spot_price / 50) * 50
    strike_range = [atm_strike + i*50 for i in range(-6, 7)]

    wb = xw.Book('NiftyTrendMulti.xlsm')
    sheet = wb.sheets['Trend']

    for i, strike in enumerate(strike_range, start=2):
        row = next((item for item in data if item['strikePrice'] == strike), None)
        if not row:
            continue

        ce = row.get('CE', {})
        pe = row.get('PE', {})

        curr_ce_price = ce.get('lastPrice', 0)
        curr_ce_oi = ce.get('openInterest', 0)
        curr_pe_price = pe.get('lastPrice', 0)
        curr_pe_oi = pe.get('openInterest', 0)

        # Read previous values
        prev_ce_price = sheet.range(f'C{i}').value or 0
        prev_ce_oi = sheet.range(f'D{i}').value or 0
        prev_pe_price = sheet.range(f'F{i}').value or 0
        prev_pe_oi = sheet.range(f'G{i}').value or 0

        # Analyze trends
        ce_trend = analyze_trend(prev_ce_price, prev_ce_oi, curr_ce_price, curr_ce_oi)
        pe_trend = analyze_trend(prev_pe_price, prev_pe_oi, curr_pe_price, curr_pe_oi)

        # Update Excel
        sheet.range(f'A{i}').value = strike
        sheet.range(f'B{i}').value = pd.Timestamp.now()
        sheet.range(f'C{i}').value = curr_ce_price
        sheet.range(f'D{i}').value = curr_ce_oi
        sheet.range(f'E{i}').value = ce_trend
        sheet.range(f'F{i}').value = curr_pe_price
        sheet.range(f'G{i}').value = curr_pe_oi
        sheet.range(f'H{i}').value = pe_trend

    wb.save()
    wb.close()

if __name__ == "__main__":
    update_excel()
