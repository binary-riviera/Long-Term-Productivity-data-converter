from openpyxl import load_workbook
from pathlib import Path
import requests
import csv
import json

def download_excel_database():
    url = 'http://longtermproductivity.com/download/BCLDatabase_online_v2.4.xlsx'
    r = requests.get(url, allow_redirects=False)
    open('BCLDatabase_online.xlsx', 'wb').write(r.content)

def extract_csv_files(workbook):
    print('Extracting core data into csv files...')
    sheets_to_export = ['GDP per capita', 'Labor Productivity', 'TFP', 'KI', 'AgeK']
    for s in sheets_to_export:
        sheet = workbook[s]
        values = sheet.values
        with  open('data/' + s + '.csv', 'w', newline='') as csv_file:
            writer = csv.writer(csv_file, delimiter=',')
            for line in values:
                writer.writerow(line)

def extract_extra_information(workbook):
    print('Extracting extra information into json files...')
    info = workbook['Info']
    citation = {'citation': {
        'authors' : info['B6'].value,
        'paper': info['B7'].value,
        'journal': info['B8'].value
    }}
    country_codes = {}
    row = 4 
    while True:
        country_code = info['F' + str(row)].value
        country_name = info['G' + str(row)].value
        if country_code is None:
            break
        country_codes[country_code] = country_name
        row += 1
    
    page_keys = {}
    row = 11
    while True:
        page = info['A' + str(row)].value
        series_name = info['B' + str(row)].value
        units = info['C' + str(row)].value
        if page is None:
            break
        page_keys[page] = {
            'Series Name' : series_name,
            'Units' : units
        }
        row += 1
    
    extra_information = {
        'citation' : citation,
        'country codes' : country_codes,
        'page key' : page_keys
    }

    print(json.dumps(extra_information, indent=4))

def extract_data(filename, extract_extra_data=True):
    wb = load_workbook(filename = filename, data_only=True)
    Path('./data').mkdir(parents=True, exist_ok=True)
    extract_csv_files(wb)
    if (extract_extra_data):
        extract_extra_information(wb)
    wb.close()

def main():
    print("Long Term Productivity data converter")
    print("Downloading excel database from 'http://longtermproductivity.com/download/BCLDatabase_online_v2.4.xlsx'")
    download_excel_database()
    extract_data('BCLDatabase_online.xlsx', extract_extra_data=True)

if __name__ == '__main__':
    main()