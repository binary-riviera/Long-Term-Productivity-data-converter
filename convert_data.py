from openpyxl import load_workbook
from pprint import pprint
from pathlib import Path
import requests
import csv

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

def extract_data(filename, extract_extra_data=True):
    wb = load_workbook(filename = filename)
    Path('./data').mkdir(parents=True, exist_ok=True)
    extract_csv_files(wb)
    if (extract_extra_data):
        extract_extra_information(wb)
    wb.close()

def main():
    print("Long Term Productivity data converter")
    print("Downloading excel database from 'http://longtermproductivity.com/download/BCLDatabase_online_v2.4.xlsx'")
    download_excel_database()
    extract_data('BCLDatabase_online.xlsx', extract_extra_data=False)

if __name__ == '__main__':
    main()