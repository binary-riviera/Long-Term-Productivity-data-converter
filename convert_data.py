from openpyxl import load_workbook
from pathlib import Path
import requests
import csv
import json
import re

BASE_URL = 'http://longtermproductivity.com'
DOCUMENT_NAME = "BCLDatabase_online"

def get_download_url(base_url, document):
    # I'm not using beautiful soup here to minimise the number of dependencies
    r = requests.get(base_url + '/download.html', allow_redirects=True)
    html = r.text
    links = re.findall('href=[\"\'](.*?)[\"\']', html) # yes, parsing html with regular expressions is bad
    download_url = base_url + '/' + next(link for link in links if document in link)
    return download_url

def download_excel_database(url, filename):
    r = requests.get(url, allow_redirects=False)
    open(filename, 'wb').write(r.content)


def extract_csv_files(workbook):
    print('Extracting core data into csv files...')
    sheets_to_export = ['GDP per capita', 'Labor Productivity', 'TFP', 'KI', 'AgeK']
    for s in sheets_to_export:
        sheet = workbook[s]
        values = sheet.values
        with open('data/' + s + '.csv', 'w', newline='') as csv_file:
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
    with open('data/extra_information.json', 'w') as json_file:
        json.dump(extra_information, json_file, ensure_ascii=False, indent=4)


def extract_data(filename, extract_extra_data):
    wb = load_workbook(filename = filename, data_only=True)
    Path('./data').mkdir(parents=True, exist_ok=True)
    extract_csv_files(wb)
    if (extract_extra_data):
        extract_extra_information(wb)
    wb.close()


def main():
    print("Long Term Productivity data converter")
    print("Downloading excel database from '${BASE_URL}'")
    url = get_download_url(BASE_URL, DOCUMENT_NAME)
    download_excel_database(url, DOCUMENT_NAME + '.xlsx')
    extract_data(DOCUMENT_NAME + '.xlsx', extract_extra_data=True)

if __name__ == '__main__':
    main()