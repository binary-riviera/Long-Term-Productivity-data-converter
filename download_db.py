import requests

def download_excel_database():
    url = 'http://longtermproductivity.com/download/BCLDatabase_online_v2.4.xlsx'
    r = requests.get(url, allow_redirects=False)
    open('BCLDatabase_online_v2.4.xlsx', 'wb').write(r.content)