import requests

def download_excel_database():
    url = 'http://longtermproductivity.com/download/BCLDatabase_online_v2.4.xlsx'
    r = requests.get(url, allow_redirects=False)
    open('BCLDatabase_online_v2.4.xlsx', 'wb').write(r.content)



def main():
    print("Long Term Productivity data converter v0.0.1")
    print("Downloading excel database from 'http://longtermproductivity.com/download/BCLDatabase_online_v2.4.xlsx'")
    download_excel_database()

if __name__ == '__main__':
    main()