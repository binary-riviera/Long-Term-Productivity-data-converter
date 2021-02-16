from download_db import download_excel_database




def main():
    print("Long Term Productivity data converter v0.0.1")
    print("Downloading excel database from 'http://longtermproductivity.com/download/BCLDatabase_online_v2.4.xlsx'")
    download_excel_database()

if __name__ == '__main__':
    main()