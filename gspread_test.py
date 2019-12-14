import os
import gspread
import openpyxl
import datetime
from oauth2client.service_account import ServiceAccountCredentials

# 구글 문서 연결
def conn_gspred(SHEET_URL, month_no):
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    credentials = ServiceAccountCredentials.from_json_keyfile_name('C:/Users/Hyun/Downloads/quickstart-1575028783267-e76b12a95dfe.json', SCOPES)
    gc = gspread.authorize(credentials)
    sh1 = gc.open_by_url(SHEET_URL)
    sheet_title = f'{month_no}월'
    ws = sh1.worksheet(sheet_title)
    cell_list = ws.range('B3:M6')
    print(len(cell_list))


if __name__ == '__main__':
    #URL_file = open("./eng_url.txt", 'rt', encoding='UTF8')
    #SHEET_URL = URL_file.read()
    #URL_file.close()
    SHEET_URL = 'https://docs.google.com/spreadsheets/d/1kZO8bNqG-DrzYTdVNFAnpIJroHr44_wQJg0n-zEq_WA/edit#gid=0'
    #print(f"eng_url.txt 파일에서 읽어온 URL : {SHEET_URL}\n")
    #month_no = input("몇 월(예 : 12)? ")
    #week_no = input("몇 주차(예 : 2)? ")
    #print(f"{month_no}월 {week_no}주차 주간 문자 생성 중.")
    #list_all = conn_gspred(SHEET_URL, month_no)
    #msg_data = write_message(list_all, month_no, week_no)
    month_no = '12'
    conn_gspred(SHEET_URL, month_no)