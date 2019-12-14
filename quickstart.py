from httplib2 import Http
from googleapiclient.discovery import build
from oauth2client.service_account import ServiceAccountCredentials
# Google API 요청 시 필요한 권한 유형
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
# 구글 시트 ID
SPREADSHEET_ID = '1dpM0Ai6w2Iz4EnaCzPlesigU4wcouIQfJRufwICTqwU'
# json 파일로 서비스 계정 credential 정의
credentials = ServiceAccountCredentials.from_json_keyfile_name('C:/Users/Hyun/Downloads/quickstart-1575028783267-e76b12a95dfe.json', SCOPES)
http_auth = credentials.authorize(Http())
service = build('sheets', 'v4', http=http_auth)
# 업데이트 요청 및 실행

def main():
    vlist = [
    ['이건', '첫 번째', '행입니다.','ㄷㄷㄷ','ㅇㄹㅇㄹ'],
    ['6 번째','ㄷㄹㄷㄹㄷ','','','ㅇㄹㅇㅇㅇㅇ'],
    ['열입니다.','ㅇㄹㅇㅇㅇ'],
    ]
    body = {
    'values': vlist
    }
    
    subject = '국어'
    var_range = f'{subject}!A:H'
    request = service.spreadsheets().values().update(spreadsheetId=SPREADSHEET_ID,
    range=var_range, # 2
    valueInputOption='RAW',
    body=body)
    try:
        request.execute()
        input('업데이트 완료')
    except Exception as e:
        input(f'에러 발생 : {e}')

if __name__ == '__main__':
    main()
