from __future__ import print_function
import os.path
import pickle
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

SCOPES = ['https://www.googleapis.com/auth/drive.file']

def main():
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('drive', 'v3', credentials=creds)

    # ✅ まずファイル一覧を表示
    results = service.files().list(pageSize=10, fields="files(id, name)").execute()
    items = results.get('files', [])
    if not items:
        print('Google Drive 上にまだファイルがありません。')
    else:
        print('Google Drive 上のファイル一覧:')
        for item in items:
            print(u'{0} ({1})'.format(item['name'], item['id']))

    # ✅ 次に book_note.xlsx をアップロード
    file_path = r"C:\Users\seki8\OneDrive\デスクトップ\python_lesson\book_note.xlsx"
    print("デバッグ: file_path =", file_path)
    print("デバッグ: exists? ->", os.path.exists(file_path))
    if os.path.exists(file_path):
        file_metadata = {'name': os.path.basename(file_path)}
        media = MediaFileUpload(file_path, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
        print(f"✅ アップロード成功！ File ID: {file.get('id')}")
    else:
        print(f"⚠️ ローカルに {file_path} が存在しません")

if __name__ == '__main__':
    main()