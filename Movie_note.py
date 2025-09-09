import streamlit as st
import requests
import os
from dotenv import load_dotenv
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as XLImage
from io import BytesIO
from openpyxl.styles import Alignment, Font, PatternFill

# .env から API_KEY を読み込み
load_dotenv()
API_KEY = os.getenv("MOVIE_API_KEY")
BASE_URL = "https://api.themoviedb.org/3"
EXCEL_FILE = "movie_note.xlsx"

# =========================================================
# Google Drive サービス取得
# =========================================================

def get_gdrive_service():

    """Google Drive API サービスを返す（OAuth 方式）"""
    creds = None

    # Render 環境: TOKEN_B64 を優先
    if "TOKEN_B64" in os.environ:
        token_bytes = base64.b64decode(os.environ["TOKEN_B64"])
        creds = pickle.loads(token_bytes)

    # ローカル環境: token.pickle を利用
    elif os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)

    # 期限切れならリフレッシュ
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())

    if not creds:
        raise FileNotFoundError("token.pickle が見つかりません。")

    return build("drive", "v3", credentials=creds)

# =========================================================
# Google Drive 上書き保存関数
# =========================================================

def upload_to_drive(excel_data, folder_id, filename="movie_note.xlsx"):
    service = get_gdrive_service()

    # 既存ファイルがあるか検索
    query = f"'{folder_id}' in parents and name='{filename}' and trashed=false"
    results = service.files().list(q=query, fields="files(id)").execute()
    items = results.get("files", [])

    media = MediaIoBaseUpload(
        io.BytesIO(excel_data),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        resumable=False
    )

    if items:
        # 更新（上書き）
        file_id = items[0]["id"]
        updated_file = service.files().update(
            fileId=file_id,
            media_body=media,
            fields="id, name, modifiedTime, version"
        ).execute()
        return updated_file["id"], updated_file["modifiedTime"], updated_file.get("version")
    else:
        # 新規作成
        file_metadata = {"name": filename, "parents": [folder_id]}
        new_file = service.files().create(
            body=file_metadata,
            media_body=media,
            fields="id, name, modifiedTime, version"
        ).execute()
        return new_file["id"], new_file["modifiedTime"], new_file.get("version")

# =========================================================
# Driveからファイルをダウンロード
# =========================================================

def download_from_drive(folder_id, filename="movie_note.xlsx"):
    service = get_gdrive_service()

    # Drive上にファイルがあるか検索
    query = f"'{folder_id}' in parents and name='{filename}' and trashed=false"
    results = service.files().list(q=query, fields="files(id)").execute()
    items = results.get("files", [])

    if not items:
        return None  # ファイルがまだ存在しない

    file_id = items[0]["id"]
    request = service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()
    fh.seek(0)    
    return fh.getvalue()

# =========================================================
# 見栄えを整える（列幅・行高さ・セル配置など）
# =========================================================

def format_excel(ws):

    # 列幅設定
    col_widths = {
        "A": 20, "B": 10, "C": 15, "D": 20,
        "E": 40, "F": 40, "G": 40
    }
    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width

    # 行の高さ：2行目以降はすべて120
    for row in range(2, ws.max_row + 1):
        ws.row_dimensions[row].height = 120

    # A〜H列：縦位置 上詰め
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=7):
        for cell in row:
            cell.alignment = Alignment(vertical="top")

    # D, E列：折り返して表示
    for col in ["D", "E"]:
        for row in range(2, ws.max_row + 1):
            ws[f"{col}{row}"].alignment = Alignment(vertical="top", wrap_text=True)

    # --- ヘッダー行の装飾（1行目） ---
    header_fill = PatternFill(start_color="87CEEB", end_color="87CEEB", fill_type="solid")  # スカイブルー
    header_font = Font(bold=True)

    for cell in ws[1]:  # 1行目の全セル
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.font = header_font
        cell.fill = header_fill

    return ws

def search_movies(query):
    url = f"{BASE_URL}/search/movie"
    params = {"api_key": API_KEY, "query": query, "language": "ja-JP"}
    res = requests.get(url, params=params)
    return res.json().get("results", [])

def get_movie_details(movie_id):
    url = f"{BASE_URL}/movie/{movie_id}"
    params = {"api_key": API_KEY, "language": "ja-JP", "append_to_response": "credits"}
    res = requests.get(url, params=params)
    return res.json()

def save_to_excel(movie_data, poster_url):
    # Excelファイルがあるか確認
    if os.path.exists(EXCEL_FILE):
        wb = load_workbook(EXCEL_FILE)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.append(["タイトル", "公開年", "監督", "出演者", "概要", "感想", "ポスター"])   # ヘッダー

    # 次の行番号
    row = ws.max_row + 1

    # 文字情報を追加
    ws.cell(row=row, column=1, value=movie_data["タイトル"])
    ws.cell(row=row, column=2, value=movie_data["公開年"])
    ws.cell(row=row, column=3, value=movie_data["監督"])
    ws.cell(row=row, column=4, value=movie_data["出演者"])
    ws.cell(row=row, column=5, value=movie_data["概要"])
    ws.cell(row=row, column=6, value=movie_data["感想"])

    # ポスター画像をダウンロードして貼り付け
    if poster_url:
        img_data = requests.get(poster_url).content
        img = XLImage(BytesIO(img_data))
        img.width, img.height = 80, 120  # サムネイルサイズ
        ws.add_image(img, f"G{row}")  # F列に配置

    format_excel(ws)

    wb.save(EXCEL_FILE)

st.title("🎬 映画検索アプリ")

query = st.text_input("映画タイトルを入力してください")
if query:
    results = search_movies(query)

    if results:
        titles = [f"{m['title']} ({m.get('release_date','')[:4]})" for m in results]
        choice = st.radio("検索結果から選択してください:", titles)

        selected = results[titles.index(choice)]
        details = get_movie_details(selected["id"])

        # ポスターURL
        poster_url = None
        if selected.get("poster_path"):
            poster_url = f"https://image.tmdb.org/t/p/w200{selected['poster_path']}"
            st.image(poster_url)

        # 監督
        director = [c["name"] for c in details["credits"]["crew"] if c["job"] == "Director"]
        director = director[0] if director else "不明"

        # 俳優（上位3人）
        cast = [c["name"] for c in details["credits"]["cast"][:3]]

        st.write(f"**監督**: {director}")
        st.write(f"**出演者**: {', '.join(cast)}")

        # 感想入力エリア
        comment = st.text_area("感想を入力してください")

        # if st.button("Excelに保存"):
        #     movie_data = {
        #         "タイトル": details["title"],
        #         "公開年": details.get("release_date", "")[:4],
        #         "監督": director,
        #         "出演者": ", ".join(cast),
        #         "概要": details.get("overview", ""),
        #         "感想": comment
        #     }
        #     save_to_excel(movie_data, poster_url)
        #     st.success("Excelに保存しました！（サムネイル付き）")

        # ✅ Streamlit Google Drive保存ボタン 
        if st.button("📤 Google Driveに保存（上書き）"):

            folder_id = "1UNBH5iMlZGyWYEXGZZOfqog2DqS1MkpQ"

            # 1. Driveから既存Excelをダウンロード
            existing_bytes = download_from_drive(folder_id, "movie_note.xlsx")

            # 既存状況のデバッグ表示
            if existing_bytes:
                wb_tmp = load_workbook(filename=BytesIO(existing_bytes))  # ← file= を使う
                st.info(f"DEBUG: 今の最終行（保存前）: {wb_tmp.active.max_row}")
            else:
                st.info("DEBUG: 既存ファイルなし（新規作成）")

            # 2. 行を追加
            # st.write("DEBUG 選択書籍:", selected_book.get('title'))
            # st.write("DEBUG 感想:", comment)
            # excel_data = create_excel_with_image(selected_book, comment, base_xlsx_bytes=existing_bytes)

            # 3. Drive へ保存（結果も確認表示）
            file_id, modified, version = upload_to_drive(excel_data, folder_id, filename="book_note.xlsx")
            st.success(f"✅ Google Driveに保存しました！\nID: {file_id}\n更新時刻: {modified}\n版: {version}")
            st.caption(f"https://drive.google.com/file/d/{file_id}/view")

    else:
        st.warning("検索結果が見つかりませんでした。")




