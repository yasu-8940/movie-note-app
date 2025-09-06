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

        if st.button("Excelに保存"):
            movie_data = {
                "タイトル": details["title"],
                "公開年": details.get("release_date", "")[:4],
                "監督": director,
                "出演者": ", ".join(cast),
                "概要": details.get("overview", ""),
                "感想": comment
            }
            save_to_excel(movie_data, poster_url)
            st.success("Excelに保存しました！（サムネイル付き）")
    else:
        st.warning("検索結果が見つかりませんでした。")