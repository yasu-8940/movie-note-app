import streamlit as st
import requests
import os
from dotenv import load_dotenv

# .env から API_KEY を読み込み
load_dotenv()
API_KEY = os.getenv("MOVIE_API_KEY")
BASE_URL = "https://api.themoviedb.org/3"

st.title("🎬 映画検索アプリ（Ver0）")

load_dotenv()
print("DEBUG: API_KEY =", os.getenv("MOVIE_API_KEY"))

# 入力フォーム
movie_title = st.text_input("映画タイトルを入力してください")

if st.button("検索"):
    if not movie_title:
        st.warning("タイトルを入力してください")
    else:
        # TMDb API で映画検索
        url = f"{BASE_URL}/search/movie"
        params = {"api_key": API_KEY, "query": movie_title, "language": "ja"}
        response = requests.get(url, params=params)

        if response.status_code == 200:
            data = response.json()
            results = data.get("results", [])

            if not results:
                st.info("映画が見つかりませんでした。")
            else:
                # 上位5件までタイトルリストを作る
                options = [f"{m.get('title')} ({m.get('release_date', '不明')})"
                           for m in results[:5]]
                selected = st.radio("検索結果から選んでください:", options)

                # 選択された映画を表示
                if selected:
                    idx = options.index(selected)
                    movie = results[idx]
                    st.subheader(movie.get("title"))
                    st.write("公開日:", movie.get("release_date", "不明"))
                    st.write("概要:", movie.get("overview", "なし"))
        else:
            st.error(f"エラーが発生しました: {response.status_code}")