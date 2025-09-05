import requests

# TMDb APIキーをここに入力
API_KEY = "YOUR_API_KEY"

# 検索したい映画タイトル
query = "七人の侍"

# 映画検索（日本語対応）
search_url = f"https://api.themoviedb.org/3/search/movie"
params = {
    "api_key": API_KEY,
    "query": query,
    "language": "ja"  # 日本語データを優先
}

response = requests.get(search_url, params=params)
results = response.json().get("results", [])

if not results:
    print("映画が見つかりませんでした")
else:
    movie = results[0]  # 一番上の検索結果を使用
    movie_id = movie["id"]
    title = movie["title"]
    release_date = movie.get("release_date", "不明")
    overview = movie.get("overview", "あらすじなし")
    poster_path = movie.get("poster_path", None)
    poster_url = f"https://image.tmdb.org/t/p/w500{poster_path}" if poster_path else "なし"

    # 監督や出演者の取得
    credits_url = f"https://api.themoviedb.org/3/movie/{movie_id}/credits"
    credits_params = {
        "api_key": API_KEY,
        "language": "ja"
    }
    credits_response = requests.get(credits_url, params=credits_params).json()
    
    # 監督を探す
    crew = credits_response.get("crew", [])
    director = next((c["name"] for c in crew if c["job"] == "Director"), "不明")
    
    # 主要キャスト（先頭3人）
    cast = credits_response.get("cast", [])
    main_cast = [c["name"] for c in cast[:3]]

    print("🎬 タイトル:", title)
    print("📅 公開日:", release_date)
    print("🎞 監督:", director)
    print("⭐ 主演:", ", ".join(main_cast))
    print("📝 あらすじ:", overview)
    print("🖼 ポスターURL:", poster_url)