# fetch_weather.py
import argparse
import csv
import os
from datetime import datetime, date
from dateutil import tz
import requests

# 기본 도시 좌표 (원하면 자유롭게 추가/수정)
CITIES = {
    "seoul":    {"name_kr": "서울",    "lat": 37.5665, "lon": 126.9780},
    "suwon":    {"name_kr": "수원",    "lat": 37.2636, "lon": 127.0286},
    "dublin":   {"name_kr": "더블린(아일랜드)",    "lat": 53.3498, "lon": -6.2603},
    "rome":  {"name_kr": "로마(이탈리아)",    "lat": 53.3498, "lon": -6.2603},
    "cusco":  {"name_kr": "쿠스코(페루)",    "lat": -13.1631, "lon": -72.5450},
}

WEATHERCODE_DESC = {
    0: "맑음", 1: "대체로 맑음", 2: "부분적으로 흐림", 3: "흐림",
    45: "안개", 48: "서리안개",
    51: "약한 이슬비", 53: "보통 이슬비", 55: "강한 이슬비",
    56: "약한 언 이슬비", 57: "강한 언 이슬비",
    61: "약한 비", 63: "보통 비", 65: "강한 비",
    66: "약한 언 비", 67: "강한 언 비",
    71: "약한 눈", 73: "보통 눈", 75: "강한 눈",
    77: "싸락눈",
    80: "약한 소나기", 81: "보통 소나기", 82: "강한 소나기",
    85: "약한 눈 소나기", 86: "강한 눈 소나기",
    95: "천둥번개", 96: "우박을 동반한 천둥번개(약~보통)", 99: "우박을 동반한 천둥번개(강함)"
}

def get_today_iso_kst():
    """KST 기준 YYYY-MM-DD 문자열"""
    kst = tz.gettz("Asia/Seoul")
    return datetime.now(tz=kst).date().isoformat()

def fetch_daily_weather(lat: float, lon: float, day: str):
    """
    Open-Meteo 일일예보 호출: 최대/최저기온, 강수량합, 강수확률최대, WMO 코드
    """
    url = (
        "https://api.open-meteo.com/v1/forecast"
        f"?latitude={lat}&longitude={lon}"
        "&daily=temperature_2m_max,temperature_2m_min,precipitation_sum,"
        "precipitation_probability_max,weathercode"
        "&timezone=Asia%2FSeoul"
    )
    r = requests.get(url, timeout=20)
    r.raise_for_status()
    data = r.json()
    daily = data.get("daily", {})

    # 날짜 배열에서 오늘 인덱스 찾기
    times = daily.get("time", [])
    if day not in times:
        raise ValueError(f"API 응답에 {day} 날짜가 없음. 응답 날짜들: {times}")

    i = times.index(day)
    def pick(key, default=None):
        arr = daily.get(key, [])
        return arr[i] if i < len(arr) else default

    return {
        "date": day,
        "tmax": pick("temperature_2m_max"),
        "tmin": pick("temperature_2m_min"),
        "precip_sum": pick("precipitation_sum"),
        "precip_prob_max": pick("precipitation_probability_max"),
        "weathercode": pick("weathercode"),
    }

def ensure_paths():
    os.makedirs("data", exist_ok=True)
    os.makedirs("logs", exist_ok=True)

def append_csv(row: dict, city_key: str, city_name_kr: str):
    path = os.path.join("data", "daily_weather.csv")
    file_exists = os.path.exists(path)
    with open(path, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow([
                "city_key","city_name_kr","date",
                "tmax_c","tmin_c","precip_mm","precip_prob_max_pct",
                "weathercode","weather_desc"
            ])
        desc = WEATHERCODE_DESC.get(int(row["weathercode"]) if row["weathercode"] is not None else -1, "")
        writer.writerow([
            city_key, city_name_kr, row["date"],
            row["tmax"], row["tmin"], row["precip_sum"], row["precip_prob_max"],
            row["weathercode"], desc
        ])

def main():
    parser = argparse.ArgumentParser(description="매일 새벽 6시 날씨 수집")
    parser.add_argument("--city", type=str, default="seoul", help="seoul/suwon/dublin/rome/cusco 등 키")
    args = parser.parse_args()

    ensure_paths()

    city_key = args.city.lower()
    if city_key not in CITIES:
        raise SystemExit(f"알 수 없는 도시 키: {city_key}. 사용 가능: {', '.join(CITIES.keys())}")

    meta = CITIES[city_key]
    today = get_today_iso_kst()
    try:
        row = fetch_daily_weather(meta["lat"], meta["lon"], today)
        append_csv(row, city_key, meta["name_kr"])
        # 콘솔 출력(로그로도 남음)
        desc = WEATHERCODE_DESC.get(int(row["weathercode"]) if row["weathercode"] is not None else -1, "")
        print(f"[{today}] {meta['name_kr']} 수집 완료 → "
              f"최고 {row['tmax']}°C / 최저 {row['tmin']}°C, 강수 {row['precip_sum']}mm,"
              f" 강수확률최대 {row['precip_prob_max']}%, {desc}")
    except Exception as e:
        print(f"수집 실패: {e}")

if __name__ == "__main__":
    main()
