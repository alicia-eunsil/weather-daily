# streamlit_app.py
import streamlit as st
import requests
from datetime import datetime
from dateutil import tz

CITIES = {
    "서울(Seoul)":   {"coords": (37.5665, 126.9780), "tz": "Asia/Seoul"},
    "더블린(Dublin)": {"coords": (53.3498, -6.2603), "tz": "Europe/Dublin"},
    "로마(Rome)":     {"coords": (41.9028, 12.4964), "tz": "Europe/Rome"},
    "마추피추(machu_picchu)":  {"coords": (-13.1631, -72.5450), "tz": "America/Lima"},
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
    95: "천둥번개", 96: "우박 동반 번개(약~보통)", 99: "우박 동반 번개(강함)"
}

st.set_page_config(page_title="Daily Weather", page_icon="🌤️")
st.title("🌤️ 오늘의 날씨")

city_label = st.selectbox("도시 선택", list(CITIES.keys()), index=0)
lat, lon = CITIES[city_label]["coords"]
city_tz = tz.gettz(CITIES[city_label]["tz"])

# Local Time (선택 도시 현지시간)
local_time = datetime.now(tz=city_tz)
local_date = local_time.date().isoformat()
local_time_str = local_time.strftime("%Y-%m-%d %H:%M:%S")

# KST Time
kst = tz.gettz("Asia/Seoul")
kst_time = datetime.now(tz=kst).strftime("%Y-%m-%d %H:%M:%S")

st.write(f"**📍 도시 현지 날짜/시간:** {local_time_str} ({CITIES[city_label]['tz']})")
st.write(f"**🇰🇷 한국 시간(KST):** {kst_time}")

# API 호출은 KST 날짜 기준 유지
today = datetime.now(tz=kst).date().isoformat()

url = (
    "https://api.open-meteo.com/v1/forecast"
    f"?latitude={lat}&longitude={lon}"
    "&daily=temperature_2m_max,temperature_2m_min,precipitation_sum,"
    "precipitation_probability_max,weathercode"
    "&timezone=Asia%2FSeoul"
)

try:
    r = requests.get(url, timeout=20)
    r.raise_for_status()
    daily = r.json().get("daily", {})
    times = daily.get("time", [])
    if today in times:
        i = times.index(today)
        tmax = daily.get("temperature_2m_max", [None])[i]
        tmin = daily.get("temperature_2m_min", [None])[i]
        psum = daily.get("precipitation_sum", [None])[i]
        pprob = daily.get("precipitation_probability_max", [None])[i]
        wcode = daily.get("weathercode", [None])[i]
        desc = WEATHERCODE_DESC.get(int(wcode) if wcode is not None else -1, "")

        st.subheader(f"{city_label} • {today} (KST 기준 데이터)")
        st.metric("최고기온(°C)", tmax)
        st.metric("최저기온(°C)", tmin)
        st.metric("강수량 합계(mm)", psum)
        st.metric("강수확률 최대(%)", pprob)
        st.write(f"**날씨 설명:** {desc} (코드 {wcode})")
    else:
        st.error(f"API 응답에 오늘({today})이 없습니다. 응답 날짜: {times}")
except Exception as e:
    st.error(f"조회 실패: {e}")

