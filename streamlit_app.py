# streamlit_app.py — CSV(저장 데이터) 기반 조회 버전
import os
import pandas as pd
import streamlit as st
from datetime import datetime, timedelta

CSV_PATH = "data/daily_weather.csv"   # fetch_weather.py가 저장하는 경로

KOR_COLS = {
    "city_key": "도시키",
    "city_name_kr": "도시명",
    "date": "날짜",
    "tmax_c": "최고기온(°C)",
    "tmin_c": "최저기온(°C)",
    "precip_mm": "강수량 합계(mm)",
    "precip_prob_max_pct": "강수확률 최대(%)",
    "weathercode": "WMO코드",
    "weather_desc": "날씨 설명",
}

st.set_page_config(page_title="Daily Weather (CSV Viewer)", page_icon="🌤️", layout="wide")
st.title("🌤️ 일일 날씨 — 저장 데이터 조회")

# 0) 파일 존재 여부 체크
if not os.path.exists(CSV_PATH):
    st.error(f"CSV 파일을 찾을 수 없어요: `{CSV_PATH}`")
    st.info(
        "먼저 GitHub Actions 또는 fetch_weather.py를 실행해 데이터를 저장하세요.\n"
        "기본 저장 파일: data/daily_weather.csv"
    )
    st.stop()

# 1) CSV 로드
@st.cache_data(ttl=300)
def load_csv(path: str) -> pd.DataFrame:
    df = pd.read_csv(path, encoding="utf-8")
    # 날짜 파싱 & 정렬
    if "date" in df.columns:
        df["date"] = pd.to_datetime(df["date"], errors="coerce")
    # 컬럼 존재 보정(오타/누락 대비)
    for col in ["city_key","city_name_kr","tmax_c","tmin_c","precip_mm","precip_prob_max_pct","weathercode","weather_desc"]:
        if col not in df.columns:
            df[col] = pd.NA
    return df

df = load_csv(CSV_PATH)
if df.empty or df["date"].isna().all():
    st.warning("CSV에 유효한 데이터가 없어요. 저장 스크립트를 먼저 실행해 주세요.")
    st.stop()

# 2) 사이드바 — 필터
with st.sidebar:
    st.header("🔎 필터")
    # 도시 목록: (라벨: 한글명, 값: city_key)로 구성
    # 같은 도시키가 여러 번 있을 수 있으므로 최신 라벨 우선
    latest_city_names = (
        df.sort_values("date")
          .groupby("city_key")["city_name_kr"]
          .last()
          .fillna(df["city_key"])
    )
    city_options = {f"{latest_city_names[k]} ({k})": k for k in latest_city_names.index}

    sel_cities = st.multiselect(
        "도시 선택",
        options=list(city_options.keys()),
        default=list(city_options.keys())[: min(3, len(city_options))],
    )
    sel_city_keys = [city_options[label] for label in sel_cities] if sel_cities else []

    # 날짜 범위 (기본: 최근 14일)
    min_date = pd.to_datetime(df["date"].min()).date()
    max_date = pd.to_datetime(df["date"].max()).date()
    default_start = max(min_date, (max_date - timedelta(days=13)))
    start_date, end_date = st.date_input(
        "날짜 범위",
        value=(default_start, max_date),
        min_value=min_date,
        max_value=max_date,
        format="YYYY-MM-DD",
    )

# 3) 필터 적용
mask = (df["date"].dt.date >= start_date) & (df["date"].dt.date <= end_date)
if sel_city_keys:
    mask &= df["city_key"].isin(sel_city_keys)

view = df.loc[mask].copy().sort_values(["date","city_key"], ascending=[False, True])

if view.empty:
    st.warning("선택한 조건에 해당하는 데이터가 없습니다. 날짜 범위나 도시를 바꿔보세요.")
    st.stop()

# 4) 상단 요약 카드 (가장 최근 날짜 기준)
latest_day = view["date"].max()
today_slice = view[view["date"] == latest_day]

st.subheader(f"📅 가장 최근 데이터: {latest_day.date()}")

# 선택된 도시들에 대해 간단 요약(평균값)
col1, col2, col3, col4 = st.columns(4)
col1.metric("도시 수", today_slice["city_key"].nunique())
col2.metric("평균 최고기온(°C)", f"{today_slice['tmax_c'].astype(float).mean():.1f}")
col3.metric("평균 최저기온(°C)", f"{today_slice['tmin_c'].astype(float).mean():.1f}")
col4.metric("평균 강수량(mm)", f"{today_slice['precip_mm'].astype(float).mean():.1f}")

st.divider()

# 5) 상세 표 (날짜/도시별)
# 한국어 컬럼명으로 변환된 표 제공
view_for_table = view[
    ["city_name_kr","city_key","date","tmax_c","tmin_c","precip_mm","precip_prob_max_pct","weathercode","weather_desc"]
].rename(columns=KOR_COLS)

st.subheader("📊 상세 데이터")
st.dataframe(
    view_for_table,
    use_container_width=True,
    hide_index=True
)

# 6) (선택) 간단한 추이 차트: 최고/최저/강수량
st.subheader("📈 추이 보기 (선택 도시 합산/평균)")
agg = (
    view.groupby("date")
        .agg(
            최고기온=("tmax_c", "mean"),
            최저기온=("tmin_c", "mean"),
            강수량mm=("precip_mm", "mean"),
        )
        .sort_index()
)
tab1, tab2, tab3 = st.tabs(["최고기온", "최저기온", "강수량"])

with tab1:
    st.line_chart(agg["최고기온"], height=260)
with tab2:
    st.line_chart(agg["최저기온"], height=260)
with tab3:
    st.line_chart(agg["강수량mm"], height=260)

st.caption("※ 차트는 선택한 도시들의 평균값 기준입니다. 수치는 저장된 CSV를 그대로 사용합니다.")
