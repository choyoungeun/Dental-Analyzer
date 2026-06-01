import streamlit as st
import pandas as pd
import requests
from geopy.geocoders import Nominatim
import folium
from streamlit_folium import st_folium
from datetime import datetime
import re
import altair as alt
import math
import numpy as np
from io import StringIO
import os
import glob
from difflib import SequenceMatcher

# ==========================================
# 0. API 키 세팅
# ==========================================
# 실제 서비스에서는 코드에 API 키를 직접 넣지 말고 st.secrets 사용 권장
# Streamlit Cloud의 Secrets에서 관리
# Secrets 예시: My_KEY = "발급받은_API_KEY"
try:
    MY_API_KEY = st.secrets[ "MY_API_KEY"]
except Exception:
    MY_API_KEY = ""

# 그래프 높이를 한 곳에서 통일 관리
CHART_HEIGHT = 260


# ==========================================
# 0-1. 세션 기본 좌표
# ==========================================
if "target_lat" not in st.session_state:
    st.session_state["target_lat"] = 37.5665  # 서울시청 위도

if "target_lon" not in st.session_state:
    st.session_state["target_lon"] = 126.9780  # 서울시청 경도


# ==========================================
# 0-2. hosp_data.csv 로드 함수
# ==========================================
@st.cache(suppress_st_warning=True)
def load_hosp_data():
    file_path = "hosp_data.xlsx"

    try:
        df = pd.read_excel(file_path, engine="openpyxl")
    except Exception as e:
        return pd.DataFrame(), f"hosp_data.xlsx 파일을 읽지 못했습니다: {e}"

    # 컬럼명 공백 제거
    df.columns = df.columns.astype(str).str.strip()

    # 필수 컬럼 확인
    required_cols = ["요양기관명", "개설일자", "좌표(X)", "좌표(Y)"]
    missing_cols = [col for col in required_cols if col not in df.columns]

    if missing_cols:
        msg = (
            f"엑셀 파일에 필요한 컬럼이 없습니다: {missing_cols}\n"
            f"현재 인식된 컬럼명: {df.columns.tolist()}"
        )
        return pd.DataFrame(), msg

    # 좌표 정리
    # 좌표(X) = 경도, 좌표(Y) = 위도
    df["경도"] = pd.to_numeric(df["좌표(X)"], errors="coerce")
    df["위도"] = pd.to_numeric(df["좌표(Y)"], errors="coerce")

    # 상호명
    df["상호명"] = df["요양기관명"].astype(str)

    # 개업일자 정리
    df["개업일_raw"] = (
        df["개설일자"]
        .astype(str)
        .str.replace("-", "", regex=False)
        .str.replace(".", "", regex=False)
        .str.replace("/", "", regex=False)
        .str.strip()
    )

    df["개업일_raw"] = df["개업일_raw"].str.extract(r"(\d{4,8})")[0]

    # 치과만 필터링
    df = df[df["상호명"].str.contains("치과", na=False)].copy()

    # 좌표 없는 행 제거
    df = df.dropna(subset=["위도", "경도"])

    return df, None

# ==========================================
# 1. 수학적 거리 계산 함수
# ==========================================
def calculate_distance(lat1, lon1, lat2, lon2):
    R = 6371.0

    dlat = math.radians(lat2 - lat1)
    dlon = math.radians(lon2 - lon1)

    a = (
        math.sin(dlat / 2) ** 2
        + math.cos(math.radians(lat1))
        * math.cos(math.radians(lat2))
        * math.sin(dlon / 2) ** 2
    )

    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))

    distance_km = R * c
    return distance_km * 1000

def normalize_clinic_name(name):
    """
    치과명 비교를 위해 불필요한 단어와 특수문자를 제거
    """
    name = str(name)

    # 공백 제거
    name = re.sub(r"\s+", "", name)

    # 괄호 내용 제거
    name = re.sub(r"\(.*?\)", "", name)

    # 특수문자 제거
    name = re.sub(r"[^가-힣A-Za-z0-9]", "", name)

    # 자주 붙는 단어 정리
    remove_words = [
        "치과의원",
        "치과병원",
        "의료법인",
        "의원",
        "병원",
        "치과"
    ]

    for word in remove_words:
        name = name.replace(word, "")

    return name.lower()

def name_similarity(name1, name2):
    n1 = normalize_clinic_name(name1)
    n2 = normalize_clinic_name(name2)

    if not n1 or not n2:
        return 0

    return SequenceMatcher(None, n1, n2).ratio()

def remove_duplicate_clinics(df, distance_threshold=10, similarity_threshold=0.9):
    """
    위치가 10m 이내이고 이름 유사도가 90% 이상이면 동일 치과로 간주하여 중복 제거
    개업일 정보가 있는 행을 우선 보존
    """

    if df.empty:
        return df

    df = df.copy().reset_index(drop=True)

    # 필수 컬럼 없으면 그대로 반환
    required_cols = ["상호명", "위도", "경도"]
    for col in required_cols:
        if col not in df.columns:
            return df

    keep_indices = []
    removed_indices = set()

    for i in range(len(df)):
        if i in removed_indices:
            continue

        base = df.loc[i]
        duplicate_group = [i]

        for j in range(i + 1, len(df)):
            if j in removed_indices:
                continue

            target = df.loc[j]

            try:
                dist = calculate_distance(
                    base["위도"],
                    base["경도"],
                    target["위도"],
                    target["경도"]
                )
            except:
                continue

            sim = name_similarity(
                base["상호명"],
                target["상호명"]
            )

            if dist <= distance_threshold and sim >= similarity_threshold:
                duplicate_group.append(j)
                removed_indices.add(j)

        # 중복 그룹 중 어떤 행을 남길지 결정
        group_df = df.loc[duplicate_group].copy()

        # 개업일 정보가 있는 행 우선
        if "개업일" in group_df.columns:
            valid_open_date = group_df[
                group_df["개업일"].astype(str).str.contains(r"\d{4}", regex=True, na=False)
                & (group_df["개업일"].astype(str) != "정보없음")
            ]

            if not valid_open_date.empty:
                chosen_index = valid_open_date.index[0]
            else:
                chosen_index = duplicate_group[0]
        else:
            chosen_index = duplicate_group[0]

        keep_indices.append(chosen_index)

    result_df = df.loc[keep_indices].copy().reset_index(drop=True)

    return result_df

# ==========================================
# 2. hosp_data.csv 기반 개업 정보 수집 함수
# ==========================================
def get_hira_opening_dates_from_csv(lat, lon, radius_m):
    df_hira, error_message = load_hosp_data()

    if error_message:
        st.error(error_message)
        return pd.DataFrame()

    df_hira = df_hira.copy()

    if df_hira.empty:
        return pd.DataFrame()

    df_hira["거리_m"] = df_hira.apply(
        lambda row: calculate_distance(
            lat,
            lon,
            row["위도"],
            row["경도"]
        ),
        axis=1
    )

    df_hira = df_hira[df_hira["거리_m"] <= radius_m].copy()

    df_hira["개업일"] = df_hira["개업일_raw"]

    return df_hira
# ==========================================
# 2-1. 개업일자 표시 포맷 함수
# ==========================================
def format_opening_date(estb_dt):
    estb_dt = str(estb_dt).strip()

    if not estb_dt or estb_dt.lower() == "nan":
        return "정보없음", 0

    estb_dt = re.sub(r"[^0-9]", "", estb_dt)

    if not estb_dt.isdigit() or len(estb_dt) < 4:
        return "정보없음", 0

    try:
        year = int(estb_dt[:4])
    except:
        return "정보없음", 0

    current_year = datetime.now().year
    years_active = current_year - year

    if len(estb_dt) >= 6:
        month = estb_dt[4:6]
        display_date = f"{year}년 {month}월"
    else:
        display_date = f"{year}년"

    return display_date, years_active


# ==========================================
# 3. 소상공인 API 기반 치과 상권 수집 함수
# ==========================================
@st.cache
def get_clinics_from_coords(lat, lon, radius_m, api_key):
    url = "http://apis.data.go.kr/B553077/api/open/sdsc2/storeListInRadius"
    all_items = []
    page_no = 1

    while True:
        params = {
            "ServiceKey": api_key,
            "pageNo": str(page_no),
            "numOfRows": "1000",
            "radius": radius_m,
            "cx": lon,
            "cy": lat,
            "type": "json"
        }

        try:
            response = requests.get(url, params=params, timeout=15)
            items = response.json().get("body", {}).get("items", [])

            if not items:
                break

            if isinstance(items, dict):
                all_items.append(items)
                break
            else:
                all_items.extend(items)

            if len(items) < 1000:
                break

            page_no += 1

            if page_no > 30:
                break

        except Exception:
            break

    if not all_items:
        return pd.DataFrame()

    df = pd.DataFrame(all_items)

    if "indsSclsNm" not in df.columns:
        return pd.DataFrame()

    dental_df = df[df["indsSclsNm"].str.contains("치과", na=False)].copy()

    if dental_df.empty:
        return pd.DataFrame()

    def extract_dong(addr):
        match = re.search(r"([가-힣A-Za-z0-9]+동)\b", str(addr))
        return match.group(1) if match else "기타"

    dental_df["동이름"] = dental_df["lnoAdr"].apply(extract_dong)

    rename_dict = {
        "bizesNm": "상호명",
        "lnoAdr": "지번주소",
        "lon": "경도",
        "lat": "위도"
    }

    existing_rename = {
        k: v for k, v in rename_dict.items()
        if k in dental_df.columns
    }

    dental_df.rename(columns=existing_rename, inplace=True)

    dental_df["위도"] = pd.to_numeric(dental_df["위도"], errors="coerce")
    dental_df["경도"] = pd.to_numeric(dental_df["경도"], errors="coerce")

    dental_df = dental_df.dropna(subset=["위도", "경도"])

    return dental_df



def _find_first_column(df, candidates):
    """후보 컬럼명 중 실제 데이터프레임에 존재하는 첫 컬럼을 반환"""
    normalized = {str(col).strip().lower(): col for col in df.columns}
    for cand in candidates:
        key = str(cand).strip().lower()
        if key in normalized:
            return normalized[key]
    return None


@st.cache(suppress_st_warning=True)
def load_local_transit_master():
    """
    대중교통 좌표 마스터 파일을 로컬에서 읽어 통합한다.
    아래 파일명 중 하나 이상을 app.py와 같은 폴더에 두면 Overpass 없이 빠르게 표시된다.
    - transit_stops.xlsx / transit_stops.csv: 버스+지하철 통합 파일
    - bus_stops.xlsx / bus_stops.csv: 버스정류장 파일
    - subway_stations.xlsx / subway_stations.csv: 지하철역 파일
    - seoul_bus_stops.xlsx / seoul_bus_stops.csv
    - seoul_subway_stations.xlsx / seoul_subway_stations.csv
    - 국토교통부_전국 버스정류장 위치정보_YYYYMMDD.csv: 공공데이터포털 15067528 버스정류장 파일
    """
    file_candidates = [
        # 통합/일반 파일명
        "transit_stops.xlsx", "transit_stops.csv",
        "bus_stops.xlsx", "bus_stops.csv",
        "subway_stations.xlsx", "subway_stations.csv",
        "seoul_bus_stops.xlsx", "seoul_bus_stops.csv",
        "seoul_subway_stations.xlsx", "seoul_subway_stations.csv",

        # 공공데이터포털 15067528: 국토교통부_전국 버스정류장 위치정보
        # 다운로드한 원본 파일명을 그대로 두거나 아래 이름 중 하나로 바꿔도 인식
        "국토교통부_전국 버스정류장 위치정보_20251031.csv",
        "국토교통부_전국 버스정류장 위치정보.csv",
        "national_bus_stops.csv",
        "bus_station_location.csv",
        "bus_stop_location.csv",

        # 공공데이터포털 15093755: 국가철도공단_도시광역철도_역사정보
        # 다운로드한 원본 파일명을 그대로 두거나 아래 이름 중 하나로 바꿔도 인식
        "국가철도공단_도시광역철도_역사정보_20241231.xlsx",
        "국가철도공단_도시광역철도_역사정보.xlsx",
        "urban_rail_station_info.xlsx",
        "rail_station_info.xlsx",
        "metro_station_info.xlsx",
    ]

    frames = []

    for file_path in file_candidates:
        if not os.path.exists(file_path):
            continue

        try:
            if file_path.lower().endswith(".xlsx"):
                raw_df = pd.read_excel(file_path, engine="openpyxl")
            else:
                # 서울/공공데이터 파일은 utf-8-sig 또는 cp949가 흔함
                try:
                    raw_df = pd.read_csv(file_path, encoding="utf-8-sig", low_memory=False)
                except Exception:
                    raw_df = pd.read_csv(file_path, encoding="cp949", low_memory=False)
        except Exception:
            continue

        if raw_df.empty:
            continue

        raw_df.columns = raw_df.columns.astype(str).str.strip()

        name_col = _find_first_column(
            raw_df,
            [
                "정류장/역명", "정류소명", "정류장명", "정류소명칭",
                "정류장명칭", "정류장 명칭", "NODE_NM", "node_nm",
                "역사명", "역명", "역명칭", "역사명칭",
                "stnKrNm", "station_nm", "station_name", "name"
            ]
        )
        lat_col = _find_first_column(
            raw_df,
            [
                "위도", "Y좌표", "Y", "lat", "latitude", "convY",
                "GPS_LATI", "gps_lati", "정류장위도", "정류소위도",
                "역위도", "역사위도"
            ]
        )
        lon_col = _find_first_column(
            raw_df,
            [
                "경도", "X좌표", "X", "lon", "lng", "longitude", "convX",
                "GPS_LONG", "gps_long", "정류장경도", "정류소경도",
                "역경도", "역사경도"
            ]
        )
        type_col = _find_first_column(
            raw_df,
            ["구분", "정류장유형", "정류소유형", "교통수단", "type", "노선유형"]
        )

        if not name_col or not lat_col or not lon_col:
            continue

        temp = pd.DataFrame()
        temp["정류장/역명"] = raw_df[name_col].astype(str)

        # 국가철도공단_도시광역철도_역사정보 파일에는 노선명이 같이 들어있을 수 있음
        # 같은 역명이 여러 노선에 반복될 수 있으므로, 노선명을 괄호로 붙여 지도/표에서 구분
        line_col = _find_first_column(raw_df, ["노선명", "노선", "route_nm", "line", "line_name"])
        if line_col:
            line_text = raw_df[line_col].astype(str).str.strip()
            station_text = temp["정류장/역명"].astype(str).str.strip()
            temp["정류장/역명"] = station_text + " (" + line_text + ")"

        temp["위도"] = pd.to_numeric(raw_df[lat_col], errors="coerce")
        temp["경도"] = pd.to_numeric(raw_df[lon_col], errors="coerce")

        if type_col:
            temp["구분"] = raw_df[type_col].astype(str)
        elif (
            "bus" in file_path.lower()
            or "버스" in file_path
            or "정류장" in file_path
            or "정류소" in file_path
            or str(name_col).upper() in ["NODE_NM"]
        ):
            temp["구분"] = "🚏 버스"
        elif (
            "subway" in file_path.lower()
            or "station" in file_path.lower()
            or "철도" in file_path
            or "도시광역철도" in file_path
            or "역사정보" in file_path
            or "역" in str(name_col)
        ):
            temp["구분"] = "🚇 지하철"
        else:
            temp["구분"] = "🚏 버스"

        # 유형 문자열 정리
        temp["구분"] = temp["구분"].apply(
            lambda x: "🚇 지하철" if any(k in str(x) for k in ["지하철", "도시철도", "전철", "역", "subway"])
            else "🚏 버스" if any(k in str(x) for k in ["버스", "정류", "bus"])
            else str(x)
        )

        temp = temp.dropna(subset=["위도", "경도"])

        # 한국 좌표 범위 밖이면 잘못된 좌표로 보고 제외
        temp = temp[
            (temp["위도"].between(33, 39))
            & (temp["경도"].between(124, 132))
        ].copy()

        frames.append(temp)

    if not frames:
        return pd.DataFrame()

    result = pd.concat(frames, ignore_index=True)
    result = result.drop_duplicates(subset=["구분", "정류장/역명", "위도", "경도"])
    return result


def get_transit_info_from_local_file(lat, lon, radius_m):
    transit_master = load_local_transit_master()

    if transit_master.empty:
        return pd.DataFrame()

    df = transit_master.copy()

    # 최초 로딩 속도 개선: apply 대신 numpy 벡터 연산으로 거리 계산
    lat1 = np.radians(float(lat))
    lon1 = np.radians(float(lon))
    lat2 = np.radians(pd.to_numeric(df["위도"], errors="coerce"))
    lon2 = np.radians(pd.to_numeric(df["경도"], errors="coerce"))

    dlat = lat2 - lat1
    dlon = lon2 - lon1

    a = np.sin(dlat / 2) ** 2 + np.cos(lat1) * np.cos(lat2) * np.sin(dlon / 2) ** 2
    c = 2 * np.arctan2(np.sqrt(a), np.sqrt(1 - a))
    df["거리_m"] = 6371000 * c

    df = df[df["거리_m"] <= radius_m].sort_values("거리_m").copy()

    if df.empty:
        return pd.DataFrame()

    st.session_state["transit_status"] = "로컬"
    return df[["구분", "정류장/역명", "위도", "경도", "거리_m"]]

# ==========================================
# 4. 대중교통 정보 수집 함수
# ==========================================
@st.cache(suppress_st_warning=True)
def get_transit_info(lat, lon, radius_m):
    """
    1순위: 로컬 대중교통 좌표 마스터 파일에서 반경 필터링
    2순위: 로컬 파일이 없을 때만 Overpass를 짧게 시도
    """

    local_df = get_transit_info_from_local_file(lat, lon, radius_m)
    if not local_df.empty:
        return local_df

    overpass_url = "https://overpass-api.de/api/interpreter"

    # 너무 넓은 반경은 대중교통 조회가 느려짐
    # 대중교통은 최대 800m까지만 조회
    transit_radius = min(radius_m, 800)

    overpass_query = f"""
    [out:json][timeout:5];
    (
      node["highway"="bus_stop"](around:{transit_radius},{lat},{lon});
      node["railway"="station"](around:{transit_radius},{lat},{lon});
      node["station"="subway"](around:{transit_radius},{lat},{lon});
    );
    out body;
    """

    try:
        response = requests.get(
            overpass_url,
            params={"data": overpass_query},
            timeout=5
        )

        response.raise_for_status()
        data = response.json()

        elements = data.get("elements", [])

        if not elements:
            st.session_state["transit_status"] = "없음"
            return pd.DataFrame()

        transit_list = []

        for element in elements:
            tags = element.get("tags", {})
            name = tags.get("name", "이름없음")

            elat = element.get("lat")
            elon = element.get("lon")

            if elat is None or elon is None:
                continue

            if tags.get("highway") == "bus_stop":
                transit_type = "🚏 버스"
            elif (
                tags.get("railway") == "station"
                or tags.get("station") == "subway"
            ):
                transit_type = "🚇 지하철"
            else:
                continue

            transit_list.append({
                "구분": transit_type,
                "정류장/역명": name,
                "위도": elat,
                "경도": elon
            })

        if not transit_list:
            st.session_state["transit_status"] = "없음"
            return pd.DataFrame()

        transit_df = pd.DataFrame(transit_list)

        transit_df = transit_df.drop_duplicates(
            subset=["구분", "정류장/역명", "위도", "경도"]
        )

        st.session_state["transit_status"] = "정상"
        return transit_df

    except Exception:
        # 여기서 에러를 띄우지 말고 조용히 생략
        st.session_state["transit_status"] = "실패"
        return pd.DataFrame()


# ==========================================
# 4-1. 상권 기반 치과 매출 추정 함수
# ==========================================
def format_won(value):
    try:
        value = float(value)
    except Exception:
        return "정보없음"

    if pd.isna(value):
        return "정보없음"

    if abs(value) >= 100000000:
        return f"{value / 100000000:.1f}억 원"
    elif abs(value) >= 10000:
        return f"{value / 10000:.0f}만 원"
    else:
        return f"{value:,.0f}원"


def _money_to_number(series):
    return pd.to_numeric(
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("원", "", regex=False)
        .str.replace("천원", "", regex=False)
        .str.replace("만원", "", regex=False)
        .str.replace("-", "", regex=False)
        .str.strip(),
        errors="coerce"
    )



def _korean_money_text_to_won(value, default_multiplier=1):
    """
    '7억4,071만원', '74071만원', '740,710,000' 같은 값을 원 단위 숫자로 변환한다.
    """
    if pd.isna(value):
        return pd.NA

    text = str(value).strip()
    if not text or text in ["-", "*", "nan", "None"]:
        return pd.NA

    text = text.replace(",", "").replace(" ", "")

    try:
        # 7.4억원, 74071만원 같이 단위가 명시된 경우
        total = 0.0
        matched = False

        m = re.search(r"([0-9]+(?:\.[0-9]+)?)억", text)
        if m:
            total += float(m.group(1)) * 100000000
            matched = True

        m = re.search(r"([0-9]+(?:\.[0-9]+)?)만", text)
        if m:
            total += float(m.group(1)) * 10000
            matched = True

        m = re.search(r"([0-9]+(?:\.[0-9]+)?)천", text)
        if m and not matched:
            total += float(m.group(1)) * 1000
            matched = True

        if matched:
            return total

        number = re.sub(r"[^0-9.-]", "", text)
        if not number:
            return pd.NA
        return float(number) * default_multiplier

    except Exception:
        return pd.NA


def _money_series_to_won(series, col_name=""):
    col_text = str(col_name)
    if "억원" in col_text:
        multiplier = 100000000
    elif "백만원" in col_text:
        multiplier = 1000000
    elif "천원" in col_text:
        multiplier = 1000
    elif "만원" in col_text:
        multiplier = 10000
    else:
        multiplier = 1

    return series.apply(lambda x: _korean_money_text_to_won(x, multiplier))


def _read_table_flexible(file_path):
    """CSV/XLSX를 최대한 유연하게 읽는다. XLSX는 모든 시트를 합친다."""
    if file_path.lower().endswith((".xlsx", ".xls")):
        sheets = pd.read_excel(file_path, sheet_name=None, engine="openpyxl")
        dfs = []
        for sheet_name, sheet_df in sheets.items():
            if sheet_df is not None and not sheet_df.empty:
                sheet_df = sheet_df.copy()
                sheet_df["__sheet_name"] = sheet_name
                dfs.append(sheet_df)
        return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

    try:
        return pd.read_csv(file_path, encoding="utf-8-sig", low_memory=False)
    except Exception:
        return pd.read_csv(file_path, encoding="cp949", low_memory=False)



def _normalize_period_to_yyyymm(value):
    """진료년도/진료월/기준기간 값을 YYYY-MM 문자열로 통일한다."""
    if pd.isna(value):
        return None
    text = str(value).strip()
    if not text or text.lower() in ["nan", "none"]:
        return None

    # 2026-01, 2026.01, 202601, 2026년 1월 모두 처리
    nums = re.findall(r"\d+", text)
    if len(nums) >= 2 and len(nums[0]) == 4:
        year = int(nums[0])
        month = int(nums[1])
        if 1 <= month <= 12:
            return f"{year:04d}-{month:02d}"

    digits = re.sub(r"[^0-9]", "", text)
    if len(digits) >= 6:
        year = int(digits[:4])
        month = int(digits[4:6])
        if 1 <= month <= 12:
            return f"{year:04d}-{month:02d}"
    elif len(digits) == 4:
        year = int(digits)
        return f"{year:04d}"

    return text

def _period_sort_key(value):
    period = _normalize_period_to_yyyymm(value)
    if period is None:
        return -1
    digits = re.sub(r"[^0-9]", "", str(period))
    if not digits:
        return -1
    if len(digits) == 4:
        digits += "00"
    return int(digits[:6])

def _region_candidates_from_dentists(df_dentist_merged):
    """치과 주소에서 시도/시군구 후보를 뽑는다."""
    candidates = {"sido": [], "sigungu": [], "dong": []}

    if df_dentist_merged is None or df_dentist_merged.empty:
        return candidates

    addr_cols = [c for c in ["지번주소", "도로명주소", "주소"] if c in df_dentist_merged.columns]
    addresses = []
    for col in addr_cols:
        addresses.extend(df_dentist_merged[col].dropna().astype(str).head(30).tolist())

    for addr in addresses:
        parts = addr.split()
        if not parts:
            continue

        # 시도 후보
        if len(parts) >= 1:
            candidates["sido"].append(parts[0])

        # 시군구 후보: 서울 중구 / 경기 안산시 단원구 같은 케이스 모두 대응
        for i, part in enumerate(parts):
            if part.endswith(("시", "군", "구")):
                candidates["sigungu"].append(part)
                if i + 1 < len(parts) and parts[i + 1].endswith("구") and part.endswith("시"):
                    candidates["sigungu"].append(part + " " + parts[i + 1])

        for part in parts:
            if part.endswith("동"):
                candidates["dong"].append(part)

    if "동이름" in df_dentist_merged.columns:
        candidates["dong"].extend(df_dentist_merged["동이름"].dropna().astype(str).unique().tolist())

    # 긴 후보를 먼저 매칭하도록 정렬
    for key in candidates:
        unique = list(dict.fromkeys([x.strip() for x in candidates[key] if str(x).strip() and str(x).strip() != "기타"]))
        candidates[key] = sorted(unique, key=len, reverse=True)

    return candidates


@st.cache(suppress_st_warning=True)
def load_health_insurance_claims_sales_master():
    """
    건강보험 시군구별 진료과목별 진료정보 파일을 읽는다.

    금액 컬럼이 있으면 금액을 그대로 사용하고,
    금액 컬럼이 없으면 진료건수 × 평균진료단가로 치과 매출을 추정한다.
    """
    file_candidates = [
        "nhis_sggu_subject_medical.csv", "nhis_sggu_subject_medical.xlsx",
        "dental_claims_sales.csv", "dental_claims_sales.xlsx",
        "health_insurance_dental_claims.csv", "health_insurance_dental_claims.xlsx",
        "국민건강보험공단_시군구별 진료과목별 진료 정보_20231231.csv",
        "국민건강보험공단_시군구별 진료과목별 진료 정보_20231231.xlsx",
        "국민건강보험공단_시군구별 진료과목별 진료 정보.csv",
        "국민건강보험공단_시군구별 진료과목별 진료 정보.xlsx",
    ]

    # 특정 2023년 파일 하나만 잡히는 문제를 막기 위해 후보명 + 국민건강보험공단 파일을 모두 합친다.
    file_paths = []
    for cand in file_candidates:
        if os.path.exists(cand):
            file_paths.append(cand)

    wildcard_patterns = [
        "*시군구별*진료과목별*진료*정보*.csv",
        "*시군구별*진료과목별*진료*정보*.xlsx",
        "*국민건강보험공단*진료과목별*.csv",
        "*국민건강보험공단*진료과목별*.xlsx",
        "*nhis*medical*.csv",
        "*nhis*medical*.xlsx",
    ]
    for pattern in wildcard_patterns:
        for path in glob.glob(pattern):
            if path not in file_paths:
                file_paths.append(path)

    if not file_paths:
        return pd.DataFrame(), (
            "건강보험 진료정보 파일이 없습니다. "
            "공공데이터포털에서 '국민건강보험공단_시군구별 진료과목별 진료 정보' CSV를 받은 뒤 "
            "nhis_sggu_subject_medical.csv 이름으로 app.py와 같은 폴더에 넣어주세요."
        )

    loaded = []
    read_errors = []
    for file_path in file_paths:
        try:
            temp_df = _read_table_flexible(file_path)
            if temp_df is not None and not temp_df.empty:
                temp_df = temp_df.copy()
                temp_df["__source_file"] = os.path.basename(file_path)
                loaded.append(temp_df)
        except Exception as e:
            read_errors.append(f"{file_path}: {e}")

    if not loaded:
        return pd.DataFrame(), f"건강보험 진료정보 파일을 읽지 못했습니다: {read_errors}"

    raw_df = pd.concat(loaded, ignore_index=True)

    if raw_df.empty:
        return pd.DataFrame(), "건강보험 진료정보 파일이 비어 있습니다."

    raw_df.columns = raw_df.columns.astype(str).str.strip()

    year_col = _find_first_column(raw_df, ["진료년도", "진료연도", "기준년도", "기준연도", "연도", "년도", "year"])
    month_col = _find_first_column(raw_df, ["진료월", "월", "month"])
    period_col = _find_first_column(raw_df, ["진료연월", "기준년월", "연월", "년월", "period"])
    sido_col = _find_first_column(raw_df, ["시도", "시도명", "지역", "지역명", "sido"])
    sigungu_col = _find_first_column(raw_df, ["시군구", "시군구명", "구군", "구군명", "sigungu"])
    category_col = _find_first_column(raw_df, ["진료과목명", "진료과목", "표시과목", "과목명", "업종명", "category"])
    patient_col = _find_first_column(raw_df, ["진료인원(명)", "진료인원", "환자수", "수진자수", "실인원", "patient"])
    case_col = _find_first_column(raw_df, ["진료건수(건)", "진료건수", "명세서청구건수", "청구건수", "건수", "case"])
    amount_col = _find_first_column(raw_df, [
        "요양급여비용총액", "요양급여비용총액(선별포함)", "진료비", "총진료비",
        "급여비용", "보험급여비", "보험자부담금", "보험자부담금(선별포함)",
        "금액", "진료금액", "amount", "cost"
    ])

    if not category_col:
        return pd.DataFrame(), (
            "건강보험 진료정보 파일에서 진료과목명 컬럼을 찾지 못했습니다. "
            f"현재 컬럼: {raw_df.columns.tolist()}"
        )

    if not case_col and not amount_col:
        return pd.DataFrame(), (
            "건강보험 진료정보 파일에 금액 컬럼도 진료건수 컬럼도 없습니다. "
            "최소한 '진료건수(건)' 컬럼은 있어야 진료건수 기반 추정매출을 계산할 수 있습니다. "
            f"현재 컬럼: {raw_df.columns.tolist()}"
        )

    df = pd.DataFrame()

    if period_col:
        df["기준기간"] = raw_df[period_col].apply(_normalize_period_to_yyyymm)
    elif year_col and month_col:
        year_part = raw_df[year_col].astype(str).str.extract(r"(\d{4})")[0]
        month_part = raw_df[month_col].astype(str).str.extract(r"(\d{1,2})")[0]
        df["기준기간"] = (year_part.fillna("") + "-" + month_part.fillna("01").str.zfill(2)).apply(_normalize_period_to_yyyymm)
    elif year_col:
        df["기준기간"] = raw_df[year_col].astype(str).str.extract(r"(\d{4})")[0].apply(_normalize_period_to_yyyymm)
    else:
        df["기준기간"] = "최근"

    df["시도"] = raw_df[sido_col].astype(str).str.strip() if sido_col else ""
    df["시군구"] = raw_df[sigungu_col].astype(str).str.strip() if sigungu_col else ""
    df["업종명"] = raw_df[category_col].astype(str).str.strip()
    if "__source_file" in raw_df.columns:
        df["원본파일"] = raw_df["__source_file"].astype(str)

    df["진료인원"] = _money_to_number(raw_df[patient_col]) if patient_col else 0
    df["진료건수"] = _money_to_number(raw_df[case_col]) if case_col else 0

    # 치과 기준 필터: 치과계 진료과목만 사용한다. 일반 내과/정형외과/피부과 등은 제외된다.
    dental_keywords = "치과|구강악안면|구강|보철|보존|교정|치주|소아치과"
    df = df[df["업종명"].astype(str).str.contains(dental_keywords, na=False, regex=True)].copy()

    if df.empty:
        return pd.DataFrame(), "건강보험 파일에서 치과 관련 진료과목 데이터를 찾지 못했습니다."

    avg_fee_per_case = 45000

    if amount_col:
        df["추정매출"] = _money_series_to_won(raw_df.loc[df.index, amount_col], amount_col)
        df["계산방식"] = "건강보험 금액 컬럼 사용"
        df["평균진료단가"] = pd.NA
    else:
        df["평균진료단가"] = avg_fee_per_case
        df["추정매출"] = pd.to_numeric(df["진료건수"], errors="coerce").fillna(0) * avg_fee_per_case
        df["계산방식"] = f"진료건수 × 평균진료단가 {avg_fee_per_case:,}원"

    df = df.dropna(subset=["추정매출"])
    df = df[df["추정매출"] > 0].copy()

    if df.empty:
        return pd.DataFrame(), "건강보험 치과 진료정보에서 추정매출을 계산할 수 있는 행이 없습니다."

    return df, None



def _simple_sido_name(value):
    text = str(value) if value is not None else ""
    return (
        text.replace("특별시", "")
        .replace("광역시", "")
        .replace("특별자치시", "")
        .replace("특별자치도", "")
        .replace("도", "")
        .strip()
    )


def _extract_region_from_address_text(address):
    """주소 문자열에서 시도/시군구를 최대한 추출한다."""
    text = str(address)
    parts = text.split()

    sido = ""
    sigungu = ""

    if parts:
        sido = parts[0]

    # 서울특별시 중구, 경기도 안산시 단원구, 충청남도 천안시 서북구 모두 대응
    for i, part in enumerate(parts):
        if part.endswith(("시", "군", "구")):
            if part.endswith("시") and i + 1 < len(parts) and parts[i + 1].endswith("구"):
                sigungu = parts[i + 1]  # 건강보험 파일은 보통 '단원구'처럼 들어감
                break
            if part.endswith(("군", "구")):
                sigungu = part
                break
            if part.endswith("시") and not sigungu:
                sigungu = part

    return sido, sigungu


def _attach_region_to_dentists(df_dentist_merged):
    """반경 내 치과별 시도/시군구 컬럼을 붙인다."""
    df = df_dentist_merged.copy()

    if df.empty:
        df["매출시도"] = "지역미상"
        df["매출시군구"] = "지역미상"
        return df

    addr_cols = [c for c in ["지번주소", "도로명주소", "주소"] if c in df.columns]
    if not addr_cols:
        df["매출시도"] = "지역미상"
        df["매출시군구"] = "지역미상"
        return df

    def pick_addr(row):
        for col in addr_cols:
            val = str(row.get(col, "")).strip()
            if val and val.lower() != "nan":
                return val
        return ""

    regions = df.apply(lambda row: _extract_region_from_address_text(pick_addr(row)), axis=1)
    df["매출시도"] = [r[0] if r[0] else "지역미상" for r in regions]
    df["매출시군구"] = [r[1] if r[1] else "지역미상" for r in regions]
    return df


def _count_total_dental_clinics_by_sggu(sido, sigungu, fallback_count=1):
    """hosp_data.xlsx 전체에서 해당 시군구 전체 치과 수를 계산한다."""
    try:
        df_hosp, err = load_hosp_data()
        if err or df_hosp.empty:
            return max(int(fallback_count), 1)

        df_hosp = df_hosp.copy()
        addr_cols = [
            "주소", "도로명주소", "지번주소", "소재지주소", "요양기관주소", "소재지",
            "시도코드명", "시군구코드명", "시도", "시군구"
        ]
        existing = [c for c in addr_cols if c in df_hosp.columns]
        if not existing:
            return max(int(fallback_count), 1)

        df_hosp["주소통합"] = ""
        for col in existing:
            df_hosp["주소통합"] += " " + df_hosp[col].astype(str)

        target = df_hosp.copy()

        if sido and sido != "지역미상":
            sido_simple = _simple_sido_name(sido)
            temp = target[target["주소통합"].astype(str).str.contains(re.escape(sido_simple), na=False)]
            if not temp.empty:
                target = temp

        if sigungu and sigungu != "지역미상":
            temp = target[target["주소통합"].astype(str).str.contains(re.escape(sigungu), na=False)]
            if not temp.empty:
                target = temp

        count = len(target)
        return max(int(count), int(fallback_count), 1)

    except Exception:
        return max(int(fallback_count), 1)


def _filter_claims_by_region(claims_df, sido, sigungu):
    """건강보험 치과 진료정보에서 특정 시도/시군구를 매칭한다."""
    target = claims_df.copy()

    if sido and sido != "지역미상" and "시도" in target.columns:
        sido_simple = _simple_sido_name(sido)
        temp = target[
            target["시도"].astype(str).str.contains(re.escape(str(sido)), na=False)
            | target["시도"].astype(str).str.contains(re.escape(sido_simple), na=False)
        ]
        if not temp.empty:
            target = temp

    if sigungu and sigungu != "지역미상" and "시군구" in target.columns:
        temp = target[target["시군구"].astype(str).str.contains(re.escape(str(sigungu)), na=False)]
        if not temp.empty:
            target = temp

    return target


def estimate_dental_sales_from_health_claims(df_dentist_merged):
    """
    반경이 여러 구/시군구에 걸치는 경우를 반영한 매출 추정.

    공식:
    각 시군구 치과 1곳당 월 추정매출 = 해당 시군구 치과 월 추정시장규모 / 해당 시군구 전체 치과 수
    반경 내 월 추정시장규모 = Σ(각 시군구 1곳당 월 추정매출 × 반경 내 해당 시군구 치과 수)
    """
    claims_df, err = load_health_insurance_claims_sales_master()
    if err:
        return None, pd.DataFrame(), err

    if df_dentist_merged is None or df_dentist_merged.empty:
        return None, pd.DataFrame(), "반경 내 치과가 없어 매출을 추정할 수 없습니다."

    df_radius = _attach_region_to_dentists(df_dentist_merged)

    # 반경 내 치과를 시도/시군구별로 묶는다.
    region_counts = (
        df_radius.groupby(["매출시도", "매출시군구"], dropna=False)
        .size()
        .reset_index(name="반경내치과수")
    )

    # 파일 전체 최신기간도 같이 계산
    claims_df = claims_df.copy()
    claims_df["__period_key"] = claims_df["기준기간"].apply(_period_sort_key)
    all_latest_key = claims_df["__period_key"].max()
    all_latest_df = claims_df[claims_df["__period_key"] == all_latest_key].copy()
    file_latest_period = str(all_latest_df["기준기간"].dropna().astype(str).iloc[0]) if not all_latest_df.empty else "최근"

    detail_rows = []

    for _, region_row in region_counts.iterrows():
        sido = str(region_row["매출시도"])
        sigungu = str(region_row["매출시군구"])
        radius_count = int(region_row["반경내치과수"])

        region_claims = _filter_claims_by_region(claims_df, sido, sigungu)

        # 해당 지역 매칭 실패 시 전체 파일 사용. 단, detail에 표시한다.
        region_match_note = f"{sido} {sigungu}"
        if region_claims.empty:
            region_claims = claims_df.copy()
            region_match_note = "지역매칭실패/전국평균"

        region_claims["__period_key"] = region_claims["기준기간"].apply(_period_sort_key)
        latest_key = region_claims["__period_key"].max()
        latest_region_df = region_claims[region_claims["__period_key"] == latest_key].copy()

        if latest_region_df.empty:
            continue

        latest_period = str(latest_region_df["기준기간"].dropna().astype(str).iloc[0]) if not latest_region_df.empty else "최근"
        sggu_total_sales = pd.to_numeric(latest_region_df["추정매출"], errors="coerce").sum()
        total_cases = pd.to_numeric(latest_region_df.get("진료건수", pd.Series(dtype=float)), errors="coerce").sum()
        total_patients = pd.to_numeric(latest_region_df.get("진료인원", pd.Series(dtype=float)), errors="coerce").sum()

        sggu_total_clinic_count = _count_total_dental_clinics_by_sggu(
            sido=sido,
            sigungu=sigungu,
            fallback_count=radius_count
        )

        sales_per_clinic = sggu_total_sales / max(sggu_total_clinic_count, 1)
        radius_region_sales = sales_per_clinic * radius_count

        calc_methods = latest_region_df.get("계산방식", pd.Series(["건강보험 진료건수 기반 추정"])).dropna().astype(str).unique().tolist()
        calc_method = calc_methods[0] if calc_methods else "건강보험 진료건수 기반 추정"

        detail_rows.append({
            "기준기간": latest_period,
            "파일최신기간": file_latest_period,
            "시도": sido,
            "시군구": sigungu,
            "매칭기준": region_match_note,
            "계산방식": calc_method,
            "업종명": ", ".join(latest_region_df["업종명"].dropna().astype(str).unique().tolist()[:4]),
            "시군구치과월추정시장규모": float(sggu_total_sales),
            "시군구전체치과수": int(sggu_total_clinic_count),
            "반경내치과수": int(radius_count),
            "치과1곳당월추정매출": float(sales_per_clinic),
            "반경내월추정시장규모": float(radius_region_sales),
            "진료건수": float(total_cases),
            "진료인원": float(total_patients),
        })

    detail_df = pd.DataFrame(detail_rows)

    if detail_df.empty:
        return None, pd.DataFrame(), "건강보험 치과 진료정보와 반경 내 치과 지역을 매칭하지 못했습니다."

    total_radius_sales = detail_df["반경내월추정시장규모"].sum()
    total_radius_clinics = int(detail_df["반경내치과수"].sum())
    weighted_avg_per_clinic = total_radius_sales / max(total_radius_clinics, 1)
    total_sggu_market = detail_df["시군구치과월추정시장규모"].sum()
    total_cases = detail_df["진료건수"].sum()
    total_patients = detail_df["진료인원"].sum()

    periods = detail_df["기준기간"].dropna().astype(str).unique().tolist()
    if len(periods) == 1:
        period_label = periods[0]
    else:
        period_label = f"혼합({', '.join(periods[:3])}{'...' if len(periods) > 3 else ''})"

    calc_methods = detail_df["계산방식"].dropna().astype(str).unique().tolist()
    calc_method = calc_methods[0] if len(calc_methods) == 1 else "지역별 건강보험 진료정보 기반 혼합 계산"

    summary = {
        "자료종류": "국민건강보험공단 시군구별 진료과목별 진료정보",
        "기준기간": period_label,
        "파일최신기간": file_latest_period,
        "매칭기준": f"시군구별 가중계산 {len(detail_df)}개 지역",
        "계산방식": calc_method,
        "매출행수": len(detail_df),
        "반경내치과수": total_radius_clinics,
        "진료건수": total_cases,
        "진료인원": total_patients,
        "시군구치과월추정시장규모합계": float(total_sggu_market),
        "반경내월추정시장규모": float(total_radius_sales),
        "반경내치과1곳당평균월추정매출": float(weighted_avg_per_clinic),
        "최고지역1곳당월추정매출": float(detail_df["치과1곳당월추정매출"].max()),
        "최저지역1곳당월추정매출": float(detail_df["치과1곳당월추정매출"].min()),
        # 이전 UI와의 호환용 키
        "상권치과업종총추정매출": float(total_sggu_market),
        "치과1곳당추정매출": float(weighted_avg_per_clinic),
        "반경내치과추정총매출": float(total_radius_sales),
        "최고1곳당추정매출": float(detail_df["치과1곳당월추정매출"].max()),
        "최저1곳당추정매출": float(detail_df["치과1곳당월추정매출"].min()),
    }

    return summary, detail_df, None


@st.cache(suppress_st_warning=True)
def load_nts_living_industry_sales_master():
    """
    국세통계포털(TASIS) '통계로 보는 생활업종'에서 받은 치과의원 매출 파일을 읽는다.

    권장 파일명:
    - nts_living_industry.xlsx / .csv
    - nts_dental_sales.xlsx / .csv
    - 국세청_통계로보는생활업종.xlsx / .csv
    - 생활업종_치과의원_매출.xlsx / .csv

    권장 컬럼:
    귀속연도, 시도, 시군구, 업종명, 연평균매출, 사업자수, 평균사업존속연수
    """
    file_candidates = [
        "nts_living_industry.xlsx", "nts_living_industry.csv",
        "nts_dental_sales.xlsx", "nts_dental_sales.csv",
        "국세청_통계로보는생활업종.xlsx", "국세청_통계로보는생활업종.csv",
        "국세청_생활업종_매출.xlsx", "국세청_생활업종_매출.csv",
        "생활업종_치과의원_매출.xlsx", "생활업종_치과의원_매출.csv",
        "치과의원_연평균매출.xlsx", "치과의원_연평균매출.csv",
    ]

    file_path = None
    for cand in file_candidates:
        if os.path.exists(cand):
            file_path = cand
            break

    if file_path is None:
        return pd.DataFrame(), "국세청 생활업종 매출 파일이 없습니다. nts_living_industry.xlsx 또는 nts_dental_sales.csv를 같은 폴더에 넣어주세요."

    try:
        raw_df = _read_table_flexible(file_path)
    except Exception as e:
        return pd.DataFrame(), f"국세청 생활업종 매출 파일을 읽지 못했습니다: {e}"

    if raw_df.empty:
        return pd.DataFrame(), "국세청 생활업종 매출 파일이 비어 있습니다."

    raw_df.columns = raw_df.columns.astype(str).str.strip()

    year_col = _find_first_column(raw_df, [
        "귀속연도", "귀속년도", "기준연도", "기준년도", "연도", "년도", "년", "sttPblYr", "year"
    ])
    sido_col = _find_first_column(raw_df, [
        "시도", "시도명", "광역시도", "광역자치단체", "광역단체", "지역", "지역명", "sido"
    ])
    sigungu_col = _find_first_column(raw_df, [
        "시군구", "시군구명", "시군구별", "기초자치단체", "구군", "구군명", "시군구코드명", "sigungu"
    ])
    category_col = _find_first_column(raw_df, [
        "업종명", "생활업종명", "생활밀접업종", "100대 생활업종", "업종", "업태", "category"
    ])
    avg_sales_col = _find_first_column(raw_df, [
        "연평균매출", "연평균 매출", "연평균매출액", "평균연매출", "평균 연매출",
        "평균매출", "평균매출액", "평균 총수입금액", "총수입금액", "매출액", "수입금액"
    ])
    biz_count_col = _find_first_column(raw_df, [
        "사업자수", "사업자 수", "개업사업자수", "가동사업자수", "사업체수", "업소수", "점포수", "업체수", "count"
    ])
    survive_col = _find_first_column(raw_df, [
        "평균사업존속연수", "평균 사업존속연수", "평균존속연수", "사업존속연수", "존속연수"
    ])

    if not avg_sales_col:
        return pd.DataFrame(), (
            "국세청 생활업종 파일에서 연평균매출/평균매출 컬럼을 찾지 못했습니다. "
            f"현재 컬럼: {raw_df.columns.tolist()}"
        )

    df = pd.DataFrame()
    df["귀속연도"] = raw_df[year_col].astype(str) if year_col else "파일 기준"
    df["시도"] = raw_df[sido_col].astype(str).str.strip() if sido_col else ""
    df["시군구"] = raw_df[sigungu_col].astype(str).str.strip() if sigungu_col else ""
    df["업종명"] = raw_df[category_col].astype(str).str.strip() if category_col else "치과의원"
    df["연평균매출"] = _money_series_to_won(raw_df[avg_sales_col], avg_sales_col)

    if biz_count_col:
        df["사업자수"] = pd.to_numeric(
            raw_df[biz_count_col].astype(str).str.replace(",", "", regex=False).str.replace("명", "", regex=False),
            errors="coerce"
        )
    else:
        df["사업자수"] = pd.NA

    if survive_col:
        df["평균사업존속연수"] = raw_df[survive_col].astype(str)
    else:
        df["평균사업존속연수"] = ""

    # 여러 업종이 섞인 파일이면 치과의원만 사용. 이미 치과 파일이면 그대로 통과.
    if category_col:
        dental_only = df[df["업종명"].str.contains("치과", na=False)].copy()
        if not dental_only.empty:
            df = dental_only

    df = df.dropna(subset=["연평균매출"])
    df = df[df["연평균매출"] > 0].copy()

    if df.empty:
        return pd.DataFrame(), "국세청 생활업종 파일에서 치과의원 연평균매출 데이터를 찾지 못했습니다."

    return df, None


def estimate_dental_sales_from_nts(df_dentist_merged):
    nts_df, err = load_nts_living_industry_sales_master()
    if err:
        return None, pd.DataFrame(), err

    df = nts_df.copy()
    candidates = _region_candidates_from_dentists(df_dentist_merged)

    match_basis = "국세청 생활업종 파일 전체 기준"
    target_df = pd.DataFrame()

    # 1) 시군구 우선 매칭
    if "시군구" in df.columns and candidates["sigungu"]:
        for sigungu in candidates["sigungu"]:
            temp = df[df["시군구"].astype(str).str.contains(re.escape(sigungu), na=False)].copy()
            if not temp.empty:
                target_df = temp
                match_basis = f"국세청 생활업종 · {sigungu} 기준"
                break

    # 2) 시도 매칭
    if target_df.empty and "시도" in df.columns and candidates["sido"]:
        for sido in candidates["sido"]:
            simple_sido = sido.replace("특별시", "").replace("광역시", "").replace("특별자치시", "").replace("특별자치도", "").replace("도", "")
            pattern = re.escape(sido) + "|" + re.escape(simple_sido)
            temp = df[df["시도"].astype(str).str.contains(pattern, na=False)].copy()
            if not temp.empty:
                target_df = temp
                match_basis = f"국세청 생활업종 · {sido} 기준"
                break

    # 3) 전체 평균
    if target_df.empty:
        target_df = df.copy()
        match_basis = "국세청 생활업종 · 전국/파일 전체 기준"

    if "귀속연도" in target_df.columns and target_df["귀속연도"].notna().any():
        latest_period = sorted(target_df["귀속연도"].astype(str).unique())[-1]
        latest_df = target_df[target_df["귀속연도"].astype(str) == latest_period].copy()
    else:
        latest_period = "파일 기준"
        latest_df = target_df.copy()

    dental_count = len(df_dentist_merged) if df_dentist_merged is not None and not df_dentist_merged.empty else 0
    avg_per_store = pd.to_numeric(latest_df["연평균매출"], errors="coerce").mean()

    # 파일 전체에서 최신연도 기준 최고/최저 지역 평균매출도 같이 계산
    if "귀속연도" in df.columns and df["귀속연도"].notna().any():
        all_latest_period = sorted(df["귀속연도"].astype(str).unique())[-1]
        all_latest_df = df[df["귀속연도"].astype(str) == all_latest_period].copy()
    else:
        all_latest_df = df.copy()

    all_sales = pd.to_numeric(all_latest_df["연평균매출"], errors="coerce").dropna()
    max_sales = all_sales.max() if not all_sales.empty else avg_per_store
    min_sales = all_sales.min() if not all_sales.empty else avg_per_store

    estimated_radius_total = avg_per_store * dental_count if dental_count > 0 else avg_per_store

    detail_df = latest_df.copy()
    detail_df["1곳당추정매출"] = detail_df["연평균매출"]
    detail_df["추정매출"] = detail_df["연평균매출"]
    detail_df["점포수"] = detail_df.get("사업자수", pd.NA)
    detail_df["기준기간"] = detail_df["귀속연도"]
    detail_df["동이름"] = detail_df.apply(
        lambda row: " ".join([str(row.get("시도", "")).strip(), str(row.get("시군구", "")).strip()]).strip(),
        axis=1
    )

    summary = {
        "기준기간": latest_period,
        "파일최신기간": str(all_latest_df["기준기간"].dropna().astype(str).iloc[0]) if not all_latest_df.empty else latest_period,
        "매칭기준": match_basis,
        "자료종류": "국세청 생활업종 연평균매출",
        "매출행수": len(latest_df),
        "반경내치과수": dental_count,
        "상권치과업종총추정매출": estimated_radius_total,
        "치과1곳당추정매출": avg_per_store,
        "반경내치과추정총매출": estimated_radius_total,
        "최고1곳당추정매출": max_sales,
        "최저1곳당추정매출": min_sales,
    }

    return summary, detail_df, None


@st.cache(suppress_st_warning=True)
def load_sales_estimate_master():
    """
    소상공인/상권분석에서 내려받은 업종별 추정매출 파일을 읽는다.
    app.py와 같은 폴더에 아래 이름 중 하나로 저장하면 자동 인식한다.
    - dental_sales.xlsx / dental_sales.csv
    - sales_estimate.xlsx / sales_estimate.csv
    - sbiz_sales.xlsx / sbiz_sales.csv
    - 소상공인_상권분석_추정매출.xlsx / .csv

    권장 컬럼 예시:
    기준년월 또는 기준분기, 행정동 또는 동이름, 업종명, 추정매출, 점포수
    """
    file_candidates = [
        "dental_sales.xlsx", "dental_sales.csv",
        "sales_estimate.xlsx", "sales_estimate.csv",
        "sbiz_sales.xlsx", "sbiz_sales.csv",
        "소상공인_상권분석_추정매출.xlsx", "소상공인_상권분석_추정매출.csv",
        "상권_추정매출.xlsx", "상권_추정매출.csv",
    ]

    file_path = None
    for cand in file_candidates:
        if os.path.exists(cand):
            file_path = cand
            break

    if file_path is None:
        return pd.DataFrame(), "매출 추정 파일이 없습니다. dental_sales.xlsx 또는 sales_estimate.csv를 app.py와 같은 폴더에 넣어주세요."

    try:
        if file_path.lower().endswith(".xlsx"):
            raw_df = pd.read_excel(file_path, engine="openpyxl")
        else:
            try:
                raw_df = pd.read_csv(file_path, encoding="utf-8-sig", low_memory=False)
            except Exception:
                raw_df = pd.read_csv(file_path, encoding="cp949", low_memory=False)
    except Exception as e:
        return pd.DataFrame(), f"매출 추정 파일을 읽지 못했습니다: {e}"

    if raw_df.empty:
        return pd.DataFrame(), "매출 추정 파일이 비어 있습니다."

    raw_df.columns = raw_df.columns.astype(str).str.strip()

    dong_col = _find_first_column(raw_df, [
        "행정동", "행정동명", "법정동", "법정동명", "동이름", "상권명",
        "ADMI_NM", "admi_nm", "dong", "area"
    ])
    category_col = _find_first_column(raw_df, [
        "업종명", "서비스업종명", "상권업종소분류명", "상권업종명", "표준산업분류명",
        "업종", "indsSclsNm", "service_nm", "category"
    ])
    sales_col = _find_first_column(raw_df, [
        "추정매출", "추정매출액", "월추정매출", "월매출", "분기매출", "당월매출금액",
        "매출금액", "매출액", "총매출", "sales", "amount"
    ])
    store_col = _find_first_column(raw_df, [
        "점포수", "업소수", "상가수", "가맹점수", "매장수", "store_count", "cnt"
    ])
    period_col = _find_first_column(raw_df, [
        "기준년월", "기준년분기", "기준분기", "년월", "분기", "기준일자", "base_ym", "period"
    ])
    lat_col = _find_first_column(raw_df, ["위도", "lat", "latitude", "Y좌표", "Y"])
    lon_col = _find_first_column(raw_df, ["경도", "lon", "lng", "longitude", "X좌표", "X"])

    if not category_col or not sales_col:
        return pd.DataFrame(), (
            "매출 추정 파일에서 업종명/매출액 컬럼을 찾지 못했습니다. "
            f"현재 컬럼: {raw_df.columns.tolist()}"
        )

    df = pd.DataFrame()
    df["업종명"] = raw_df[category_col].astype(str)
    df["추정매출"] = _money_to_number(raw_df[sales_col])

    if dong_col:
        df["동이름"] = raw_df[dong_col].astype(str).str.strip()
    else:
        df["동이름"] = "전체"

    if store_col:
        df["점포수"] = pd.to_numeric(raw_df[store_col], errors="coerce")
    else:
        df["점포수"] = pd.NA

    if period_col:
        df["기준기간"] = raw_df[period_col].astype(str)
    else:
        df["기준기간"] = "파일 기준"

    if lat_col and lon_col:
        df["위도"] = pd.to_numeric(raw_df[lat_col], errors="coerce")
        df["경도"] = pd.to_numeric(raw_df[lon_col], errors="coerce")
    else:
        df["위도"] = pd.NA
        df["경도"] = pd.NA

    df = df.dropna(subset=["추정매출"])
    df = df[df["업종명"].str.contains("치과|의료|병원|의원", na=False)].copy()

    if df.empty:
        return pd.DataFrame(), "매출 파일에는 치과/의료 관련 업종 매출 행이 없습니다."

    return df, None


def estimate_dental_sales(lat, lon, radius_m, df_dentist_merged):
    # 1순위: 공공데이터포털 국민건강보험공단 시군구별 진료과목별 진료비 파일
    # 치과 실제 총매출은 아니지만, 무료로 받을 수 있는 전국 단위 '금액' 근거 데이터다.
    claims_summary, claims_detail_df, claims_err = estimate_dental_sales_from_health_claims(df_dentist_merged)
    if claims_summary is not None:
        return claims_summary, claims_detail_df, None

    # 2순위: 사용자가 직접 수집한 국세청/생활업종 매출 파일
    nts_summary, nts_detail_df, nts_err = estimate_dental_sales_from_nts(df_dentist_merged)
    if nts_summary is not None:
        return nts_summary, nts_detail_df, None

    # 3순위: 사용자가 별도로 보유한 상권/카드/소상공인 추정매출 파일
    sales_df, err = load_sales_estimate_master()

    if err:
        return None, pd.DataFrame(), "\n".join([x for x in [claims_err, nts_err, err] if x])

    df = sales_df.copy()

    # 좌표가 있는 매출 파일이면 반경으로 직접 필터링
    has_coords = (
        "위도" in df.columns and "경도" in df.columns
        and df["위도"].notna().any() and df["경도"].notna().any()
    )

    if has_coords:
        df = df.dropna(subset=["위도", "경도"]).copy()
        df["거리_m"] = df.apply(
            lambda row: calculate_distance(lat, lon, row["위도"], row["경도"]),
            axis=1
        )
        target_df = df[df["거리_m"] <= radius_m].copy()
        match_basis = f"반경 {radius_m}m 좌표 기준"
    else:
        dong_list = []
        if not df_dentist_merged.empty and "동이름" in df_dentist_merged.columns:
            dong_list = [str(x).strip() for x in df_dentist_merged["동이름"].dropna().unique().tolist()]
            dong_list = [x for x in dong_list if x and x != "기타"]

        if dong_list:
            pattern = "|".join([re.escape(x) for x in dong_list])
            target_df = df[df["동이름"].astype(str).str.contains(pattern, na=False)].copy()
            match_basis = "반경 내 치과의 동이름 기준"
        else:
            target_df = df.copy()
            match_basis = "전체 파일 기준"

    if target_df.empty:
        return None, pd.DataFrame(), "현재 반경/동 기준에 맞는 치과업종 매출 데이터가 없습니다."

    # 최신 기준기간만 사용 가능하면 최신값으로 제한
    if "기준기간" in target_df.columns and target_df["기준기간"].notna().any():
        latest_period = sorted(target_df["기준기간"].astype(str).unique())[-1]
        latest_df = target_df[target_df["기준기간"].astype(str) == latest_period].copy()
    else:
        latest_period = "파일 기준"
        latest_df = target_df.copy()

    total_sales = latest_df["추정매출"].sum()
    dental_count = len(df_dentist_merged) if df_dentist_merged is not None and not df_dentist_merged.empty else 0

    store_count = pd.to_numeric(latest_df.get("점포수", pd.Series(dtype=float)), errors="coerce").sum()
    if pd.isna(store_count) or store_count <= 0:
        store_count = dental_count

    avg_per_store = total_sales / store_count if store_count and store_count > 0 else 0
    estimated_radius_total = avg_per_store * dental_count if dental_count > 0 else total_sales

    detail_df = latest_df.copy()
    detail_df["1곳당추정매출"] = detail_df.apply(
        lambda row: row["추정매출"] / row["점포수"]
        if pd.notna(row.get("점포수")) and row.get("점포수") not in [0, "0"]
        else pd.NA,
        axis=1
    )

    per_store_values = pd.to_numeric(detail_df["1곳당추정매출"], errors="coerce").dropna()
    max_sales = per_store_values.max() if not per_store_values.empty else avg_per_store
    min_sales = per_store_values.min() if not per_store_values.empty else avg_per_store

    summary = {
        "기준기간": latest_period,
        "파일최신기간": str(all_latest_df["기준기간"].dropna().astype(str).iloc[0]) if not all_latest_df.empty else latest_period,
        "매칭기준": match_basis,
        "매출행수": len(latest_df),
        "반경내치과수": dental_count,
        "상권치과업종총추정매출": total_sales,
        "치과1곳당추정매출": avg_per_store,
        "반경내치과추정총매출": estimated_radius_total,
        "최고1곳당추정매출": max_sales,
        "최저1곳당추정매출": min_sales,
    }

    return summary, detail_df, None


# ==========================================
# 4-1. 주소 → 좌표 변환
# ==========================================
def get_coords_from_address(address):
    geolocator = Nominatim(user_agent="clinic_analyzer")
    location = geolocator.geocode(address)

    if location:
        return location.latitude, location.longitude

    return None, None


# ==========================================
# 5. 지도 생성 함수
# ==========================================
def create_map(df, df_hira, transit_df, center_lat, center_lon, radius_m):
    # 반경에 따라 적당한 줌 레벨 설정
    if radius_m <= 300:
        zoom_level = 18
    elif radius_m <= 500:
        zoom_level = 17
    elif radius_m <= 1000:
        zoom_level = 16
    elif radius_m <= 1500:
        zoom_level = 15
    else:
        zoom_level = 14

    m = folium.Map(
        location=[center_lat, center_lon],
        zoom_start=zoom_level
    )

    star_html = """
    <div style='font-size: 32px; text-shadow: 2px 2px 5px rgba(0,0,0,0.7);'>
        ⭐
    </div>
    """

    folium.Marker(
        [center_lat, center_lon],
        popup="분석 기준점",
        icon=folium.DivIcon(html=star_html)
    ).add_to(m)

    folium.Circle(
        location=[center_lat, center_lon],
        radius=radius_m,
        color="#3186cc",
        fill=True,
        fill_opacity=0.1
    ).add_to(m)

    # 치과 마커는 df_dentist_merged 기준으로 한 번만 표시
    # df_hira를 따로 찍으면 같은 치과가 중복 마커로 보일 수 있음
    if not df.empty:
        for _, row in df.iterrows():
            name = row["상호명"]
            years_active = row.get("업력(년)", 0)
            open_date = row.get("개업일", "정보없음")

            try:
                years_active = int(years_active)
            except Exception:
                years_active = 0

            marker_color = "gray"

            if years_active > 0:
                if years_active < 3:
                    marker_color = "red"
                elif years_active < 5:
                    marker_color = "orange"
                elif years_active < 10:
                    marker_color = "green"
                elif years_active < 15:
                    marker_color = "blue"
                else:
                    marker_color = "purple"

            popup_html = f"""
            <div style='width:180px'>
                <b>🦷 {name}</b><br>
                개업: {open_date}<br>
                업력: {years_active}년차
            </div>
            """

            folium.Marker(
                location=[row["위도"], row["경도"]],
                popup=folium.Popup(popup_html, max_width=300),
                tooltip=name,
                icon=folium.Icon(
                    color=marker_color,
                    icon="plus",
                    prefix="fa"
                )
            ).add_to(m)

    # 대중교통 마커
    if not transit_df.empty:
        for _, row in transit_df.iterrows():
            name = row["정류장/역명"]
            t_type = row["구분"]

            if "버스" in t_type:
                t_color = "green"
                t_icon = "bus"
            elif "지하철" in t_type:
                t_color = "blue"
                t_icon = "train"
            else:
                t_color = "cadetblue"
                t_icon = "info-sign"

            popup_html = f"""
            <div style='width:120px'>
                <b>{t_type}</b><br>
                {name}
            </div>
            """

            folium.Marker(
                location=[row["위도"], row["경도"]],
                popup=folium.Popup(popup_html, max_width=300),
                tooltip=f"{t_type} - {name}",
                icon=folium.Icon(
                    color=t_color,
                    icon=t_icon,
                    prefix="fa"
                )
            ).add_to(m)

    return m


# ==========================================
# 6. Streamlit UI 화면 구성
# ==========================================
st.set_page_config(
    page_title="치과 상권 분석 대시보드",
    layout="wide"
)

st.markdown(
    """
    <style>
        .block-container {
            max-width: 100% !important;
            padding-top: 0.45rem !important;
            padding-left: 0.75rem !important;
            padding-right: 0.75rem !important;
            padding-bottom: 0.75rem !important;
            overflow-x: hidden;
        }
        h1, h2, h3, h4, h5 {
            white-space: normal !important;
            word-break: keep-all !important;
            line-height: 1.22 !important;
            margin-bottom: 0.35rem !important;
        }
        div[data-testid="stHorizontalBlock"] {
            gap: 0.8rem;
        }
        .element-container {
            margin-bottom: 0.25rem;
        }
        .dash-title {
            font-size: 2rem;
            font-weight: 800;
            line-height: 1.25;
            margin: 0 0 0.45rem 0;
            white-space: normal;
            word-break: keep-all;
        }
        .mini-note {
            font-size: 0.82rem;
            color: #666;
            line-height: 1.35;
        }
    </style>
    <div class="dash-title">🦷 치과 상권 분석 대시보드</div>
    """,
    unsafe_allow_html=True
)

with st.sidebar:
    st.header("⚙️ 기준 위치 설정")

    address_input = st.text_input(
        "개원 후보지 주소 입력:",
        "서울특별시 중구 세종대로 110"
    )

    run_btn = st.button("검색")

    if run_btn:
        lat, lon = get_coords_from_address(address_input)

        if lat and lon:
            st.session_state["target_lat"] = lat
            st.session_state["target_lon"] = lon
        else:
            st.error("주소를 찾을 수 없습니다.")

    st.markdown("---")

    radius_input = st.slider(
        "탐색 반경 조절 (m)",
        100,
        2000,
        500,
        100
    )

    st.markdown("---")
    enable_transit = st.checkbox("대중교통 조회", value=False)
    enable_sales = st.checkbox("매출 추정 계산", value=False)
    st.caption("최초 로딩을 빠르게 하려고 무거운 대중교통/매출 파일 조회는 필요할 때만 켜도록 했습니다.")

    st.info("💡 지도 위 원하는 곳을 클릭하면 분석 기준점이 이동합니다.")


def classify_year_group(years):
    try:
        years = int(years)
    except Exception:
        return "정보없음"

    if years <= 0:
        return "정보없음"
    elif years < 3:
        return "3년 미만"
    elif years < 5:
        return "3~5년"
    elif years < 10:
        return "5~10년"
    elif years < 15:
        return "10~15년"
    else:
        return "15년 이상"


def render_graph_summary_vertical(df_dentist_merged):
    st.markdown("### 📈 그래프 요약")

    group_order = [
        "3년 미만",
        "3~5년",
        "5~10년",
        "10~15년",
        "15년 이상",
        "정보없음"
    ]

    st.markdown("##### 📊 치과 업력별 분포")
    if df_dentist_merged.empty:
        st.info("업력 그래프를 표시할 치과 데이터가 없습니다.")
    else:
        df_chart = df_dentist_merged.copy()
        df_chart["업력구간"] = df_chart["업력(년)"].apply(classify_year_group)

        chart_df = (
            df_chart["업력구간"]
            .value_counts()
            .reindex(group_order, fill_value=0)
            .reset_index()
        )
        chart_df.columns = ["업력구간", "치과수"]

        color_scale = alt.Scale(
            domain=group_order,
            range=["red", "orange", "green", "blue", "purple", "gray"]
        )

        year_chart = alt.Chart(chart_df).mark_bar(size=44).encode(
            x=alt.X(
                "업력구간:N",
                sort=group_order,
                axis=alt.Axis(title=None, labelAngle=0)
            ),
            y=alt.Y(
                "치과수:Q",
                axis=alt.Axis(title="치과 수")
            ),
            color=alt.Color(
                "업력구간:N",
                scale=color_scale,
                legend=None
            ),
            tooltip=["업력구간", "치과수"]
        ).properties(
            height=CHART_HEIGHT,
            width=690
        )

        st.altair_chart(year_chart)

    st.markdown("##### 📍 동별 치과 밀집도")
    if not df_dentist_merged.empty and "동이름" in df_dentist_merged.columns:
        dong_counts = df_dentist_merged["동이름"].value_counts().reset_index()
        dong_counts.columns = ["동이름", "치과수"]

        dong_chart = alt.Chart(dong_counts).mark_bar(
            size=24,
            color="#3186cc"
        ).encode(
            x=alt.X(
                "동이름:N",
                axis=alt.Axis(labelAngle=-90, title=None),
                sort="-y"
            ),
            y=alt.Y(
                "치과수:Q",
                axis=alt.Axis(title="치과 개수")
            ),
            tooltip=["동이름", "치과수"]
        ).properties(
            height=CHART_HEIGHT,
            width=690
        )

        st.altair_chart(dong_chart)
    else:
        st.info("동별 밀집도 그래프를 표시할 치과 데이터가 없습니다.")


if not MY_API_KEY:
    st.error("⚠️ 코드 상단의 MY_API_KEY에 소상공인시장진흥공단 API 키를 입력해 주세요.")

else:
    with st.spinner("상권 데이터를 분석 중입니다..."):
        lat = st.session_state["target_lat"]
        lon = st.session_state["target_lon"]

        # 1. 소상공인 API 기반 치과 목록
        df_dentist_all = get_clinics_from_coords(
            lat,
            lon,
            radius_input,
            MY_API_KEY
        )

        # 2. hosp_data.xlsx 기반 개업일자 목록
        df_dentist_hira = get_hira_opening_dates_from_csv(
            lat,
            lon,
            radius_input
        )

        # 3. 대중교통 정보: 최초 로딩 속도를 위해 사용자가 켤 때만 조회
        if enable_transit:
            df_transit = get_transit_info(
                lat,
                lon,
                radius_input
            )
        else:
            df_transit = pd.DataFrame(columns=["구분", "정류장/역명", "위도", "경도", "거리_m"])
            st.session_state["transit_status"] = "비활성화"

        # 매출 추정은 개업정보 병합 후 계산
        sales_summary = None
        sales_detail_df = pd.DataFrame()
        sales_error_message = None

        # ==========================================
        # 7. 개업 정보 병합
        # ==========================================
        df_dentist_merged = df_dentist_all.copy()

        if not df_dentist_merged.empty:
            df_dentist_merged["개업일"] = "정보없음"
            df_dentist_merged["업력(년)"] = 0

            if not df_dentist_hira.empty:

                def find_opening_date(row):
                    temp = df_dentist_hira.copy()

                    temp["match_dist"] = temp.apply(
                        lambda hira_row: calculate_distance(
                            row["위도"],
                            row["경도"],
                            hira_row["위도"],
                            hira_row["경도"]
                        ),
                        axis=1
                    )

                    # 좌표가 50m 이내면 같은 병원으로 간주
                    match_row = temp[temp["match_dist"] <= 50].sort_values(
                        by="match_dist"
                    )

                    if not match_row.empty:
                        estb_dt = match_row.iloc[0]["개업일"]
                        return format_opening_date(estb_dt)

                    return "정보없음", 0

                df_dentist_merged[["개업일", "업력(년)"]] = df_dentist_merged.apply(
                    find_opening_date,
                    axis=1,
                    result_type="expand"
                )

            # 이름 90% 이상 + 위치 10m 이내인 치과는 동일 치과로 간주
            df_dentist_merged = remove_duplicate_clinics(
                df_dentist_merged,
                distance_threshold=10,
                similarity_threshold=0.9
            )

        else:
            df_dentist_merged = pd.DataFrame(
                columns=[
                    "상호명",
                    "개업일",
                    "업력(년)",
                    "지번주소",
                    "위도",
                    "경도",
                    "동이름"
                ]
            )

        # 4. 매출 추정: 최초 로딩 속도를 위해 사용자가 켤 때만 대용량 파일 계산
        if enable_sales:
            sales_summary, sales_detail_df, sales_error_message = estimate_dental_sales(
                lat,
                lon,
                radius_input,
                df_dentist_merged
            )
        else:
            sales_error_message = "매출 추정 계산이 꺼져 있습니다. 왼쪽 사이드바에서 '매출 추정 계산'을 켜면 계산됩니다."

        # ==========================================
        # 8. 화면 배치
        # ==========================================
        col1, col2 = st.columns([1.0, 1.0])

        with col1:
            st.subheader("🗺️ 통합 상권 지도")

            st.markdown(
                "⭐ **기준점** | "
                "➕ **치과 업력** | "
                "🔴 **3년미만** 🟠 **3~5년** 🟢 **5~10년** 🔵 **10~15년** 🟣 **15년이상** ⚪ **기타** | "
                "🚏 버스 🚇 지하철"
            )

            clinic_map = create_map(
                df_dentist_merged,
                df_dentist_hira,
                df_transit,
                lat,
                lon,
                radius_input
            )

            map_data = st_folium(
                clinic_map,
                width=690,
                height=455,
                returned_objects=["last_clicked"]
            )

            if map_data and map_data.get("last_clicked"):
                clicked_lat = map_data["last_clicked"]["lat"]
                clicked_lon = map_data["last_clicked"]["lng"]

                if (
                    st.session_state["target_lat"] != clicked_lat
                    or st.session_state["target_lon"] != clicked_lon
                ):
                    st.session_state["target_lat"] = clicked_lat
                    st.session_state["target_lon"] = clicked_lon
                    st.experimental_rerun()

            # 그래프 요약을 지도 바로 아래에 수직 배치
            render_graph_summary_vertical(df_dentist_merged)

        with col2:
            st.subheader("📊 상권 및 교통 요약")

            # 매출 추정 요약
            st.markdown("##### 💰 치과 매출 추정")

            if sales_summary is None:
                st.info(sales_error_message or "매출 추정 데이터를 불러오지 못했습니다.")
                st.caption("공공데이터포털의 국민건강보험공단_시군구별 진료과목별 진료 정보 CSV를 넣고 사이드바에서 매출 추정 계산을 켜면 자동 계산됩니다.")
            else:
                latest_note = ""
                if sales_summary.get("파일최신기간") and sales_summary.get("파일최신기간") != sales_summary.get("기준기간"):
                    latest_note = f" · 파일 전체 최신: {sales_summary.get('파일최신기간')}"

                st.caption(
                    f"{sales_summary['기준기간']} · {sales_summary['매칭기준']}"
                    f"{latest_note} · {sales_summary.get('계산방식', '상권 단위 추정')} · "
                    "개별 치과 실제 총매출이 아닌 건강보험/상권 단위 추정값"
                )

                st.markdown(
                    f"""
                    <div style='display:grid; grid-template-columns:1fr 1fr; gap:8px; margin:6px 0 10px 0;'>
                        <div style='background:#fff8e8; border:1px solid #f0d39a; border-radius:12px; padding:10px; text-align:center;'>
                            <div style='font-size:12px; font-weight:700;'>시군구 치과 월 추정시장규모 합계</div>
                            <div style='font-size:21px; font-weight:800; margin-top:3px;'>{format_won(sales_summary['시군구치과월추정시장규모합계'])}</div>
                        </div>
                        <div style='background:#eef9ff; border:1px solid #afd8ee; border-radius:12px; padding:10px; text-align:center;'>
                            <div style='font-size:12px; font-weight:700;'>반경 내 1곳당 평균 월 추정매출</div>
                            <div style='font-size:21px; font-weight:800; margin-top:3px;'>{format_won(sales_summary['반경내치과1곳당평균월추정매출'])}</div>
                        </div>
                        <div style='background:#f8f0ff; border:1px solid #d9b9ef; border-radius:12px; padding:10px; text-align:center;'>
                            <div style='font-size:12px; font-weight:700;'>반경 내 월 추정시장규모</div>
                            <div style='font-size:21px; font-weight:800; margin-top:3px;'>{format_won(sales_summary['반경내월추정시장규모'])}</div>
                            <div style='font-size:11px; color:#555;'>치과 {sales_summary['반경내치과수']}곳 기준</div>
                        </div>
                        <div style='background:#f5f5f5; border:1px solid #d6d6d6; border-radius:12px; padding:10px; text-align:center;'>
                            <div style='font-size:12px; font-weight:700;'>최고 / 최저 지역 1곳당</div>
                            <div style='font-size:16px; font-weight:800; margin-top:4px;'>{format_won(sales_summary['최고지역1곳당월추정매출'])}</div>
                            <div style='font-size:16px; font-weight:800;'>{format_won(sales_summary['최저지역1곳당월추정매출'])}</div>
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

                if not sales_detail_df.empty:
                    detail_view = sales_detail_df.copy()
                    show_cols = [c for c in ["기준기간", "시도", "시군구", "시군구전체치과수", "반경내치과수", "시군구치과월추정시장규모", "치과1곳당월추정매출", "반경내월추정시장규모", "진료건수"] if c in detail_view.columns]
                    detail_view = detail_view[show_cols].head(20)
                    for money_col in ["추정매출", "1곳당추정매출", "시군구치과월추정시장규모", "치과1곳당월추정매출", "반경내월추정시장규모"]:
                        if money_col in detail_view.columns:
                            detail_view[money_col] = detail_view[money_col].apply(format_won)
                    st.dataframe(detail_view, height=105)

            # 대중교통 요약
            st.markdown("##### 🚉 인근 대중교통")

            transit_status = st.session_state.get("transit_status", "정상")

            if not enable_transit:
                st.info("대중교통 조회가 꺼져 있습니다. 왼쪽 사이드바에서 '대중교통 조회'를 켜면 표시됩니다.")
            elif df_transit.empty:
                if transit_status == "실패":
                    st.info("대중교통 정보는 현재 불러오지 못했습니다. 로컬 대중교통 파일을 추가하면 안정적으로 표시됩니다.")
                else:
                    st.info("설정하신 반경 내에 대중교통 정보가 없습니다.")

            else:
                if transit_status == "로컬":
                    st.caption("로컬 대중교통 좌표 파일 기준으로 표시 중입니다.")

                bus_df = df_transit[df_transit["구분"].str.contains("버스", na=False)].copy()
                subway_df = df_transit[df_transit["구분"].str.contains("지하철", na=False)].copy()

                bus_cnt = len(bus_df)
                sub_cnt = len(subway_df)

                st.markdown(
                    f"""
                    <div style='display:flex; gap:10px; margin:6px 0 10px 0;'>
                        <div style='flex:1; background:#f1fff1; border:1px solid #b6e2b6; border-radius:12px; padding:12px; text-align:center;'>
                            <div style='font-size:14px; font-weight:700;'>🚏 버스정류장</div>
                            <div style='font-size:27px; font-weight:800; margin-top:3px;'>{bus_cnt}개</div>
                            <div style='font-size:11px; color:#555;'>반경 {radius_input}m 이내</div>
                        </div>
                        <div style='flex:1; background:#f2f6ff; border:1px solid #b8c8f0; border-radius:12px; padding:12px; text-align:center;'>
                            <div style='font-size:14px; font-weight:700;'>🚇 지하철역</div>
                            <div style='font-size:27px; font-weight:800; margin-top:3px;'>{sub_cnt}개</div>
                            <div style='font-size:11px; color:#555;'>반경 {radius_input}m 이내</div>
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

                st.markdown("###### 🚏 버스정류장 목록")
                if bus_df.empty:
                    st.caption("반경 내 버스정류장이 없습니다.")
                else:
                    st.dataframe(
                        bus_df[["정류장/역명"]].rename(columns={"정류장/역명": "버스정류장명"}),
                        height=105
                    )

                st.markdown("###### 🚇 지하철역 목록")
                if subway_df.empty:
                    st.caption("반경 내 지하철역이 없습니다.")
                else:
                    st.dataframe(
                        subway_df[["정류장/역명"]].rename(columns={"정류장/역명": "지하철역명"}),
                        height=105
                    )

            # 경쟁 치과 목록
            st.markdown(
                f"##### 📋 경쟁 치과 세부 목록 ({len(df_dentist_merged)}개)"
            )

            if df_dentist_merged.empty:
                st.warning("설정하신 반경 내에 치과가 없습니다.")

            else:
                df_dentist_merged = df_dentist_merged.sort_values(
                    by="업력(년)",
                    ascending=False
                )

                display_cols = ["상호명", "개업일", "업력(년)"]

                if "지번주소" in df_dentist_merged.columns:
                    display_cols.append("지번주소")

                st.dataframe(
                    df_dentist_merged[display_cols],
                    height=235
                )
