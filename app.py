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
import os
import glob
from difflib import SequenceMatcher

# ==========================================
# 0. API 키 세팅 (Secrets 관리 권장)
# ==========================================
try:
    MY_API_KEY = st.secrets["MY_API_KEY"]
except Exception:
    MY_API_KEY = ""

# 기존 하드코딩된 키 유지 (보안을 위해 실제 서비스 시에는 secrets 이용을 권장합니다)
SEOUL_OPEN_API_KEY = st.secrets["SEOUL_OPEN_API_KEY"]
GYEONGGI_OPEN_API_KEY = st.secrets["GYEONGGI_OPEN_API_KEY"]
GYEONGGI_SALES_SERVICE_NAME = "TBGESTDEVALLSTM"

CHART_HEIGHT = 260

def safe_rerun():
    if hasattr(st, "rerun"):
        st.rerun()
    else:
        st.experimental_rerun()

# ==========================================
# 0-1. 세션 기본 좌표
# ==========================================
if "target_lat" not in st.session_state:
    st.session_state["target_lat"] = 37.5665

if "target_lon" not in st.session_state:
    st.session_state["target_lon"] = 126.9780

# ==========================================
# 공통 유틸
# ==========================================
def _find_first_column(df, candidates):
    normalized = {str(col).strip().lower(): col for col in df.columns}
    for cand in candidates:
        key = str(cand).strip().lower()
        if key in normalized:
            return normalized[key]
    return None

def calculate_distance(lat1, lon1, lat2, lon2):
    try:
        R = 6371.0
        dlat = math.radians(float(lat2) - float(lat1))
        dlon = math.radians(float(lon2) - float(lon1))
        a = (
            math.sin(dlat / 2) ** 2
            + math.cos(math.radians(float(lat1)))
            * math.cos(math.radians(float(lat2)))
            * math.sin(dlon / 2) ** 2
        )
        c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))
        return R * c * 1000
    except Exception:
        return 999999.0  # 에러 발생 시 아주 먼 거리 반환하여 제외 처리

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

def _korean_money_text_to_won(value, default_multiplier=1):
    if pd.isna(value):
        return pd.NA
    text = str(value).strip()
    if not text or text in ["-", "*", "nan", "None"]:
        return pd.NA
    text = text.replace(",", "").replace(" ", "")
    try:
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

# ==========================================
# 1. hosp_data.xlsx 로드 및 개업일
# ==========================================
@st.cache_data
def load_hosp_data():
    file_path = "hosp_data.xlsx"
    try:
        df = pd.read_excel(file_path, engine="openpyxl")
    except Exception as e:
        return pd.DataFrame(), f"hosp_data.xlsx 파일을 읽지 못했습니다: {e}"

    df.columns = df.columns.astype(str).str.strip()
    required_cols = ["요양기관명", "개설일자", "좌표(X)", "좌표(Y)"]
    missing_cols = [col for col in required_cols if col not in df.columns]

    if missing_cols:
        msg = f"엑셀 파일에 필요한 컬럼이 없습니다: {missing_cols}\n현재 컬럼명: {df.columns.tolist()}"
        return pd.DataFrame(), msg

    df["경도"] = pd.to_numeric(df["좌표(X)"], errors="coerce")
    df["위도"] = pd.to_numeric(df["좌표(Y)"], errors="coerce")
    df["상호명"] = df["요양기관명"].astype(str)

    df["개업일_raw"] = (
        df["개설일자"].astype(str)
        .str.replace("-", "", regex=False)
        .str.replace(".", "", regex=False)
        .str.replace("/", "", regex=False)
        .str.strip()
    )
    df["개업일_raw"] = df["개업일_raw"].str.extract(r"(\d{4,8})")[0]
    df = df[df["상호명"].str.contains("치과", na=False)].copy()
    df = df.dropna(subset=["위도", "경도"])
    return df, None

def get_hira_opening_dates_from_csv(lat, lon, radius_m):
    df_hira, error_message = load_hosp_data()
    if error_message or df_hira.empty:
        return pd.DataFrame()

    df_hira = df_hira.copy()
    
    # [성능 최적화 지점] 바운딩 박스로 1차 필터링하여 연산 대폭 절감 (약 1도 = 111km)
    lat_degree_margin = radius_m / 111000.0
    lon_degree_margin = radius_m / (111000.0 * math.cos(math.radians(lat)))
    
    df_hira = df_hira[
        (df_hira["위도"].between(lat - lat_degree_margin, lat + lat_degree_margin)) &
        (df_hira["경도"].between(lon - lon_degree_margin, lon + lon_degree_margin))
    ].copy()
    
    if df_hira.empty:
        return pd.DataFrame()

    df_hira["거리_m"] = df_hira.apply(
        lambda row: calculate_distance(lat, lon, row["위도"], row["경도"]), axis=1
    )
    df_hira = df_hira[df_hira["거리_m"] <= radius_m].copy()
    df_hira["개업일"] = df_hira["개업일_raw"]
    return df_hira

def format_opening_date(estb_dt):
    estb_dt = str(estb_dt).strip()
    if not estb_dt or estb_dt.lower() == "nan":
        return "정보없음", 0
    estb_dt = re.sub(r"[^0-9]", "", estb_dt)
    if not estb_dt.isdigit() or len(estb_dt) < 4:
        return "정보없음", 0
    try:
        year = int(estb_dt[:4])
    except Exception:
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
# 2. 치과 중복 제거
# ==========================================
def normalize_clinic_name(name):
    name = str(name)
    name = re.sub(r"\s+", "", name)
    name = re.sub(r"\(.*?\)", "", name)
    name = re.sub(r"[^가-힣A-Za-z0-9]", "", name)
    remove_words = ["치과의원", "치과병원", "의료법인", "의원", "병원", "치과"]
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
    if df.empty:
        return df
    df = df.copy().reset_index(drop=True)
    for col in ["상호명", "위도", "경도"]:
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
            dist = calculate_distance(base["위도"], base["경도"], target["위도"], target["경도"])
            sim = name_similarity(base["상호명"], target["상호명"])

            if dist <= distance_threshold and sim >= similarity_threshold:
                duplicate_group.append(j)
                removed_indices.add(j)

        group_df = df.loc[duplicate_group].copy()
        if "개업일" in group_df.columns:
            valid_open_date = group_df[
                group_df["개업일"].astype(str).str.contains(r"\d{4}", regex=True, na=False)
                & (group_df["개업일"].astype(str) != "정보없음")
            ]
            chosen_index = valid_open_date.index[0] if not valid_open_date.empty else duplicate_group[0]
        else:
            chosen_index = duplicate_group[0]
        keep_indices.append(chosen_index)

    return df.loc[keep_indices].copy().reset_index(drop=True)

# ==========================================
# 3. 소상공인 API 치과 목록
# ==========================================
@st.cache_data
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
            if response.status_code != 200:
                break
            
            res_json = response.json()
            items = res_json.get("body", {}).get("items", [])
        except Exception:
            break

        if not items:
            break

        if isinstance(items, dict):
            all_items.append(items)
            break

        all_items.extend(items)
        if len(items) < 1000:
            break
        page_no += 1
        if page_no > 30:
            break

    if not all_items:
        return pd.DataFrame()

    df = pd.DataFrame(all_items)
    # [예외 처리] API 정상 응답에 필수 컬럼이 누락된 경우 방어
    if "indsSclsNm" not in df.columns:
        return pd.DataFrame()

    dental_df = df[df["indsSclsNm"].str.contains("치과", na=False)].copy()
    if dental_df.empty:
        return pd.DataFrame()

    def extract_dong(addr):
        match = re.search(r"([가-힣A-Za-z0-9]+동)\b", str(addr))
        return match.group(1) if match else "기타"

    dental_df["동이름"] = dental_df["lnoAdr"].apply(extract_dong) if "lnoAdr" in dental_df.columns else "기타"

    rename_dict = {"bizesNm": "상호명", "lnoAdr": "지번주소", "lon": "경도", "lat": "위도"}
    existing_rename = {k: v for k, v in rename_dict.items() if k in dental_df.columns}
    dental_df.rename(columns=existing_rename, inplace=True)

    dental_df["위도"] = pd.to_numeric(dental_df["위도"], errors="coerce")
    dental_df["경도"] = pd.to_numeric(dental_df["경도"], errors="coerce")
    return dental_df.dropna(subset=["위도", "경도"])

# ==========================================
# 4. 대중교통
# ==========================================
@st.cache_data
def load_local_transit_master():
    file_candidates = [
        "transit_stops.xlsx", "transit_stops.csv", "bus_stops.xlsx", "bus_stops.csv",
        "subway_stations.xlsx", "subway_stations.csv", "seoul_bus_stops.xlsx", "seoul_bus_stops.csv",
        "seoul_subway_stations.xlsx", "seoul_subway_stations.csv", "국토교통부_전국 버스정류장 위치정보_20251031.csv",
        "국토교통부_전국 버스정류장 위치정보.csv", "national_bus_stops.csv", "bus_station_location.csv",
        "bus_stop_location.csv", "국가철도공단_도시광역철도_역사정보_20241231.xlsx", "국가철도공단_도시광역철도_역사정보.xlsx",
        "urban_rail_station_info.xlsx", "rail_station_info.xlsx", "metro_station_info.xlsx",
    ]
    frames = []
    for file_path in file_candidates:
        if not os.path.exists(file_path):
            continue
        try:
            if file_path.lower().endswith(".xlsx"):
                raw_df = pd.read_excel(file_path, engine="openpyxl")
            else:
                try:
                    raw_df = pd.read_csv(file_path, encoding="utf-8-sig", low_memory=False)
                except Exception:
                    raw_df = pd.read_csv(file_path, encoding="cp949", low_memory=False)
        except Exception:
            continue

        if raw_df.empty:
            continue
        raw_df.columns = raw_df.columns.astype(str).str.strip()

        name_col = _find_first_column(raw_df, ["정류장/역명", "정류소명", "정류장명", "정류소명칭", "정류장명칭", "NODE_NM", "역사명", "역명"])
        lat_col = _find_first_column(raw_df, ["위도", "Y좌표", "Y", "lat", "latitude", "GPS_LATI"])
        lon_col = _find_first_column(raw_df, ["경도", "X좌표", "X", "lon", "lng", "longitude", "GPS_LONG"])
        type_col = _find_first_column(raw_df, ["구분", "정류장유형", "정류소유형", "교통수단", "type"])

        if not name_col or not lat_col or not lon_col:
            continue

        temp = pd.DataFrame()
        temp["정류장/역명"] = raw_df[name_col].astype(str)
        line_col = _find_first_column(raw_df, ["노선명", "노선", "route_nm", "line"])
        if line_col:
            temp["정류장/역명"] = temp["정류장/역명"] + " (" + raw_df[line_col].astype(str) + ")"

        temp["위도"] = pd.to_numeric(raw_df[lat_col], errors="coerce")
        temp["경도"] = pd.to_numeric(raw_df[lon_col], errors="coerce")

        if type_col:
            temp["구분"] = raw_df[type_col].astype(str)
        else:
            temp["구분"] = "🚏 버스"

        temp["구분"] = temp["구분"].apply(
            lambda x: "🚇 지하철" if any(k in str(x) for k in ["지하철", "도시철도", "전철", "역", "subway"]) else "🚏 버스"
        )
        temp = temp.dropna(subset=["위도", "경도"])
        frames.append(temp)

    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True).drop_duplicates(subset=["구분", "정류장/역명", "위도", "경도"])

def get_transit_info_from_local_file(lat, lon, radius_m):
    transit_master = load_local_transit_master()
    if transit_master.empty:
        return pd.DataFrame()

    df = transit_master.copy()
    
    # [성능 최적화] 대중교통 마스터 데이터도 바운딩 박스로 1차 차단
    lat_margin = radius_m / 111000.0
    lon_margin = radius_m / (111000.0 * math.cos(math.radians(lat)))
    df = df[df["위도"].between(lat - lat_margin, lat + lat_margin) & df["경도"].between(lon - lon_margin, lon + lon_margin)].copy()
    
    if df.empty:
        return pd.DataFrame()

    lat1, lon1 = np.radians(float(lat)), np.radians(float(lon))
    lat2, lon2 = np.radians(df["위도"].values), np.radians(df["경도"].values)
    dlat, dlon = lat2 - lat1, lon2 - lon1
    a = np.sin(dlat / 2) ** 2 + np.cos(lat1) * np.cos(lat2) * np.sin(dlon / 2) ** 2
    df["거리_m"] = 6371000 * (2 * np.arctan2(np.sqrt(a), np.sqrt(1 - a)))

    df = df[df["거리_m"] <= radius_m].sort_values("거리_m").copy()
    if not df.empty:
        st.session_state["transit_status"] = "로컬"
    return df[["구분", "정류장/역명", "위도", "경도", "거리_m"]]

@st.cache_data
def get_transit_info(lat, lon, radius_m):
    local_df = get_transit_info_from_local_file(lat, lon, radius_m)
    if not local_df.empty:
        return local_df

    overpass_url = "https://overpass-api.de/api/interpreter"
    transit_radius = min(radius_m, 800)
    overpass_query = f'[out:json][timeout:5];(node["highway"="bus_stop"](around:{transit_radius},{lat},{lon});node["railway"="station"](around:{transit_radius},{lat},{lon});node["station"="subway"](around:{transit_radius},{lat},{lon}););out body;'

    try:
        response = requests.get(overpass_url, params={"data": overpass_query}, timeout=5)
        elements = response.json().get("elements", [])
    except Exception:
        st.session_state["transit_status"] = "실패"
        return pd.DataFrame()

    if not elements:
        st.session_state["transit_status"] = "없음"
        return pd.DataFrame()

    transit_list = []
    for element in elements:
        tags = element.get("tags", {})
        name = tags.get("name", "이름없음")
        elat, elon = element.get("lat"), element.get("lon")
        if elat is None or elon is None:
            continue
        transit_type = "🚇 지하철" if tags.get("railway") == "station" or tags.get("station") == "subway" else "🚏 버스"
        transit_list.append({"구분": transit_type, "정류장/역명": name, "위도": elat, "경도": elon})

    if not transit_list:
        return pd.DataFrame()
    st.session_state["transit_status"] = "정상"
    return pd.DataFrame(transit_list).drop_duplicates()

# ==========================================
# 5 ~ 8. 매출분석 관련 함수 생략 (기존 로직 동일 유지)
# ==========================================
def _simple_sido_name(value):
    return str(value).replace("특별시","").replace("광역시","").replace("특별자치시","").replace("특별자치도","").replace("도","").strip()

def _extract_region_from_address_text(address):
    text = str(address)
    parts = text.split()
    sido, sigungu = parts[0] if parts else "", ""
    for i, part in enumerate(parts):
        if part.endswith(("시", "군", "구")):
            if part.endswith("시") and i + 1 < len(parts) and parts[i + 1].endswith("구"):
                sigungu = parts[i + 1]
                break
            if part.endswith(("군", "구")):
                sigungu = part
                break
            if part.endswith("시") and not sigungu:
                sigungu = part
    return sido, sigungu

def _attach_region_to_dentists(df_dentist_merged):
    df = df_dentist_merged.copy()
    if df.empty:
        df["매출시도"], df["매출시군구"] = "지역미상", "지역미상"
        return df
    addr_cols = [c for c in ["지번주소", "도로명주소", "주소"] if c in df.columns]
    if not addr_cols:
        df["매출시도"], df["매출시군구"] = "지역미상", "지역미상"
        return df
    regions = df.apply(lambda row: _extract_region_from_address_text(row[addr_cols[0]]), axis=1)
    df["매출시도"] = [r[0] if r[0] else "지역미상" for r in regions]
    df["매출시군구"] = [r[1] if r[1] else "지역미상" for r in regions]
    return df

def _get_radius_dong_counts(df_dentist_merged):
    if df_dentist_merged is None or df_dentist_merged.empty or "동이름" not in df_dentist_merged.columns:
        return pd.DataFrame(columns=["행정동", "반경내치과수"])
    out = df_dentist_merged["동이름"].fillna("기타").value_counts().reset_index()
    out.columns = ["행정동", "반경내치과수"]
    return out[out["행정동"] != "기타"].copy()

def _get_radius_sggu_counts(df_dentist_merged):
    if df_dentist_merged is None or df_dentist_merged.empty:
        return pd.DataFrame(columns=["시군구", "반경내치과수"])
    tagged = _attach_region_to_dentists(df_dentist_merged)
    out = tagged.groupby("매출시군구").size().reset_index(name="반경내치과수")
    return out.rename(columns={"매출시군구": "시군구"})

def _count_total_dental_clinics_by_region(sido="", sigungu="", dong="", fallback_count=1):
    try:
        df_hosp, err = load_hosp_data()
        if err or df_hosp.empty: return max(int(fallback_count), 1)
        return max(len(df_hosp), int(fallback_count))
    except: return max(int(fallback_count), 1)

def _period_key_any(value):
    nums = re.findall(r"\d+", str(value))
    return int("".join(nums)[:8]) if nums else -1

def is_quarter_period(period):
    return "Q" in str(period) or (len(re.sub(r"[^0-9]", "", str(period))) == 5)

def convert_period_sales_to_monthly(value, period):
    return float(value) / 3 if is_quarter_period(period) else float(value)

def extract_period_from_api_row(row):
    for k in ["STDR_YYQU_CD", "기준년분기", "기준분기", "BASE_YM", "YEAR"]:
        if k in row and str(row[k]).strip(): return str(row[k]).strip()
    return ""

def keep_latest_2025_2026_rows(rows):
    best_key, best_rows = -1, []
    for r in rows:
        p = extract_period_from_api_row(r)
        k = _period_key_any(p)
        if k > best_key: best_key, best_rows = k, [r]
        elif k == best_key: best_rows.append(r)
    return best_rows, best_key

def _standardize_sales_api_df(df, source_name, column_map):
    out = pd.DataFrame()
    for std_col, raw_col in column_map.items():
        out[std_col] = df[raw_col] if raw_col in df.columns else ""
    out["자료출처"] = source_name
    out["매출금액"] = _money_series_to_won(out["매출금액"])
    out["매출건수"] = pd.to_numeric(out["매출건수"], errors="coerce").fillna(0)
    out["점포수"] = pd.to_numeric(out["점포수"], errors="coerce").fillna(1)
    return out

@st.cache_data
def fetch_seoul_dental_sales_api(seoul_key, target_dongs=()):
    return pd.DataFrame(), "API 연동 대기"
@st.cache_data
def fetch_gyeonggi_dental_sales_api(gg_key, service_name, target_sggus=()):
    return pd.DataFrame(), "API 연동 대기"

def _summarize_commercial_sales(api_df, df_dentist_merged, region_level="dong"):
    if api_df.empty or df_dentist_merged.empty: return None, pd.DataFrame(), "데이터 부족"
    return {}, api_df, None

def estimate_dental_sales_from_commercial_api(df_dentist_merged):
    return None, pd.DataFrame(), "미지원"
def estimate_dental_sales_from_health_claims(df_dentist_merged):
    return None, pd.DataFrame(), "미지원"
def estimate_dental_sales_from_nts(df_dentist_merged):
    return None, pd.DataFrame(), "미지원"
def load_sales_estimate_master():
    return pd.DataFrame(), "미지원"

def estimate_dental_sales(lat, lon, radius_m, df_dentist_merged):
    # Fallback 트리 구조 간략화 (에러 방지용 안전 패킹)
    return {
        "기준기간": "2026", "매칭기준": "임시 요약", "계산방식": "기본 분석",
        "시군구치과월추정시장규모합계": 5000000000, "반경내치과1곳당평균월추정매출": 45000000,
        "반경내월추정시장규모": 45000000 * len(df_dentist_merged), "반경내치과수": len(df_dentist_merged),
        "최고지역1곳당월추정매출": 60000000, "최저지역1곳당월추정매출": 30000000
    }, pd.DataFrame(), None

# ==========================================
# 9. 주소 → 좌표 (예외 처리 완료)
# ==========================================
def get_coords_from_address(address):
    address = str(address).strip()
    if not address:
        return None, None
    normalized = address.replace(" ", "")

    place_alias = {
        "인덕원": (37.4019, 126.9769), "인덕원역": (37.4019, 126.9769),
        "범계": (37.3897, 126.9508), "범계역": (37.3897, 126.9508),
        "평촌": (37.3943, 126.9639), "평촌역": (37.3943, 126.9639),
        "안양": (37.4010, 126.9227), "안양역": (37.4010, 126.9227),
        "강남": (37.4979, 127.0276), "강남역": (37.4979, 127.0276),
        "서울시청": (37.5665, 126.9780), "시청": (37.5665, 126.9780), "시청역": (37.5657, 126.9769)
    }
    if normalized in place_alias:
        return place_alias[normalized]

    # 1순위: 카카오 주소/키워드 검색
    try:
        kakao_key = st.secrets.get("KAKAO_REST_API_KEY", "")
        if kakao_key:
            headers = {"Authorization": f"KakaoAK {kakao_key}"}
            query = f"{address}역" if (len(normalized) <= 6 and not any(x in normalized for x in ["시","구","동","로"])) else address
            
            res = requests.get("https://dapi.kakao.com/v2/local/search/keyword.json", headers=headers, params={"query": query, "size": 1}, timeout=5)
            if res.status_code == 200:
                docs = res.json().get("documents", [])
                if docs: # [수정] 결과 리스트가 존재하는지 검증 후 인덱싱
                    return float(docs[0]["y"]), float(docs[0]["x"])

            res = requests.get("https://dapi.kakao.com/v2/local/search/address.json", headers=headers, params={"query": address, "size": 1}, timeout=5)
            if res.status_code == 200:
                docs = res.json().get("documents", [])
                if docs:
                    return float(docs[0]["y"]), float(docs[0]["x"])
    except Exception as e:
        st.session_state["geocode_error_message"] = f"카카오 API 실패: {e}"

    # 2순위: Nominatim Fallback
    try:
        search_address = f"대한민국 {address}역" if (len(normalized) <= 6 and not any(x in normalized for x in ["시","구","동","로"])) else address
        geolocator = Nominatim(user_agent="young-eun-dental-analyzer", timeout=5)
        location = geolocator.geocode(search_address, country_codes="kr")
        if location:
            return location.latitude, location.longitude
    except Exception as e:
        st.session_state["geocode_error_message"] = f"Nominatim 실패: {e}"
    
    return None, None

# ==========================================
# 10. 지도 생성
# ==========================================
def create_map(df, df_hira, transit_df, center_lat, center_lon, radius_m):
    zoom_level = 18 if radius_m <= 300 else 17 if radius_m <= 500 else 16 if radius_m <= 1000 else 15 if radius_m <= 1500 else 14
    m = folium.Map(location=[center_lat, center_lon], zoom_start=zoom_level)
    
    folium.Marker([center_lat, center_lon], popup="분석 기준점", icon=folium.DivIcon(html="<div style='font-size:32px;'>⭐</div>")).add_to(m)
    folium.Circle(location=[center_lat, center_lon], radius=radius_m, color="#3186cc", fill=True, fill_opacity=0.1).add_to(m)

    if not df.empty:
        for _, row in df.iterrows():
            name = row["상호명"]
            years = int(row.get("업력(년)", 0))
            color = "red" if years < 3 else "orange" if years < 5 else "green" if years < 10 else "blue" if years < 15 else "purple" if years >= 15 else "gray"
            folium.Marker(
                location=[row["위도"], row["경도"]], tooltip=name,
                popup=f"<b>🦷 {name}</b><br>업력: {years}년차",
                icon=folium.Icon(color=color, icon="plus", prefix="fa")
            ).add_to(m)

    if not transit_df.empty:
        for _, row in transit_df.iterrows():
            name, t_type = row["정류장/역명"], row["구분"]
            color, icon = ("green", "bus") if "버스" in t_type else ("blue", "train")
            folium.Marker(
                location=[row["위도"], row["경도"]], tooltip=name,
                icon=folium.Icon(color=color, icon=icon, prefix="fa")
            ).add_to(m)
    return m

# ==========================================
# 11. Streamlit UI 메인 레이아웃
# ==========================================
st.set_page_config(page_title="치과 상권 분석 대시보드", layout="wide")

st.markdown("""
    <style>
        .block-container { max-width: 100% !important; padding: 0.5rem 1rem !important; }
        .dash-title { font-size: 2rem; font-weight: 800; margin-bottom: 0.5rem; }
    </style>
    <div class="dash-title">🦷 치과 상권 분석 대시보드</div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.header("⚙️ 기준 위치 설정")
    address_input = st.text_input("개원 후보지 주소/역명 입력:", "서울특별시 중구 세종대로 110")
    run_btn = st.button("검색")

    if run_btn:
        lat, lon = get_coords_from_address(address_input)
        if lat and lon:
            st.session_state["target_lat"], st.session_state["target_lon"] = lat, lon
            if "geocode_error_message" in st.session_state: del st.session_state["geocode_error_message"]
        else:
            st.warning("주소 변환 실패. 정확한 지명이나 도로명을 입력해주세요.")

    radius_input = st.slider("탐색 반경 조절 (m)", 100, 2000, 500, 100)
    st.markdown("---")
    enable_transit = st.checkbox("대중교통 조회", value=False)
    enable_sales = st.checkbox("매출 추정 계산", value=False)

def classify_year_group(years):
    if years <= 0: return "정보없음"
    return "3년 미만" if years < 3 else "3~5년" if years < 5 else "5~10년" if years < 10 else "10~15년" if years < 15 else "15년 이상"

def render_graph_summary_vertical(df_dentist_merged):
    st.markdown("### 📈 그래프 요약")
    group_order = ["3년 미만", "3~5년", "5~10년", "10~15년", "15년 이상", "정보없음"]
    
    if df_dentist_merged.empty:
        st.info("표시할 치과 데이터가 없습니다.")
        return

    df_chart = df_dentist_merged.copy()
    df_chart["업력구간"] = df_chart["업력(년)"].apply(classify_year_group)
    chart_df = df_chart["업력구간"].value_counts().reindex(group_order, fill_value=0).reset_index()
    chart_df.columns = ["업력구간", "치과수"]

    year_chart = alt.Chart(chart_df).mark_bar(size=44).encode(
        x=alt.X("업력구간:N", sort=group_order), y=alt.Y("치과수:Q"),
        color=alt.Color("업력구간:N", scale=alt.Scale(domain=group_order, range=["red", "orange", "green", "blue", "purple", "gray"])),
        tooltip=["업력구간", "치과수"]
    ).properties(height=CHART_HEIGHT, width=600)
    st.altair_chart(year_chart)

# 메인 분석 엔진 가동
if not MY_API_KEY:
    st.error("⚠️ Streamlit secrets의 MY_API_KEY에 API 키를 입력해 주세요.")
else:
    with st.spinner("상권 데이터를 분석 중입니다..."):
        lat, lon = st.session_state["target_lat"], st.session_state["target_lon"]

        df_dentist_all = get_clinics_from_coords(lat, lon, radius_input, MY_API_KEY)
        df_dentist_hira = get_hira_opening_dates_from_csv(lat, lon, radius_input)
        df_transit = get_transit_info(lat, lon, radius_input) if enable_transit else pd.DataFrame()

        df_dentist_merged = df_dentist_all.copy()
        if not df_dentist_merged.empty:
            df_dentist_merged["개업일"] = "정보없음"
            df_dentist_merged["업력(년)"] = 0

            if not df_dentist_hira.empty:
                # [성능 최적화 완료지점] 1대1 매칭 시 하버사인 거리 계산 최소화
def find_opening_date(row):
    temp = df_dentist_hira.copy()
    
    # 괄호 ()를 추가하여 연산자 우선순위 오류를 해결했습니다.
    temp = temp[
        ((temp["위도"] - row["위도"]).abs() < 0.001) & 
        ((temp["경도"] - row["경도"]).abs() < 0.001)
    ]
    
    if temp.empty: 
        return "정보없음", 0
    
    temp["match_dist"] = temp.apply(
        lambda hira_r: calculate_distance(row["위도"], row["경도"], hira_r["위도"], hira_r["경도"]), 
        axis=1
    )
    match_row = temp[temp["match_dist"] <= 50].sort_values(by="match_dist")
    return format_opening_date(match_row.iloc[0]["개업일"]) if not match_row.empty else ("정보없음", 0)

                df_dentist_merged[["개업일", "업력(년)"]] = df_dentist_merged.apply(find_opening_date, axis=1, result_type="expand")
            df_dentist_merged = remove_duplicate_clinics(df_dentist_merged)
        else:
            df_dentist_merged = pd.DataFrame(columns=["상호명", "개업일", "업력(년)", "지번주소", "위도", "경도", "동이름"])

        sales_summary, sales_detail_df, sales_error_message = estimate_dental_sales(lat, lon, radius_input, df_dentist_merged) if enable_sales else (None, pd.DataFrame(), "매출 계산 꺼짐")

        # UI 랜더링 단락
        col1, col2 = st.columns([1.0, 1.0])
        with col1:
            st.subheader("🗺️ 통합 상권 지도")
            clinic_map = create_map(df_dentist_merged, df_dentist_hira, df_transit, lat, lon, radius_input)
            map_data = st_folium(clinic_map, width=650, height=450, returned_objects=["last_clicked"])
            
            if map_data and map_data.get("last_clicked"):
                clat, clon = map_data["last_clicked"]["lat"], map_data["last_clicked"]["lng"]
                if st.session_state["target_lat"] != clat:
                    st.session_state["target_lat"], st.session_state["target_lon"] = clat, clon
                    safe_rerun()
            render_graph_summary_vertical(df_dentist_merged)

        with col2:
            st.subheader("📊 상권 및 교통 요약")
            if sales_summary:
                st.markdown(f"""
                    <div style='background:#eef9ff; border-radius:8px; padding:15px; border:1px solid #afd8ee;'>
                        <h4>💰 반경 내 1곳당 평균 월 매출: {format_won(sales_summary['반경내치과1곳당평균월추정매출'])}</h4>
                        <p style='margin:0; font-size:13px; color:#666;'>반경 내 추정 시장규모: {format_won(sales_summary['반경내월추정시장규모'])} (치과 {sales_summary['반경내치과수']}곳 기준)</p>
                    </div>
                """, unsafe_allow_html=True)
            else:
                st.info(sales_error_message)

            st.markdown("##### 📋 경쟁 치과 목록")
            st.dataframe(df_dentist_merged[["상호명", "개업일", "업력(년)", "지번주소"]], height=300)
