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
from io import StringIO
import os
from difflib import SequenceMatcher

# ==========================================
# 0. API 키 
MY_API_KEY = st.secrets[ "MY_API_KEY"]


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


# ==========================================
# 4. 대중교통 정보 수집 함수
# ==========================================
@st.cache(suppress_st_warning=True)
def get_transit_info(lat, lon, radius_m):
    """
    대중교통 정보는 Overpass 서버가 자주 느리거나 실패하므로
    앱 전체 속도를 위해 짧게만 시도하고 실패하면 바로 빈 DataFrame 반환
    """

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

st.title("🦷 치과 상권 분석 대시보드")


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

    st.info("💡 **팁:** 지도 위 원하는 곳을 직접 클릭해서 분석 기준점을 이동해 보세요!")


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


if not MY_API_KEY:
    st.error("⚠️ 코드 상단의 MY_API_KEY에 소상공인시장진흥공단 API 키를 입력해 주세요.")

else:
    with st.spinner("해당 반경 내 상권, 개업일자, 교통 정보를 통합 분석 중입니다..."):
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

        # 3. 대중교통 정보
        df_transit = get_transit_info(
            lat,
            lon,
            radius_input
        )

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

        # ==========================================
        # 8. 화면 배치
        # ==========================================
        col1, col2 = st.columns([1.2, 1.0])

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
                width=700,
                height=480,
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

            st.markdown("##### 📊 치과 업력별 분포")

            if df_dentist_merged.empty:
                st.info("업력 그래프를 표시할 치과 데이터가 없습니다.")

            else:
                df_chart = df_dentist_merged.copy()
                df_chart["업력구간"] = df_chart["업력(년)"].apply(classify_year_group)

                group_order = [
                    "3년 미만",
                    "3~5년",
                    "5~10년",
                    "10~15년",
                    "15년 이상",
                    "정보없음"
                ]

                chart_df = (
                    df_chart["업력구간"]
                    .value_counts()
                    .reindex(group_order, fill_value=0)
                    .reset_index()
                )

                chart_df.columns = ["업력구간", "치과수"]

                color_scale = alt.Scale(
                    domain=group_order,
                    range=[
                        "red",
                        "orange",
                        "green",
                        "blue",
                        "purple",
                        "gray"
                    ]
                )

                year_chart = alt.Chart(chart_df).mark_bar(size=45).encode(
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
                    height=220,
                    width=650
                )

                st.altair_chart(year_chart)

        with col2:
            st.subheader("📊 상권 및 교통 요약")

            # 대중교통 요약
            st.markdown("##### 🚉 인근 대중교통")

            transit_status = st.session_state.get("transit_status", "정상")

            if df_transit.empty:
                if transit_status == "실패":
                    st.info("대중교통 정보는 현재 불러오지 못했습니다. 치과 상권 분석은 계속 진행됩니다.")
                else:
                    st.info("설정하신 반경 내에 대중교통 정보가 없습니다.")

            else:
                bus_cnt = len(df_transit[df_transit["구분"].str.contains("버스", na=False)])
                sub_cnt = len(df_transit[df_transit["구분"].str.contains("지하철", na=False)])
                etc_cnt = len(df_transit) - bus_cnt - sub_cnt

                st.caption(
                    f"총 버스정류장: {bus_cnt}개 | 지하철역: {sub_cnt}개 | 기타 교통: {etc_cnt}개"
                )

                st.dataframe(
                    df_transit[["구분", "정류장/역명"]],
                    height=160
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
                    height=250
                )

            # 동별 치과 밀집도
            st.markdown("##### 📍 동별 치과 밀집도")

            if not df_dentist_merged.empty and "동이름" in df_dentist_merged.columns:
                dong_counts = df_dentist_merged["동이름"].value_counts().reset_index()
                dong_counts.columns = ["동이름", "치과수"]

                chart = alt.Chart(dong_counts).mark_bar(
                    size=25,
                    color="#3186cc"
                ).encode(
                    x=alt.X(
                        "동이름:N",
                        axis=alt.Axis(
                            labelAngle=-90,
                            title=None
                        ),
                        sort="-y"
                    ),
                    y=alt.Y(
                        "치과수:Q",
                        axis=alt.Axis(title="치과 개수")
                    ),
                    tooltip=["동이름", "치과수"]
                ).properties(
                    height=260,
                    width=alt.Step(45)
                )

                st.altair_chart(chart)
