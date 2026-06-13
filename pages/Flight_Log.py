import math
import base64
from html import escape
from pathlib import Path

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

try:
    import plotly.graph_objects as go
except Exception:
    go = None


st.set_page_config(page_title="個人飛行履歷", page_icon="✈️", layout="wide", initial_sidebar_state="collapsed")


# =========================================================
# 基本設定
# =========================================================
APP_DIR = Path(__file__).resolve().parent
BASE_DIR = APP_DIR.parent if APP_DIR.name.lower() == "pages" else APP_DIR
DATA_CANDIDATES = [
    BASE_DIR / "data" / "flight_data.xlsx",
    BASE_DIR / "flight_data.xlsx",
    APP_DIR / "flight_data.xlsx",
    Path.cwd() / "data" / "flight_data.xlsx",
    Path.cwd() / "flight_data.xlsx",
]

# 使用者提供的人頭圖像；若 data/profile_avatar.png 存在，會優先使用外部檔案。
PROFILE_AVATAR_FALLBACK_DATA_URI = "data:image/png;base64,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"
PROFILE_AVATAR_CANDIDATES = [
    BASE_DIR / "data" / "profile_avatar.png",
    BASE_DIR / "profile_avatar.png",
    APP_DIR / "profile_avatar.png",
    Path.cwd() / "data" / "profile_avatar.png",
    Path.cwd() / "profile_avatar.png",
]

AIRPORT_FALLBACK = {
    "TPE": {"IATA": "TPE", "ICAO": "RCTP", "City": "Taipei", "AirportName": "Taoyuan", "Country": "Taiwan", "CountryCode": "TW", "Continent": "Asia", "Latitude": 25.0777, "Longitude": 121.2328},
    "TSA": {"IATA": "TSA", "ICAO": "RCSS", "City": "Taipei", "AirportName": "Taipei Songshan Airport", "Country": "Taiwan", "CountryCode": "TW", "Continent": "Asia", "Latitude": 25.0697, "Longitude": 121.5525},
    "HND": {"IATA": "HND", "ICAO": "RJTT", "City": "Tokyo", "AirportName": "Haneda", "Country": "Japan", "CountryCode": "JP", "Continent": "Asia", "Latitude": 35.5494, "Longitude": 139.7798},
    "HIJ": {"IATA": "HIJ", "ICAO": "RJOA", "City": "Hiroshima", "AirportName": "Hiroshima", "Country": "Japan", "CountryCode": "JP", "Continent": "Asia", "Latitude": 34.4361, "Longitude": 132.9194},
    "KIX": {"IATA": "KIX", "ICAO": "RJBB", "City": "Osaka", "AirportName": "Kansai", "Country": "Japan", "CountryCode": "JP", "Continent": "Asia", "Latitude": 34.4347, "Longitude": 135.2440},
    "ICN": {"IATA": "ICN", "ICAO": "RKSI", "City": "Seoul", "AirportName": "Incheon", "Country": "South Korea", "CountryCode": "KR", "Continent": "Asia", "Latitude": 37.4602, "Longitude": 126.4407},
    "NRT": {"IATA": "NRT", "ICAO": "RJAA", "City": "Tokyo", "AirportName": "Narita", "Country": "Japan", "CountryCode": "JP", "Continent": "Asia", "Latitude": 35.7720, "Longitude": 140.3929},
    "HKG": {"IATA": "HKG", "ICAO": "VHHH", "City": "Hong Kong", "AirportName": "Hong Kong International", "Country": "Hong Kong", "CountryCode": "HK", "Continent": "Asia", "Latitude": 22.3080, "Longitude": 113.9185},
    "BKK": {"IATA": "BKK", "ICAO": "VTBS", "City": "Bangkok", "AirportName": "Suvarnabhumi", "Country": "Thailand", "CountryCode": "TH", "Continent": "Asia", "Latitude": 13.6900, "Longitude": 100.7501},
}

CONTINENT_BY_COUNTRY_CODE = {
    "TW": "Asia", "JP": "Asia", "KR": "Asia", "HK": "Asia", "TH": "Asia", "CN": "Asia", "SG": "Asia", "MY": "Asia", "VN": "Asia", "PH": "Asia",
    "US": "N America", "CA": "N America", "MX": "N America",
    "BR": "S America", "AR": "S America", "CL": "S America", "PE": "S America",
    "GB": "Europe", "FR": "Europe", "DE": "Europe", "IT": "Europe", "ES": "Europe", "NL": "Europe", "CH": "Europe",
    "AU": "Oceania", "NZ": "Oceania",
    "ZA": "Africa", "EG": "Africa", "KE": "Africa", "MA": "Africa",
}

CLASS_LABELS = {1: "economy", 2: "economy+", 3: "business", 4: "first", 5: "private"}
SEAT_LABELS = {1: "window", 2: "middle", 3: "aisle"}
REASON_LABELS = {1: "leisure", 2: "business", 3: "crew", 4: "other"}


# =========================================================
# CSS：改成 myFlightradar24 風格
# =========================================================
st.markdown(
    """
<style>
    .block-container {
        max-width: 100% !important;
        padding-top: 0 !important;
        padding-left: 0 !important;
        padding-right: 0 !important;
        padding-bottom: 2.0rem !important;
        background: #f4f5f7;
    }
    div[data-testid="stAppViewContainer"] {
        background: #f4f5f7;
    }
    /* Streamlit 預設 header 是 fixed，會壓到自訂上方導覽列，導致畫面只露出下半截。 */
    header[data-testid="stHeader"],
    div[data-testid="stToolbar"],
    #MainMenu,
    footer {
        display: none !important;
        visibility: hidden !important;
        height: 0 !important;
    }
    .mfr-topbar {
        background: #ffffff;
        border-bottom: 1px solid #dfe3e8;
        min-height: 58px;
        display: flex;
        align-items: center;
        justify-content: flex-start;
        padding: 0 18px;
        margin: 0;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        position: relative;
        z-index: 5;
    }
    .mfr-brand {
        display: flex;
        align-items: center;
        gap: 10px;
        font-size: 25px;
        font-weight: 700;
        color: #2f2f2f;
        letter-spacing: -0.5px;
    }
    .mfr-logo {
        width: 38px;
        height: 38px;
        border-radius: 50%;
        background: #d94030;
        color: white;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 18px;
        font-style: italic;
        font-weight: 800;
    }
    .mfr-menu {
        display: flex;
        gap: 22px;
        align-items: center;
        color: #4a4a4a;
        font-size: 14px;
    }
    .mfr-add {
        color: #1da1f2;
        border: 1px solid #1da1f2;
        border-radius: 18px;
        padding: 6px 16px;
        font-weight: 700;
    }
    .map-wrap {
        background: #fff;
        border: 0;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        margin: 0;
        overflow: hidden;
    }
    .profile-strip {
        display: grid;
        grid-template-columns: 176px repeat(4, 1fr);
        background: #fff;
        margin: 0 0 0 0;
        border: 1px solid #d9dee5;
        border-top: 0;
        min-height: 132px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
    }
    .avatar-card {
        background: #2c2c2c;
        color: white;
        display: flex;
        align-items: stretch;
        justify-content: center;
        overflow: hidden;
    }
    .avatar-image {
        width: 100%;
        height: 100%;
        min-height: 132px;
        object-fit: cover;
        display: block;
    }
    .avatar-circle {
        width: 70px;
        height: 70px;
        border-radius: 50%;
        border: 3px solid #f4f4f4;
        background: radial-gradient(circle at 50% 36%, #eeeeee 0 20%, transparent 21%), linear-gradient(180deg, transparent 0 45%, #eeeeee 46% 100%);
    }
    .user-name {
        font-weight: 700;
        font-size: 14px;
    }
    /* 儀表板右上方年度選單：保留功能，但不再佔用地圖上方空間。 */
    div[data-testid="stSelectbox"] {
        position: relative;
        z-index: 30;
        width: 150px !important;
        margin-left: auto;
        margin-top: 10px;
        margin-bottom: -46px;
    }
    div[data-testid="stSelectbox"] label {
        display: none !important;
    }
    div[data-testid="stSelectbox"] > div {
        background: #ffffff;
    }
    .summary-cell {
        padding: 24px 26px 16px 26px;
        display: flex;
        flex-direction: column;
        justify-content: center;
        min-width: 0;
    }
    .summary-main {
        color: #333333;
        font-size: 38px;
        font-weight: 800;
        line-height: 1.05;
        letter-spacing: -0.8px;
        white-space: nowrap;
    }
    .summary-main span {
        font-size: 24px;
        font-weight: 400;
        color: #666;
        margin-left: 4px;
    }
    .summary-sub {
        margin-top: 6px;
        color: #555;
        font-size: 15px;
        line-height: 1.35;
    }
    .color-rail {
        height: 4px;
        background: linear-gradient(90deg,#77b900 0 10%,#ffc400 10% 20%,#d7443f 20% 30%,#91418e 30% 40%,#009996 40% 50%,#bde5ff 50% 60%,#77b900 60% 70%,#ffc400 70% 80%,#d7443f 80% 90%,#91418e 90% 100%);
        margin: 0 0 20px 0;
    }
    .section-card {
        background: #ffffff;
        border: 1px solid #dfe3e8;
        box-shadow: 0 2px 8px rgba(0,0,0,0.06);
        margin-bottom: 18px;
    }
    .pie-grid {
        display: grid;
        grid-template-columns: repeat(4, 1fr);
        gap: 0;
        padding: 20px 28px 22px 28px;
    }
    .pie-box {
        display: grid;
        grid-template-columns: 96px 1fr;
        gap: 18px;
        align-items: center;
        min-height: 130px;
    }
    .pie-title {
        grid-column: 1 / 3;
        font-weight: 800;
        color: #222;
        font-size: 13px;
        text-transform: uppercase;
        margin-bottom: 10px;
    }
    .fake-pie {
        width: 88px;
        height: 88px;
        border-radius: 50%;
        background: conic-gradient(#292929 var(--deg), #e1e1e1 0deg);
        position: relative;
    }
    .fake-pie::after {
        content: "";
        position: absolute;
        width: 2px;
        height: 42px;
        background: white;
        left: calc(50% - 1px);
        top: 0;
    }
    .pie-list {
        font-size: 14px;
        color: #333;
        line-height: 1.25;
        word-break: break-word;
    }
    .rank-grid {
        display: grid;
        grid-template-columns: repeat(5, 1fr);
        gap: 22px;
        margin: 8px 0 22px 0;
    }
    .rank-card {
        background: #fff;
        box-shadow: 0 2px 10px rgba(0,0,0,0.12);
        min-height: 340px;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
    }
    .rank-head {
        padding: 18px 20px 16px 20px;
        min-height: 226px;
        color: #fff;
    }
    .rank-title {
        text-align: center;
        font-weight: 800;
        font-size: 15px;
        margin-bottom: 22px;
        text-transform: uppercase;
    }
    .rank-row {
        display: grid;
        grid-template-columns: 58px 1fr 26px;
        gap: 9px;
        align-items: center;
        padding: 8px 0;
        border-bottom: 1px solid rgba(255,255,255,0.25);
        font-size: 13px;
    }
    .rank-label { white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .bar-bg { height: 8px; background: rgba(255,255,255,0.25); }
    .bar-fill { height: 8px; background: rgba(255,255,255,0.92); }
    .rank-foot {
        background: #fff;
        padding: 16px 20px 18px 20px;
        color: #333;
        min-height: 82px;
    }
    .rank-number {
        font-size: 34px;
        line-height: 1;
        font-weight: 800;
    }
    .rank-caption {
        margin-top: 4px;
        font-size: 16px;
        color: #444;
    }
    .pill-wrap {
        float: right;
        margin-top: -44px;
        display: flex;
        flex-direction: column;
        gap: 5px;
    }
    .pill {
        background: #b7b7b7;
        color: #fff;
        font-size: 11px;
        font-weight: 800;
        border-radius: 14px;
        padding: 5px 13px;
        text-align: center;
    }
    .pill.alt {
        background: #fff;
        color: #9a9a9a;
        border: 1px solid #b7b7b7;
    }
    .chart-card {
        background: #fff;
        border: 1px solid #dfe3e8;
        box-shadow: 0 2px 8px rgba(0,0,0,0.06);
        padding: 8px 12px 10px 12px;
        margin-bottom: 18px;
    }
    .small-note {
        color: #777;
        font-size: 12px;
    }
    div[data-testid="stMetric"] {
        background: #ffffff;
        border: 1px solid #dfe3e8;
        padding: 16px 18px;
        box-shadow: 0 1px 6px rgba(0,0,0,0.06);
    }

    .topbar-widget-row {
        background: #ffffff;
        border-bottom: 1px solid #dfe3e8;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        padding: 8px 16px 8px 18px;
        min-height: 58px;
    }
    .mfr-brand-inline {
        display: flex;
        align-items: center;
        gap: 10px;
        min-height: 44px;
        font-size: 25px;
        font-weight: 700;
        color: #2f2f2f;
        letter-spacing: -0.5px;
    }
    .map-size-panel-note {
        color: #666;
        font-size: 12px;
        text-align: right;
        margin-top: -4px;
        margin-bottom: 6px;
    }
    /* 目前頁面只有一個主要 button，用它做成地圖右下角的縮放/尺寸控制。 */
    div[data-testid="stButton"] {
        position: relative;
        margin-top: -62px;
        margin-right: 18px;
        z-index: 20;
        display: flex;
        justify-content: flex-end;
        pointer-events: none;
    }
    div[data-testid="stButton"] > button {
        width: 46px !important;
        height: 46px !important;
        border-radius: 50% !important;
        padding: 0 !important;
        background: #ffffff !important;
        border: 1px solid #d5dbe3 !important;
        box-shadow: 0 4px 12px rgba(0,0,0,0.22) !important;
        color: #606975 !important;
        font-size: 23px !important;
        font-weight: 700 !important;
        pointer-events: auto;
    }
    div[data-testid="stButton"] > button:hover {
        background: #f4f7fb !important;
        border-color: #9aa7b6 !important;
        color: #222 !important;
    }

    @media (max-width: 1200px) {
        .profile-strip { grid-template-columns: 1fr 1fr; }
        .avatar-card { grid-column: 1 / 3; padding: 18px 0; }
        .pie-grid { grid-template-columns: 1fr 1fr; }
        .rank-grid { grid-template-columns: 1fr 1fr; }
    }
    @media (max-width: 760px) {
        .mfr-menu { display: none; }
        .profile-strip { grid-template-columns: 1fr; }
        .avatar-card { grid-column: 1; }
        .pie-grid { grid-template-columns: 1fr; }
        .rank-grid { grid-template-columns: 1fr; }
        .summary-main { font-size: 32px; }
    }
</style>
""",
    unsafe_allow_html=True,
)


# =========================================================
# 工具函式
# =========================================================
def find_data_file() -> Path | None:
    for file in DATA_CANDIDATES:
        if file.exists():
            return file
    return None


def get_profile_avatar_data_uri() -> str:
    for file in PROFILE_AVATAR_CANDIDATES:
        if file.exists():
            try:
                suffix = file.suffix.lower().replace('.', '') or 'png'
                mime = 'jpeg' if suffix in {'jpg', 'jpeg'} else 'png'
                return f"data:image/{mime};base64," + base64.b64encode(file.read_bytes()).decode('ascii')
            except Exception:
                pass
    return PROFILE_AVATAR_FALLBACK_DATA_URI


def make_year_options(df: pd.DataFrame):
    years = sorted([int(y) for y in df['Year'].dropna().unique()]) if 'Year' in df.columns else []
    options = ['Select a year']
    for y in years:
        count = int((df['Year'] == y).sum())
        options.append(f"{y} - {count} flights")
    return options


def parse_year_option(option: str):
    if not option or option in {'All', 'Select a year'}:
        return None
    try:
        return int(str(option).split('-')[0].strip())
    except Exception:
        return None


def haversine_km(lat1, lon1, lat2, lon2):
    try:
        lat1, lon1, lat2, lon2 = map(float, [lat1, lon1, lat2, lon2])
    except Exception:
        return None
    radius = 6371.0
    phi1, phi2 = math.radians(lat1), math.radians(lat2)
    d_phi = math.radians(lat2 - lat1)
    d_lambda = math.radians(lon2 - lon1)
    a = math.sin(d_phi / 2) ** 2 + math.cos(phi1) * math.cos(phi2) * math.sin(d_lambda / 2) ** 2
    return round(2 * radius * math.atan2(math.sqrt(a), math.sqrt(1 - a)))


def safe_number(value, digits=0):
    try:
        value = float(value)
        return f"{value:,.{digits}f}"
    except Exception:
        return "-"


def format_hours(hours):
    try:
        total_minutes = int(round(float(hours) * 60))
    except Exception:
        return "0 h 00 min"
    h, m = divmod(total_minutes, 60)
    return f"{h} h {m:02d} min"


def safe_top(series: pd.Series, default="-"):
    s = series.dropna().astype(str).str.strip()
    s = s[s != ""]
    if s.empty:
        return default
    return s.value_counts().index[0]


def normalize_code(value):
    if pd.isna(value):
        return None
    try:
        return int(float(value))
    except Exception:
        return None


def choose_display(series: pd.Series, fallback="-"):
    s = series.dropna().astype(str).str.strip()
    s = s[s != ""]
    return s.iloc[0] if not s.empty else fallback


def add_continent(row, prefix):
    code = row.get(f"{prefix}_CountryCode")
    if pd.notna(code) and str(code).strip():
        return CONTINENT_BY_COUNTRY_CODE.get(str(code).strip().upper(), "Other")
    country = str(row.get(f"{prefix}Country", "")).strip().lower()
    if country in {"taiwan", "japan", "south korea", "hong kong", "thailand", "china"}:
        return "Asia"
    return "Other" if country else None


@st.cache_data(show_spinner=False)
def load_data():
    data_file = find_data_file()
    if data_file is None:
        return None, None, None

    flight = pd.read_excel(data_file, sheet_name="FlightLog")
    try:
        airport = pd.read_excel(data_file, sheet_name="AirportMaster")
    except Exception:
        airport = pd.DataFrame(AIRPORT_FALLBACK.values())

    flight.columns = [str(c).strip() for c in flight.columns]
    airport.columns = [str(c).strip() for c in airport.columns]

    if "Date" in flight.columns:
        raw_date = flight["Date"]
        if pd.api.types.is_numeric_dtype(raw_date):
            flight["Date"] = pd.to_datetime(raw_date, unit="D", origin="1899-12-30", errors="coerce")
        else:
            numeric_date = pd.to_numeric(raw_date, errors="coerce")
            text_date = pd.to_datetime(raw_date, errors="coerce")
            serial_date = pd.to_datetime(numeric_date, unit="D", origin="1899-12-30", errors="coerce")
            flight["Date"] = text_date.fillna(serial_date)
        flight["Year"] = flight["Date"].dt.year
        flight["Month"] = flight["Date"].dt.month
        flight["Weekday"] = flight["Date"].dt.day_name()

    fallback_df = pd.DataFrame(AIRPORT_FALLBACK.values())
    airport = pd.concat([airport, fallback_df], ignore_index=True)
    if "Continent" not in airport.columns:
        airport["Continent"] = airport.get("CountryCode", pd.Series(dtype=str)).map(CONTINENT_BY_COUNTRY_CODE).fillna("Other")
    airport = airport.dropna(subset=["IATA"]).drop_duplicates(subset=["IATA"], keep="first")

    air_cols = ["IATA", "City", "AirportName", "Country", "CountryCode", "Continent", "Latitude", "Longitude"]
    air_cols = [c for c in air_cols if c in airport.columns]
    airport_small = airport[air_cols].copy()

    flight = flight.merge(airport_small.add_prefix("From_"), left_on="FromIATA", right_on="From_IATA", how="left")
    flight = flight.merge(airport_small.add_prefix("To_"), left_on="ToIATA", right_on="To_IATA", how="left")

    for col in ["FromCountry", "ToCountry"]:
        if col not in flight.columns:
            flight[col] = None
    flight["FromCountry"] = flight["FromCountry"].fillna(flight.get("From_Country"))
    flight["ToCountry"] = flight["ToCountry"].fillna(flight.get("To_Country"))

    if "DurationHours" in flight.columns:
        flight["DurationHours"] = pd.to_numeric(flight["DurationHours"], errors="coerce")
    else:
        flight["DurationHours"] = None

    if "DistanceKm" in flight.columns:
        flight["DistanceKm"] = pd.to_numeric(flight["DistanceKm"], errors="coerce")
    else:
        flight["DistanceKm"] = None

    missing_distance = flight["DistanceKm"].isna()
    if missing_distance.any():
        flight.loc[missing_distance, "DistanceKm"] = flight.loc[missing_distance].apply(
            lambda r: haversine_km(r.get("From_Latitude"), r.get("From_Longitude"), r.get("To_Latitude"), r.get("To_Longitude")),
            axis=1,
        )

    if "Route" not in flight.columns:
        flight["Route"] = flight["FromIATA"].astype(str) + " → " + flight["ToIATA"].astype(str)
    flight["RouteDash"] = flight["FromIATA"].astype(str) + "-" + flight["ToIATA"].astype(str)

    flight["ClassLabel"] = flight.get("FlightClassCode", pd.Series(dtype=float)).apply(lambda x: CLASS_LABELS.get(normalize_code(x)))
    flight["SeatLabel"] = flight.get("SeatTypeCode", pd.Series(dtype=float)).apply(lambda x: SEAT_LABELS.get(normalize_code(x)))
    flight["ReasonLabel"] = flight.get("FlightReasonCode", pd.Series(dtype=float)).apply(lambda x: REASON_LABELS.get(normalize_code(x)))

    flight["FromContinent"] = flight.apply(lambda r: r.get("From_Continent") if pd.notna(r.get("From_Continent")) else add_continent(r, "From"), axis=1)
    flight["ToContinent"] = flight.apply(lambda r: r.get("To_Continent") if pd.notna(r.get("To_Continent")) else add_continent(r, "To"), axis=1)

    return flight, airport, data_file


def render_topbar(df: pd.DataFrame):
    years = sorted([int(y) for y in df["Year"].dropna().unique()]) if "Year" in df.columns else []
    year_options = ["Select a year"] + [str(y) for y in years]

    st.markdown('<div class="topbar-widget-row">', unsafe_allow_html=True)
    c_brand, c_year = st.columns([0.86, 0.14], gap="small")
    with c_brand:
        st.markdown(
            '<div class="mfr-brand-inline"><div class="mfr-logo">my</div>'
            '<div>flightlog<span style="color:#777;font-weight:500">24</span></div></div>',
            unsafe_allow_html=True,
        )
    with c_year:
        selected_year = st.selectbox("Select a year", year_options, index=0, label_visibility="collapsed")
    st.markdown('</div>', unsafe_allow_html=True)
    return selected_year


def apply_filters(df: pd.DataFrame, selected_year: str) -> pd.DataFrame:
    result = df.copy()
    if selected_year != "All":
        result = result[result["Year"] == int(selected_year)]

    # 進階條件移到 sidebar，避免佔用地圖上方空間。
    with st.sidebar:
        st.header("進階篩選")
        st.caption("主要年度篩選已移到右上角；這裡只放進階條件。")

        airlines = sorted([x for x in result["Airline"].dropna().unique() if str(x).strip()]) if "Airline" in result.columns else []
        selected_airlines = st.multiselect("航空公司", airlines, default=airlines)

        routes = sorted([x for x in result["Route"].dropna().unique() if str(x).strip()]) if "Route" in result.columns else []
        selected_routes = st.multiselect("航線", routes, default=routes)

        trips = sorted([x for x in result.get("TripName", pd.Series(dtype=str)).dropna().unique() if str(x).strip()])
        selected_trips = st.multiselect("旅程名稱", trips, default=trips)

    if airlines and selected_airlines:
        result = result[result["Airline"].isin(selected_airlines)]
    if routes and selected_routes:
        result = result[result["Route"].isin(selected_routes)]
    if trips and selected_trips:
        result = result[result["TripName"].isin(selected_trips)]

    return result

def build_route_layer_data(df: pd.DataFrame) -> pd.DataFrame:
    required = ["FromIATA", "ToIATA", "Route", "RouteDash", "From_Latitude", "From_Longitude", "To_Latitude", "To_Longitude"]
    for col in required:
        if col not in df.columns:
            return pd.DataFrame()

    map_df = df.dropna(subset=["From_Latitude", "From_Longitude", "To_Latitude", "To_Longitude"]).copy()
    if map_df.empty:
        return map_df

    route_df = (
        map_df.groupby(["FromIATA", "ToIATA", "Route", "RouteDash", "From_Latitude", "From_Longitude", "To_Latitude", "To_Longitude"], as_index=False)
        .agg(
            FlightCount=("FlightNo", "count"),
            DistanceKm=("DistanceKm", "sum"),
            DurationHours=("DurationHours", "sum"),
            Flights=("FlightNo", lambda x: ", ".join([str(v) for v in x if pd.notna(v)])),
        )
    )
    route_df["LineWidth"] = route_df["FlightCount"].apply(lambda x: min(4.5, max(2.0, 1.4 + float(x) * 0.8)))
    return route_df


def map_zoom(lat_range, lon_range):
    max_range = max(float(lat_range), float(lon_range))
    if max_range < 2:
        return 6
    if max_range < 5:
        return 5
    if max_range < 10:
        return 4
    if max_range < 22:
        return 3.4
    if max_range < 45:
        return 2.4
    return 1.4


def show_map(df: pd.DataFrame, height=520):
    route_df = build_route_layer_data(df)
    if route_df.empty:
        st.warning("目前缺少機場座標，無法顯示航線地圖。請在 flight_data.xlsx 的 AirportMaster 工作表補上 IATA、Latitude、Longitude。")
        return

    point_df = pd.concat(
        [
            route_df[["FromIATA", "From_Latitude", "From_Longitude"]].rename(columns={"FromIATA": "IATA", "From_Latitude": "Latitude", "From_Longitude": "Longitude"}),
            route_df[["ToIATA", "To_Latitude", "To_Longitude"]].rename(columns={"ToIATA": "IATA", "To_Latitude": "Latitude", "To_Longitude": "Longitude"}),
        ],
        ignore_index=True,
    ).drop_duplicates("IATA")

    if go is None:
        st.map(point_df.rename(columns={"Latitude": "lat", "Longitude": "lon"})[["lat", "lon"]])
        return

    lats = pd.concat([route_df["From_Latitude"], route_df["To_Latitude"]]).astype(float)
    lons = pd.concat([route_df["From_Longitude"], route_df["To_Longitude"]]).astype(float)
    center = {"lat": float(lats.mean()), "lon": float(lons.mean())}
    zoom = map_zoom(lats.max() - lats.min(), lons.max() - lons.min())

    fig = go.Figure()
    for _, r in route_df.iterrows():
        hover = f"<b>{escape(str(r['RouteDash']))}</b><br>Flights: {int(r['FlightCount'])}<br>{escape(str(r['Flights']))}<br>{r['DistanceKm']:,.0f} km"
        fig.add_trace(
            go.Scattermapbox(
                lon=[r["From_Longitude"], r["To_Longitude"]],
                lat=[r["From_Latitude"], r["To_Latitude"]],
                mode="lines",
                line=dict(width=r["LineWidth"], color="#ff2b2b"),
                opacity=0.88,
                hoverinfo="text",
                text=hover,
                showlegend=False,
            )
        )

    fig.add_trace(
        go.Scattermapbox(
            lon=point_df["Longitude"],
            lat=point_df["Latitude"],
            text=point_df["IATA"],
            mode="markers+text",
            textposition="top center",
            textfont=dict(size=12, color="#333333"),
            marker=dict(size=10, color="#e53935"),
            hovertemplate="<b>%{text}</b><extra></extra>",
            showlegend=False,
        )
    )

    # Use English Google map tiles. If Google blocks third-party tile loading in a given environment,
    # Plotly will still show the route layer on a clean white background instead of breaking the page.
    google_tile = "https://mt1.google.com/vt/lyrs=m&x={x}&y={y}&z={z}&hl=en"
    fig.update_layout(
        height=height,
        autosize=True,
        margin=dict(l=0, r=0, t=0, b=0),
        paper_bgcolor="#ffffff",
        plot_bgcolor="#ffffff",
        mapbox=dict(
            style="white-bg",
            center=center,
            zoom=zoom,
            layers=[
                dict(
                    sourcetype="raster",
                    source=[google_tile],
                    sourceattribution="Google Maps",
                    below="traces",
                )
            ],
        ),
    )

    # 用 components.html 自己包 Plotly 地圖，才能把「全螢幕」與「縮放」控制放在地圖右上角。
    # requestFullscreen 只作用在 map-stage，因此不會把整個 Streamlit 頁面一起放大。
    plot_html = fig.to_html(
        full_html=False,
        include_plotlyjs=True,
        div_id="flight-map-plot",
        config={
            "displayModeBar": False,
            "scrollZoom": True,
            "responsive": True,
        },
        default_width="100%",
        default_height=f"{height}px",
    )
    map_component_html = f"""
    <html>
    <head>
        <style>
            html, body {{ margin:0; padding:0; width:100%; height:100%; overflow:hidden; background:#ffffff; }}
            #map-stage {{ position:relative; width:100%; height:{height}px; background:#ffffff; overflow:hidden; }}
            #flight-map-plot {{ width:100% !important; height:{height}px !important; }}
            .map-controls {{
                position:absolute;
                top:14px;
                right:14px;
                z-index:9999;
                display:flex;
                flex-direction:column;
                gap:8px;
                pointer-events:auto;
            }}
            .map-btn {{
                width:40px;
                height:40px;
                border-radius:50%;
                border:1px solid rgba(0,0,0,0.16);
                background:#ffffff;
                color:#404852;
                font-size:24px;
                font-weight:700;
                line-height:36px;
                text-align:center;
                cursor:pointer;
                box-shadow:0 3px 10px rgba(0,0,0,0.22);
                user-select:none;
            }}
            .map-btn.full {{ font-size:21px; line-height:38px; }}
            .map-btn:hover {{ background:#f4f7fb; color:#111; }}
            #map-stage:fullscreen {{ width:100vw; height:100vh; background:#ffffff; }}
            #map-stage:fullscreen #flight-map-plot {{ width:100vw !important; height:100vh !important; }}
            #map-stage:fullscreen .map-controls {{ top:18px; right:18px; }}
            #map-stage:-webkit-full-screen {{ width:100vw; height:100vh; background:#ffffff; }}
            #map-stage:-webkit-full-screen #flight-map-plot {{ width:100vw !important; height:100vh !important; }}
        </style>
    </head>
    <body>
        <div id="map-stage">
            {plot_html}
            <div class="map-controls" aria-label="Map controls">
                <button id="btn-fullscreen" class="map-btn full" title="Fullscreen">⛶</button>
                <button id="btn-zoom-in" class="map-btn" title="Zoom in">+</button>
                <button id="btn-zoom-out" class="map-btn" title="Zoom out">−</button>
            </div>
        </div>
        <script>
            const stage = document.getElementById('map-stage');
            const plot = document.getElementById('flight-map-plot');
            const baseHeight = {height};

            function currentZoom() {{
                const z = plot && plot.layout && plot.layout.mapbox ? plot.layout.mapbox.zoom : null;
                return (typeof z === 'number') ? z : {zoom};
            }}
            function setZoom(z) {{
                if (!plot || !window.Plotly) return;
                Plotly.relayout(plot, {{'mapbox.zoom': Math.max(0.2, Math.min(18, z))}});
            }}
            document.getElementById('btn-zoom-in').addEventListener('click', () => setZoom(currentZoom() + 0.7));
            document.getElementById('btn-zoom-out').addEventListener('click', () => setZoom(currentZoom() - 0.7));

            async function toggleFullscreen() {{
                try {{
                    if (!document.fullscreenElement && !document.webkitFullscreenElement) {{
                        if (stage.requestFullscreen) await stage.requestFullscreen();
                        else if (stage.webkitRequestFullscreen) await stage.webkitRequestFullscreen();
                    }} else {{
                        if (document.exitFullscreen) await document.exitFullscreen();
                        else if (document.webkitExitFullscreen) await document.webkitExitFullscreen();
                    }}
                }} catch (err) {{
                    console.warn('Fullscreen is not available in this browser context.', err);
                }}
                setTimeout(resizePlot, 180);
            }}
            function resizePlot() {{
                if (!plot || !window.Plotly) return;
                const isFs = !!(document.fullscreenElement || document.webkitFullscreenElement);
                const h = isFs ? window.innerHeight : baseHeight;
                Plotly.relayout(plot, {{height: h}});
                Plotly.Plots.resize(plot);
            }}
            document.getElementById('btn-fullscreen').addEventListener('click', toggleFullscreen);
            document.addEventListener('fullscreenchange', resizePlot);
            document.addEventListener('webkitfullscreenchange', resizePlot);
            window.addEventListener('resize', resizePlot);
        </script>
    </body>
    </html>
    """
    components.html(map_component_html, height=height, scrolling=False)


def count_known(series: pd.Series, ordered_labels):
    s = series.dropna().astype(str).str.strip()
    s = s[s != ""]
    return {label: int((s == label).sum()) for label in ordered_labels}


def pie_box(title, counts: dict):
    total = sum(counts.values())
    first_value = next((v for v in counts.values() if v > 0), 0)
    deg = 0 if total == 0 else max(2, int(round(first_value / total * 360)))
    items = "".join([f"<div><b>{v}</b> {escape(str(k))}</div>" for k, v in counts.items()])
    return f"""
    <div class="pie-box">
        <div class="pie-title">{escape(title)}</div>
        <div class="fake-pie" style="--deg:{deg}deg"></div>
        <div class="pie-list">{items}</div>
    </div>
    """


def summary_html(df: pd.DataFrame):
    total_flights = len(df)
    total_distance_km = df["DistanceKm"].sum(skipna=True)
    total_miles = total_distance_km * 0.621371
    total_hours = df["DurationHours"].sum(skipna=True)
    days = total_hours / 24 if total_hours else 0
    months = days / 30.4375 if days else 0
    moon = total_distance_km / 384400 if total_distance_km else 0
    co2_tons = total_miles * 0.196 / 1000 if total_miles else 0
    methane_kg = total_miles * 0.00001055 if total_miles else 0
    nitrous_kg = total_miles * 0.0000090 if total_miles else 0

    domestic = 0
    international = 0
    for _, r in df.iterrows():
        if pd.notna(r.get("FromCountry")) and pd.notna(r.get("ToCountry")) and str(r.get("FromCountry")) == str(r.get("ToCountry")):
            domestic += 1
        else:
            international += 1

    return f"""
    <div class="profile-strip">
        <div class="avatar-card">
            <img class="avatar-image" src="{get_profile_avatar_data_uri()}" alt="yuhui0427 profile avatar" />
        </div>
        <div class="summary-cell">
            <div class="summary-main">{total_flights:,}<span>flights</span></div>
            <div class="summary-sub"><b>{domestic:,}</b> domestic<br><b>{international:,}</b> international</div>
        </div>
        <div class="summary-cell">
            <div class="summary-main">{safe_number(total_miles)}<span>miles</span></div>
            <div class="summary-sub"><b>{safe_number(total_distance_km)}</b> km<br><b>{moon:.2f}x</b> to the moon</div>
        </div>
        <div class="summary-cell">
            <div class="summary-main">{format_hours(total_hours).replace(' h ', '<span> h </span>').replace(' min', '<span> min</span>')}</div>
            <div class="summary-sub"><b>{days:.1f}</b> days<br><b>{months:.2f}</b> months</div>
        </div>
        <div class="summary-cell">
            <div class="summary-main">{co2_tons:.1f}<span>tons CO₂</span></div>
            <div class="summary-sub"><b>{methane_kg:.2f}</b> kg methane<br><b>{nitrous_kg:.2f}</b> kg nitrous oxide</div>
        </div>
    </div>
    <div class="color-rail"></div>
    """


def pie_section_html(df: pd.DataFrame):
    class_counts = count_known(df["ClassLabel"], ["economy", "economy+", "business", "first", "private"])
    seat_counts = count_known(df["SeatLabel"], ["window", "middle", "aisle"])
    reason_counts = count_known(df["ReasonLabel"], ["leisure", "business", "crew", "other"])
    continents = pd.concat([df["FromContinent"], df["ToContinent"]], ignore_index=True).dropna().astype(str)
    continent_counts = {k: int((continents == k).sum()) for k in ["Asia", "S America", "Oceania", "Europe", "Africa", "N America"]}
    return f"""
    <div class="section-card">
        <div class="pie-grid">
            {pie_box("CLASS", class_counts)}
            {pie_box("SEAT", seat_counts)}
            {pie_box("REASON", reason_counts)}
            {pie_box("CONTINENTS", continent_counts)}
        </div>
    </div>
    """


def top_counts(series: pd.Series, top_n=5):
    s = series.dropna().astype(str).str.strip()
    s = s[s != ""]
    if s.empty:
        return pd.Series(dtype=int)
    return s.value_counts().head(top_n)


def rank_card(title, series, total_count, caption, color, top_n=5, show_pills=True):
    counts = top_counts(series, top_n)
    max_value = int(counts.max()) if not counts.empty else 1
    row_parts = []
    for label, value in counts.items():
        width = int(value / max_value * 100) if max_value else 0
        row_parts.append(
            f'<div class="rank-row">'
            f'<div class="rank-label">{escape(str(label))}</div>'
            f'<div class="bar-bg"><div class="bar-fill" style="width:{width}%"></div></div>'
            f'<div>{int(value)}</div>'
            f'</div>'
        )
    if not row_parts:
        row_parts.append('<div class="rank-row"><div>-</div><div></div><div>0</div></div>')

    pills = ''
    if show_pills:
        pills = (
            '<div class="pill-wrap">'
            '<div class="pill">FLIGHTS</div>'
            '<div class="pill alt">DISTANCE</div>'
            '</div>'
        )

    # Streamlit st.markdown 會先經過 Markdown parser。
    # 若 HTML 中有空白行或大量縮排，會被判定為 code block，畫面就會顯示原始 <div> 文字。
    # 因此這裡刻意輸出成 compact HTML，不保留縮排與空白行。
    rows = ''.join(row_parts)
    return (
        f'<div class="rank-card">'
        f'<div class="rank-head" style="background:{color}">'
        f'<div class="rank-title">{escape(title)}</div>'
        f'{rows}'
        f'</div>'
        f'<div class="rank-foot">'
        f'<div class="rank-number">{int(total_count):,}</div>'
        f'<div class="rank-caption">{escape(caption)}</div>'
        f'{pills}'
        f'</div>'
        f'</div>'
    )

def rank_cards_html(df: pd.DataFrame):
    airports = pd.concat([df["FromIATA"], df["ToIATA"]], ignore_index=True)
    airlines = df.get("AirlineIATA", df["Airline"])
    aircraft = df.get("AircraftCode", df.get("Aircraft", pd.Series(dtype=str)))
    countries = pd.concat([df.get("From_CountryCode", df["FromCountry"]), df.get("To_CountryCode", df["ToCountry"])], ignore_index=True)

    html = "<div class='rank-grid'>"
    html += rank_card("TOP AIRPORTS", airports, airports.dropna().astype(str).str.strip().replace("", pd.NA).dropna().nunique(), "total airports", "#7fba00")
    html += rank_card("TOP AIRLINES", airlines, airlines.dropna().astype(str).str.strip().replace("", pd.NA).dropna().nunique(), "total airlines", "#f7b800")
    html += rank_card("TOP AIRCRAFT", aircraft, aircraft.dropna().astype(str).str.strip().replace("", pd.NA).dropna().nunique(), "total aircraft", "#d8433e")
    html += rank_card("TOP ROUTES", df["RouteDash"], df["RouteDash"].dropna().astype(str).str.strip().replace("", pd.NA).dropna().nunique(), "total routes", "#8e3f8f")
    html += rank_card("TOP COUNTRIES", countries, countries.dropna().astype(str).str.strip().replace("", pd.NA).dropna().nunique(), "total countries", "#009892", show_pills=False)
    html += "</div>"
    return html


def mfr_line_chart(x, y, title, bg_color, y_range_pad=1):
    # 年份、月份、星期都必須強制用 category 軸。
    # 否則 Plotly 會把 2023 / 2024 當連續數字，導致畫面出現 2,023.2、2,023.4，
    # 看起來像沒有顯示完整年度。
    x = [str(v) for v in x]
    if go is None:
        st.line_chart(pd.DataFrame({"value": y}, index=x), use_container_width=True)
        return
    ymax = max(y) if len(y) else 1
    fig = go.Figure()
    fig.add_trace(
        go.Scatter(
            x=x,
            y=y,
            mode="lines+markers",
            line=dict(width=5, color="rgba(255,255,255,0.95)"),
            marker=dict(size=10, color="rgba(255,255,255,1)"),
            fill="tozeroy",
            fillcolor="rgba(255,255,255,0.12)",
            hovertemplate="%{x}: %{y}<extra></extra>",
        )
    )
    fig.update_layout(
        height=260,
        title=dict(text=f"<b>{title}</b>", x=0.5, y=0.92, font=dict(color="white", size=14)),
        margin=dict(l=28, r=28, t=46, b=36),
        paper_bgcolor=bg_color,
        plot_bgcolor=bg_color,
        xaxis=dict(
            type="category",
            categoryorder="array",
            categoryarray=x,
            tickmode="array",
            tickvals=x,
            ticktext=x,
            showgrid=False,
            tickfont=dict(color="white", size=12),
            zeroline=False,
        ),
        yaxis=dict(
            showgrid=True,
            gridcolor="rgba(255,255,255,0.22)",
            tickfont=dict(color="white", size=12),
            zeroline=False,
            rangemode="tozero",
            dtick=1,
            range=[0, max(1, ymax + y_range_pad)],
        ),
        showlegend=False,
    )
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def show_charts(df: pd.DataFrame):
    years = sorted([int(y) for y in df["Year"].dropna().unique()]) if "Year" in df.columns else []
    if years:
        # myFlightradar24 的年度圖會把中間沒有航班的年份也補出來，
        # 因此這裡從第一筆航班年度一路顯示到「目前年份」或資料最新年份。
        current_year = pd.Timestamp.today().year
        end_year = max(max(years), current_year)
        x = list(range(min(years), end_year + 1))
        year_counts = df.groupby("Year").size().to_dict()
        y = [int(year_counts.get(year, 0)) for year in x]
        st.markdown('<div class="chart-card">', unsafe_allow_html=True)
        mfr_line_chart(x, y, "FLIGHTS PER YEAR", "#8fc6ec")
        year_span = max(1, end_year - min(years))
        st.markdown(f"<div style='padding:0 14px 12px 14px'><span style='font-size:34px;font-weight:800'>{year_span}</span><br><span style='color:#555'>years of flying</span></div>", unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    months = list(range(1, 13))
    month_names = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    month_counts = df.groupby("Month").size().to_dict() if "Month" in df.columns else {}
    month_y = [int(month_counts.get(m, 0)) for m in months]

    weekday_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    weekday_short = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
    weekday_counts = df.groupby("Weekday").size().to_dict() if "Weekday" in df.columns else {}
    weekday_y = [int(weekday_counts.get(d, 0)) for d in weekday_order]

    with c1:
        mfr_line_chart(month_names, month_y, "FLIGHTS PER MONTH", "#e8b4d4")
    with c2:
        mfr_line_chart(weekday_short, weekday_y, "FLIGHTS PER WEEKDAY", "#e8b4d4")


def show_detail_table(df: pd.DataFrame):
    st.markdown("### Flight details")
    display_cols = [
        "Date", "FlightNo", "Airline", "Route", "DepTime", "ArrTime", "Duration",
        "DistanceKm", "Aircraft", "Registration", "SeatNo", "TripName", "Companion", "Note",
    ]
    display_cols = [c for c in display_cols if c in df.columns]
    detail = df[display_cols].sort_values("Date", ascending=False).copy()
    if "Date" in detail.columns:
        detail["Date"] = detail["Date"].dt.strftime("%Y-%m-%d")
    if "DistanceKm" in detail.columns:
        detail["DistanceKm"] = detail["DistanceKm"].round(0).astype("Int64")
    st.dataframe(detail, use_container_width=True, hide_index=True)


# =========================================================
# 主畫面
# =========================================================
_df, airport_df, data_file = load_data()

if _df is None:
    st.error("找不到 flight_data.xlsx")
    st.markdown(
        """
請把檔案放在以下其中一個位置：

```text
data/flight_data.xlsx
flight_data.xlsx
pages/flight_data.xlsx
```

建議放在：

```text
ai-monitor-center/
  pages/Flight_Log.py
  data/flight_data.xlsx
```
        """
    )
    st.stop()

# 依使用者需求：刪除頁面上方篩選與 flightlog logo；年度選單放在儀表板右上方。
# Streamlit widget state 在 rerun 開始時已可讀取，因此地圖也會跟著目前選取年度切換。
current_year_option = st.session_state.get("dashboard_year_select", "All")
current_selected_year = parse_year_option(current_year_option)
map_df = _df.copy() if current_selected_year is None else _df[_df["Year"] == current_selected_year].copy()

if map_df.empty:
    st.warning("目前年度篩選沒有航班資料。")
    st.stop()

st.markdown('<div class="map-wrap">', unsafe_allow_html=True)
show_map(map_df, height=620)
st.markdown('</div>', unsafe_allow_html=True)

# 地圖維持滿版；下方 Profile / 統計儀表板置中，左右保留空白。
left_pad, main_area, right_pad = st.columns([0.07, 0.86, 0.07], gap="small")
with main_area:
    year_options = make_year_options(_df)
    default_index = year_options.index(current_year_option) if current_year_option in year_options else 0
    _, year_col = st.columns([0.91, 0.09], gap="small")
    with year_col:
        selected_year_option = st.selectbox("Select a year", year_options, index=default_index, key="dashboard_year_select")
    selected_year = parse_year_option(selected_year_option)
    filtered = _df.copy() if selected_year is None else _df[_df["Year"] == selected_year].copy()

    if filtered.empty:
        st.warning("目前年度篩選沒有航班資料。")
        st.stop()

    st.markdown(summary_html(filtered), unsafe_allow_html=True)
    st.markdown(pie_section_html(filtered), unsafe_allow_html=True)
    st.markdown(rank_cards_html(filtered), unsafe_allow_html=True)

    show_charts(filtered)

    st.markdown("---")
    show_detail_table(filtered)

    if "DataQualityNote" in filtered.columns:
        notes = filtered["DataQualityNote"].dropna()
        notes = notes[notes.astype(str).str.strip() != ""].unique()
        if len(notes) > 0:
            with st.expander("資料品質備註"):
                for note in notes:
                    st.write(f"- {note}")

    with st.expander("資料維護方式"):
        st.markdown(
            f"""
目前讀取資料檔：`{data_file}`

### 新增航班
請直接編輯 `data/flight_data.xlsx` 的 `FlightLog` 工作表。

最重要欄位：

```text
Date, FlightNo, Airline, AirlineIATA, FromIATA, ToIATA, DepTime, ArrTime,
Duration, DurationHours, Aircraft, AircraftCode, Registration, SeatNo,
SeatTypeCode, FlightClassCode, FlightReasonCode, TripName, Companion, Note
```

### 新增機場
如果地圖沒有顯示某條航線，通常是 `AirportMaster` 缺少該機場座標。

請在 `AirportMaster` 新增：

```text
IATA, ICAO, City, AirportName, Country, CountryCode, Latitude, Longitude
```
            """
        )
