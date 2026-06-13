import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import folium
from streamlit_folium import st_folium
from pathlib import Path

st.set_page_config(
    page_title="My Flight Log",
    page_icon="✈️",
    layout="wide"
)

DATA_PATH = Path("data/flight_data.xlsx")

@st.cache_data
def load_data():
    flights = pd.read_excel(DATA_PATH, sheet_name="FlightLog")
    airports = pd.read_excel(DATA_PATH, sheet_name="AirportMaster")
    flights["Date"] = pd.to_datetime(flights["Date"])
    flights["Year"] = flights["Date"].dt.year
    flights["Month"] = flights["Date"].dt.month
    flights["Weekday"] = flights["Date"].dt.day_name()
    return flights, airports

if not DATA_PATH.exists():
    st.error("找不到 `data/flight_data.xlsx`，請確認檔案已放入 data 資料夾。")
    st.stop()

flights, airports = load_data()

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("## ✈️ My Flight Log")
st.caption("個人飛行紀錄 · 航線地圖 · 統計分析")
st.divider()

# ── Year filter ───────────────────────────────────────────────────────────────
years = sorted(flights["Year"].unique().tolist())
year_options = ["全部"] + [str(y) for y in years]
selected_year = st.sidebar.selectbox("篩選年份", year_options, index=0)

if selected_year != "全部":
    df = flights[flights["Year"] == int(selected_year)].copy()
else:
    df = flights.copy()

# ── Metrics ───────────────────────────────────────────────────────────────────
total_flights = len(df)
total_km = int(df["DistanceKm"].sum())
total_hours = round(df["DurationHours"].sum(), 1)
total_co2 = round(total_km * 0.000121, 2)  # rough ICAO factor (tons)

c1, c2, c3, c4 = st.columns(4)
c1.metric("✈️ 航班數", total_flights)
c2.metric("📏 總距離", f"{total_km:,} km")
c3.metric("⏱️ 總飛行時間", f"{total_hours} 小時")
c4.metric("🌿 CO₂ 排放", f"{total_co2} 噸")

st.divider()

# ── Map ───────────────────────────────────────────────────────────────────────
st.subheader("🗺️ 航線地圖")

airport_coords = airports.set_index("IATA")[["Latitude", "Longitude", "City"]].to_dict("index")

m = folium.Map(location=[30, 130], zoom_start=4, tiles="CartoDB dark_matter")

plotted_airports = set()
for _, row in df.iterrows():
    src = row["FromIATA"]
    dst = row["ToIATA"]
    if src in airport_coords and dst in airport_coords:
        s = airport_coords[src]
        d = airport_coords[dst]
        folium.PolyLine(
            locations=[[s["Latitude"], s["Longitude"]], [d["Latitude"], d["Longitude"]]],
            color="#378ADD", weight=2, opacity=0.7,
            tooltip=f"{row['Route']}  {row['FlightNo']}"
        ).add_to(m)
        for iata, info in [(src, s), (dst, d)]:
            if iata not in plotted_airports:
                folium.CircleMarker(
                    location=[info["Latitude"], info["Longitude"]],
                    radius=6, color="#E24B4A", fill=True, fill_color="#E24B4A",
                    fill_opacity=0.9,
                    tooltip=f"{iata} · {info['City']}"
                ).add_to(m)
                plotted_airports.add(iata)

st_folium(m, width="100%", height=360, returned_objects=[])

st.divider()

# ── Charts ────────────────────────────────────────────────────────────────────
col_l, col_r = st.columns(2)

with col_l:
    st.subheader("📅 每年航班數")
    year_counts = flights.groupby("Year").size().reset_index(name="航班數")
    fig_year = px.bar(
        year_counts, x="Year", y="航班數",
        color_discrete_sequence=["#378ADD"],
        text="航班數"
    )
    fig_year.update_layout(
        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(tickmode="array", tickvals=year_counts["Year"].tolist(), title=""),
        yaxis=dict(title="", gridcolor="rgba(128,128,128,0.15)"),
        margin=dict(t=10, b=10, l=0, r=0), height=220
    )
    fig_year.update_traces(textposition="outside")
    st.plotly_chart(fig_year, use_container_width=True)

with col_r:
    st.subheader("📆 每月航班數")
    month_counts = df.groupby("Month").size().reindex(range(1, 13), fill_value=0).reset_index()
    month_counts.columns = ["月份", "航班數"]
    month_labels = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
    month_counts["月份名"] = month_counts["月份"].apply(lambda x: month_labels[x-1])
    fig_month = px.line(
        month_counts, x="月份名", y="航班數",
        markers=True, color_discrete_sequence=["#378ADD"]
    )
    fig_month.update_layout(
        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(title="", categoryorder="array", categoryarray=month_labels),
        yaxis=dict(title="", gridcolor="rgba(128,128,128,0.15)"),
        margin=dict(t=10, b=10, l=0, r=0), height=220
    )
    st.plotly_chart(fig_month, use_container_width=True)

# ── Top stats ─────────────────────────────────────────────────────────────────
st.divider()
col_a, col_b, col_c = st.columns(3)

with col_a:
    st.subheader("🏆 機場")
    ap_from = df["FromIATA"].value_counts()
    ap_to = df["ToIATA"].value_counts()
    ap_total = (ap_from.add(ap_to, fill_value=0)).astype(int).sort_values(ascending=False).head(6)
    fig_ap = px.bar(
        x=ap_total.values, y=ap_total.index,
        orientation="h", color_discrete_sequence=["#378ADD"],
        text=ap_total.values
    )
    fig_ap.update_layout(
        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(title="", gridcolor="rgba(128,128,128,0.15)"),
        yaxis=dict(title="", autorange="reversed"),
        margin=dict(t=5, b=5, l=0, r=0), height=220
    )
    fig_ap.update_traces(textposition="outside")
    st.plotly_chart(fig_ap, use_container_width=True)

with col_b:
    st.subheader("✈️ 航空公司")
    al_counts = df["AirlineIATA"].value_counts().head(5)
    al_names = df.drop_duplicates("AirlineIATA").set_index("AirlineIATA")["Airline"]
    labels = [f"{k} · {al_names.get(k, k)}" for k in al_counts.index]
    fig_al = px.bar(
        x=al_counts.values, y=labels,
        orientation="h", color_discrete_sequence=["#185FA5"],
        text=al_counts.values
    )
    fig_al.update_layout(
        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(title="", gridcolor="rgba(128,128,128,0.15)"),
        yaxis=dict(title="", autorange="reversed"),
        margin=dict(t=5, b=5, l=0, r=0), height=220
    )
    fig_al.update_traces(textposition="outside")
    st.plotly_chart(fig_al, use_container_width=True)

with col_c:
    st.subheader("🛩️ 機型")
    ac_counts = df["AircraftCode"].value_counts().head(5)
    ac_names = df.drop_duplicates("AircraftCode").set_index("AircraftCode")["Aircraft"]
    ac_labels = [f"{k} · {ac_names.get(k,k).split()[1] if len(ac_names.get(k,k).split())>1 else k}" for k in ac_counts.index]
    fig_ac = px.bar(
        x=ac_counts.values, y=ac_labels,
        orientation="h", color_discrete_sequence=["#534AB7"],
        text=ac_counts.values
    )
    fig_ac.update_layout(
        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(title="", gridcolor="rgba(128,128,128,0.15)"),
        yaxis=dict(title="", autorange="reversed"),
        margin=dict(t=5, b=5, l=0, r=0), height=220
    )
    fig_ac.update_traces(textposition="outside")
    st.plotly_chart(fig_ac, use_container_width=True)

# ── Flight log table ──────────────────────────────────────────────────────────
st.divider()
st.subheader("📋 飛行紀錄")

display_cols = {
    "Date": "日期",
    "FlightNo": "班號",
    "Route": "航線",
    "Airline": "航空公司",
    "Aircraft": "機型",
    "Registration": "機號",
    "SeatNo": "座位",
    "DurationHours": "飛行時數",
    "DistanceKm": "距離 (km)"
}

table = df[list(display_cols.keys())].copy()
table.rename(columns=display_cols, inplace=True)
table["日期"] = table["日期"].dt.strftime("%Y-%m-%d")
table["飛行時數"] = table["飛行時數"].apply(lambda x: f"{x:.1f}h")

st.dataframe(table, use_container_width=True, hide_index=True)