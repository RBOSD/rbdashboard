# -*- coding: utf-8 -*-
"""
鐵道事故事件動態分析看板 - Python 版
資料來源：Excel (tra_test_data.xlsx)
"""

import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import os

# ========== 設定 ==========
EXCEL_PATH = r"C:\Users\hckuo\Desktop\Power BI\tra_test_data.xlsx"
OP_MAP = {"TRA": "臺鐵", "THSR": "高鐵", "AFR": "林鐵", "TSR": "糖鐵"}
CLS_COLORS = {"重大行車事故": "#E02F44", "一般行車事故": "#FF9830", "異常事件": "#32CEE6"}
OP_COLORS = {"臺鐵": "#42A5F5", "高鐵": "#FF7043", "林鐵": "#66BB6A", "糖鐵": "#D4E157"}

# ========== 載入資料 ==========
@st.cache_data
def load_data(path):
    """從 Excel 載入並整理資料（欄位：項次、發生時間、營運機構、事故事件分類、事故事件種類、原因第一階～第三階）"""
    df = pd.read_excel(path, sheet_name=0)
    # 發生時間轉 datetime（支援 Excel 序列值與 datetime）
    if df["發生時間"].dtype in ("float64", "int64"):
        df["發生時間"] = pd.to_datetime(df["發生時間"], unit="D", origin="1899-12-30", errors="coerce")
    else:
        df["發生時間"] = pd.to_datetime(df["發生時間"], errors="coerce")
    df = df.dropna(subset=["發生時間"])
    # 年月
    df["年月"] = df["發生時間"].dt.strftime("%Y/%m")
    # 營運機構中文
    df["營運機構名稱"] = df["營運機構"].map(lambda x: OP_MAP.get(str(x).upper(), str(x)))
    # 件數（每筆 1 件）
    df["件數"] = 1
    return df

# ========== 主程式 ==========
st.set_page_config(page_title="鐵道事故事件監理中心", page_icon="🚆", layout="wide")
st.markdown("""
<style>
    .stMetric { background: #181B1F; padding: 15px; border-radius: 6px; border: 1px solid #2B2D37; }
    div[data-testid="stMetricValue"] { font-size: 2rem !important; }
</style>
""", unsafe_allow_html=True)

# 側邊欄：資料路徑與篩選
with st.sidebar:
    st.header("📂 資料與篩選")
    excel_path = st.text_input("Excel 路徑", value=EXCEL_PATH, help="請輸入 tra_test_data.xlsx 的完整路徑")
    
    if not os.path.exists(excel_path):
        st.error(f"找不到檔案：{excel_path}")
        st.stop()
    
    df_raw = load_data(excel_path)
    
    st.subheader("篩選條件")
    op_filter = st.multiselect("營運機構", options=df_raw["營運機構名稱"].unique().tolist(), default=df_raw["營運機構名稱"].unique().tolist())
    cls_filter = st.multiselect("事故事件分類", options=df_raw["事故事件分類"].unique().tolist(), default=df_raw["事故事件分類"].unique().tolist())
    st.caption(f"原始資料共 {len(df_raw)} 筆")

# 套用篩選
df = df_raw[
    (df_raw["營運機構名稱"].isin(op_filter if op_filter else df_raw["營運機構名稱"].unique())) &
    (df_raw["事故事件分類"].isin(cls_filter if cls_filter else df_raw["事故事件分類"].unique()))
].copy()

# ========== 頂部指標卡 ==========
st.title("🚆 鐵道事故事件監理中心")
st.caption("資料來源：Excel | 點擊圖表可互動")

stats_total = df["件數"].sum()
stats_major = df[df["事故事件分類"] == "重大行車事故"]["件數"].sum()
stats_general = df[df["事故事件分類"] == "一般行車事故"]["件數"].sum()
stats_anomaly = df[df["事故事件分類"] == "異常事件"]["件數"].sum()

col1, col2, col3, col4 = st.columns(4)
with col1:
    st.metric("當前篩選累計件數", f"{stats_total:,.0f}")
with col2:
    st.metric("重大行車事故", f"{stats_major:,.0f}")
with col3:
    st.metric("一般行車事故", f"{stats_general:,.0f}")
with col4:
    st.metric("異常事件", f"{stats_anomaly:,.0f}")

# ========== 第一列：趨勢圖 + 圓餅圖 ==========
row1_col1, row1_col2, row1_col3 = st.columns([2, 1, 1])

with row1_col1:
    st.subheader("事故發生趨勢")
    trend = df.groupby(["年月", "事故事件分類"])["件數"].sum().reset_index()
    fig_trend = px.area(trend, x="年月", y="件數", color="事故事件分類", color_discrete_map=CLS_COLORS)
    fig_trend.update_layout(height=280, margin=dict(l=0, r=0, t=30, b=0), legend=dict(orientation="h", y=1.1), template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
    st.plotly_chart(fig_trend, use_container_width=True)

with row1_col2:
    st.subheader("營運機構佔比")
    op_pie = df.groupby("營運機構名稱")["件數"].sum().reset_index()
    fig_op = px.pie(op_pie, values="件數", names="營運機構名稱", color="營運機構名稱", color_discrete_map=OP_COLORS, hole=0.6)
    fig_op.update_layout(height=280, margin=dict(l=0, r=0, t=30, b=0), showlegend=True, legend=dict(orientation="h"), template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)")
    st.plotly_chart(fig_op, use_container_width=True)

with row1_col3:
    st.subheader("事件分類佔比")
    cls_pie = df.groupby("事故事件分類")["件數"].sum().reset_index()
    fig_cls = px.pie(cls_pie, values="件數", names="事故事件分類", color="事故事件分類", color_discrete_map=CLS_COLORS, hole=0.6)
    fig_cls.update_layout(height=280, margin=dict(l=0, r=0, t=30, b=0), showlegend=True, legend=dict(orientation="h"), template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)")
    st.plotly_chart(fig_cls, use_container_width=True)

# ========== 第二列：肇因鑽研 + 熱區矩陣 ==========
row2_col1, row2_col2 = st.columns([1, 1.5])

with row2_col1:
    st.subheader("肇因深度鑽研")
    drill_level = st.radio("選擇階層", ["原因第一階", "原因第二階", "原因第三階"], horizontal=True, key="drill")
    cause_col = drill_level
    cause_df = df.groupby(cause_col)["件數"].sum().reset_index().sort_values("件數", ascending=True)
    fig_cause = px.bar(cause_df, y=cause_col, x="件數", orientation="h", color="件數", color_continuous_scale="Blues")
    fig_cause.update_layout(height=350, margin=dict(l=0, r=0, t=30, b=0), showlegend=False, template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", xaxis=dict(showgrid=False), yaxis=dict(showgrid=False, categoryorder="total ascending"))
    st.plotly_chart(fig_cause, use_container_width=True)

with row2_col2:
    st.subheader("營運機構 vs 肇因 熱區矩陣")
    matrix = df.pivot_table(index=cause_col, columns="營運機構名稱", values="件數", aggfunc="sum", fill_value=0)
    matrix = matrix.reindex(matrix.sum(axis=1).sort_values(ascending=False).index)
    fig_heat = px.imshow(matrix, text_auto=".0f", color_continuous_scale="Reds", aspect="auto")
    fig_heat.update_layout(height=350, margin=dict(l=0, r=0, t=30, b=0), template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", xaxis=dict(side="bottom"))
    st.plotly_chart(fig_heat, use_container_width=True)

st.caption("💡 顏色越紅代表該機構在該肇因下發生的事故件數越密集")
