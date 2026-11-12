# -*- coding: utf-8 -*-
"""
旅館房價建議工具（Streamlit 原型）
隱私設計：
  - 上傳的檔案只在這次瀏覽的記憶體中處理，程式不寫入任何檔案、資料庫或快取。
  - 估計結果只放在 st.session_state（每個瀏覽分頁各自獨立），關閉分頁後就會被清除。
  - 刻意不使用 st.cache_data / st.cache_resource，避免資料在伺服器上跨使用者保留。
執行：streamlit run streamlit_app.py
"""
import datetime as dt
import hashlib
from pathlib import Path

import pandas as pd
import streamlit as st

import pricing_core as pc

st.set_page_config(page_title="房價建議工具", page_icon="🏨", layout="wide")
SAMPLE = Path(__file__).with_name("sample_data.xlsx")
ss = st.session_state

st.title("🏨 房價建議工具")
st.caption("在整個月平均成交價不低於目標的前提下，讓總營收最大的每日、每房型建議價格。")
st.info("🔒 上傳的資料只在這次瀏覽中暫存在記憶體，不會寫入硬碟或資料庫；"
        "關閉或重新整理網頁後就會清除，下次使用需要重新上傳。", icon="🔒")

# ---------------- 側欄：範例檔與參數 ----------------
with st.sidebar:
    st.header("① 準備資料")
    if SAMPLE.exists():
        st.download_button("下載範例 Excel（模擬資料）", SAMPLE.read_bytes(), file_name="sample_data.xlsx",
                           help="包含格式說明與兩年的模擬資料，可以直接上傳試用。")
    with st.expander("Excel 需要哪些工作表？"):
        st.markdown(
            "- **房型**：房型、房量（可加最低價、最高價）\n"
            "- **歷史時段**：房型、入住日、牌價、開始提前天數、結束提前天數、訂房數、時段開始剩餘房數\n"
            "- **放假日**：日期（國定假日與補假，週六日不用列）\n"
            "- **目前訂房**：入住日、房型、已售房數、已售營收\n\n"
            "詳細說明在範例檔的「說明」工作表。")

    st.header("③ 定價參數")
    today = st.date_input("計算日期（今天）", value=dt.date.today())
    default_month = (today.replace(day=1) + dt.timedelta(days=32)).replace(day=1)
    month_first = st.date_input("要定價的月份（選該月任一天）", value=default_month)
    target = st.number_input("平均成交價目標（元）", min_value=0, value=6000, step=100)
    buffer = st.number_input("安全緩衝（元）", min_value=0, value=100, step=50,
                             help="計算時用「目標 + 緩衝」，降低實際結果略低於目標的機會。")
    step = st.selectbox("價格級距（元）", [50, 100, 200, 500], index=1)

# ---------------- ② 上傳與估計 ----------------
st.header("② 上傳資料並估計需求模型")
up = st.file_uploader("上傳 Excel 檔（.xlsx）", type=["xlsx"])

if up is None:
    for k in ["file_id", "data", "model", "result"]:
        ss.pop(k, None)
    st.stop()

raw = up.getvalue()
file_id = hashlib.sha256(raw).hexdigest()
if ss.get("file_id") != file_id:                   # 換了檔案：清掉舊的結果
    for k in ["data", "model", "result"]:
        ss.pop(k, None)
    try:
        data, warns = pc.read_workbook(up)
    except pc.InputError as e:
        st.error("資料有問題，請修正後重新上傳：\n\n" + str(e).replace("\n", "\n\n"))
        st.stop()
    ss.file_id, ss.data, ss.read_warns = file_id, data, warns

data = ss.data
h = data["hist"]
c1, c2, c3 = st.columns(3)
c1.metric("房型數", len(data["rooms"]))
c2.metric("歷史價格時段", f"{len(h):,}")
c3.metric("放假日", len(data["holidays"]))
st.caption(f"歷史入住日範圍：{h['入住日'].min()} ～ {h['入住日'].max()}")
for w in ss.read_warns:
    st.warning(w)

if "model" not in ss:
    if st.button("估計需求模型", type="primary"):
        bar = st.progress(0.0, text="開始估計…")
        try:
            ss.model = pc.fit_model(data, progress=lambda f, msg: bar.progress(f, text=msg))
        except pc.InputError as e:
            st.error(str(e))
            st.stop()
        bar.empty()
        st.rerun()
    else:
        st.stop()

m = ss.model
price_t, night_t, month_t = pc.model_tables(m)
with st.expander("需求模型估計結果", expanded=False):
    st.markdown(f"價格時段 {m.n_spells:,} 個，其中 {m.cens_rate:.1%} 在時段內賣光（以設限處理）。")
    if m.interact:
        st.markdown(f"週末 × 假日夜交互作用：**顯著**（p = {m.p_int:.2g}），採用四類夜晚類型。")
    elif m.p_int == m.p_int:
        st.markdown(f"週末 × 假日夜交互作用：不顯著（p = {m.p_int:.2g}），採用「週末 + 假日夜」。")
    st.subheader("價格效果")
    st.dataframe(price_t, hide_index=True, use_container_width=True)
    a, b = st.columns(2)
    a.subheader("夜晚類型")
    a.dataframe(night_t, hide_index=True, use_container_width=True)
    b.subheader("月份")
    b.dataframe(month_t, hide_index=True, use_container_width=True)
for n in m.notes:
    st.warning(n)

# ---------------- ④ 產生建議價格 ----------------
st.header("④ 產生建議價格")
rooms = data["rooms"]
bounds = {}
if {"最低價", "最高價"} <= set(rooms.columns):
    bounds = {r["房型"]: (r["最低價"], r["最高價"]) for _, r in rooms.iterrows()}
year, month = month_first.year, month_first.month
st.markdown(f"定價月份：**{year} 年 {month} 月**，計算日期：**{today}**，"
            f"計算用目標：**{target + buffer:,} 元**（目標 {target:,} + 緩衝 {buffer:,}）")

if st.button("產生建議價格", type="primary"):
    try:
        ss.result = pc.price_month(m, data["holidays"], year, month, today, data["otb"],
                                   target, buffer, step, bounds)
    except pc.InputError as e:
        st.error(str(e))
        ss.pop("result", None)

if "result" in ss:
    res = ss.result
    s = res["summary"]
    fmt = lambda v: "—" if v is None else f"{v:,.0f}"
    k1, k2 = st.columns(2)
    k3, k4 = st.columns(2)
    k1.metric("已成交", f"{s['已成交房數']:,} 間", f"平均 {fmt(s['已成交平均價'])} 元", delta_color="off")
    k2.metric("預期全月平均價", f"{fmt(s['預期全月平均價'])} 元",
              f"不設限時 {fmt(s['不設限時預期平均價'])} 元", delta_color="off")
    k3.metric("預期全月營收", f"{fmt(s['預期全月營收'])} 元",
              f"{s['預期全月營收'] - s['不設限時預期營收']:+,.0f} 元（相對不設限）", delta_color="off")
    k4.metric("θ（虛擬成本）", f"{fmt(s['θ（每間房的虛擬成本）'])} 元",
              "目標吃緊" if s["θ（每間房的虛擬成本）"] > 0 else "目標沒有發揮作用", delta_color="off")
    for w in res["warnings"]:
        st.warning(w)

    st.subheader("建議價格（元）")
    st.dataframe(res["price_table"], hide_index=True, use_container_width=True, height=600)
    with st.expander("預測明細"):
        st.dataframe(res["detail"], hide_index=True, use_container_width=True)

    st.download_button("下載結果 Excel", pc.to_excel(res, m),
                       file_name=f"建議價格_{res['month']}_{res['today']}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       type="primary")
    st.caption("建議價格是「從今天到下次重算前」要掛的價格。建議每週（或更頻繁）用最新的訂房資料重新上傳計算。")
