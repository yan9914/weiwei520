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
import importlib
from pathlib import Path

import streamlit as st

import pricing_core as pc

# Streamlit Cloud 從 GitHub 更新程式時，只會重跑這個主程式，已載入的 pricing_core 會沿用舊版，
# 導致找不到新函數。每次執行都重新載入，確保用的是最新版（只定義函數，成本很低）。
pc = importlib.reload(pc)

st.set_page_config(page_title="房價建議工具", page_icon="🏨", layout="wide")
SAMPLE = Path(__file__).with_name("sample_data.xlsx")
ss = st.session_state


def as_range(v):
    """date_input 的區間選擇在只點了第一天時會回傳一個日期，這時回傳 None。"""
    return (v[0], v[1]) if isinstance(v, (tuple, list)) and len(v) == 2 else None


st.title("🏨 房價建議工具")
st.caption("產生每日、每房型的建議價格，讓總淨收益（成交價 − 房型成本）最大；也可以加上「平均淨收益不低於目標」的條件。")
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
            "- **房型**：房型、房量、成本（每間夜，元；可加最低價、最高價）\n"
            "- **歷史時段**：房型、入住日、牌價、開始提前天數、結束提前天數、訂房數、時段開始剩餘房數\n"
            "- **放假日**：日期（國定假日與補假，週六日不用列）\n"
            "- **目前訂房**：入住日、房型、已售房數、已售營收\n\n"
            "詳細說明在範例檔的「說明」工作表。")

    st.header("③ 定價參數")
    today = st.date_input("計算日期（今天）", value=dt.date.today())
    # 預設 = 下個月整月；用真實的今天計算，避免改「計算日期」時把已選好的區間重設
    nm = (dt.date.today().replace(day=1) + dt.timedelta(days=32)).replace(day=1)
    nm_end = (nm + dt.timedelta(days=32)).replace(day=1) - dt.timedelta(days=1)
    price_rng = as_range(st.date_input("要定價的日期區間（可跨月）", value=(nm, nm_end), key="price_rng",
                                       help="先點開始日，再點結束日。"))
    step = st.selectbox("價格級距（元）", [1, 10, 50, 100, 200, 500], index=3,
                        help="建議價格會是這個數字的倍數。級距越小計算越久，但營收幾乎不會增加。")

    method = st.radio("需求預測方式", ["只用歷史模型", "貝氏更新"], index=0,
                      help="貝氏更新：把歷史模型當先驗，用「目前訂房」比歷史模型預期多或少，"
                           "調整這段期間的整體需求水準 u。")
    use_bayes = method == "貝氏更新"
    alpha_user = None
    if use_bayes:
        if not st.checkbox("先驗強度 α 用歷史資料自動估計", value=True,
                           help="α 越大，越相信歷史模型；越小，越相信目前訂房。"):
            alpha_user = st.number_input("先驗強度 α", min_value=1.0, value=100.0, step=10.0,
                                         help="大約等於「歷史模型值得多少間訂房的證據」。")

    use_target = st.checkbox("設定平均淨收益目標", value=True,
                             help="平均淨收益 = 總淨收益 ÷ 總房數（每間夜）。不勾選：只做總淨收益最大化。")
    target = buffer = window = None
    if use_target:
        target = st.number_input("平均淨收益目標（元／間夜）", min_value=0, value=5000, step=100)
        buffer = st.number_input("安全緩衝（元）", min_value=0, value=100, step=50,
                                 help="計算時用「目標 + 緩衝」，降低實際結果略低於目標的機會。")
        same = st.checkbox("平均淨收益計算區間 = 定價區間", value=True)
        if not same:
            window = as_range(st.date_input("平均淨收益計算區間", value=(nm, nm_end), key="window",
                                            help="這段期間所有晚上（含已過去的）合計的平均淨收益要達標。"))

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
    st.subheader("貝氏先驗：歷史各月的需求水準")
    st.markdown(f"歷史上每個月「實際訂房 ÷ 模型預期」偏離 1 的程度，換算成先驗強度 **α ≈ {m.alpha_hist:.0f}**"
                "（扣除隨機誤差後；α 越大代表歷史上各月越接近模型）。")
    if m.level_hist is not None:
        st.dataframe(m.level_hist, hide_index=True, use_container_width=True)
for n in m.notes:
    st.warning(n)

# ---------------- ④ 產生建議價格 ----------------
st.header("④ 產生建議價格")
if price_rng is None:
    st.info("請在左側選好定價日期區間的開始日與結束日。")
    st.stop()
if use_target and not same and window is None:
    st.info("請在左側選好平均淨收益計算區間的開始日與結束日。")
    st.stop()

rooms = data["rooms"]
costs = dict(zip(rooms["房型"], rooms["成本"]))
bounds = {}
if {"最低價", "最高價"} <= set(rooms.columns):
    bounds = {r["房型"]: (r["最低價"], r["最高價"]) for _, r in rooms.iterrows()}
desc = f"定價區間：**{price_rng[0]} ～ {price_rng[1]}**，計算日期：**{today}**，價格級距：**{step} 元**"
if use_target:
    w_txt = "與定價區間相同" if window is None else f"{window[0]} ～ {window[1]}"
    desc += (f"<br>平均淨收益目標：**{target:,} 元／間夜** + 緩衝 {buffer:,} 元 = 計算用 **{target + buffer:,} 元**，"
             f"計算區間：**{w_txt}**")
else:
    desc += "<br>不設平均淨收益目標：只做總淨收益最大化"
st.markdown(desc, unsafe_allow_html=True)

if st.button("產生建議價格", type="primary"):
    try:
        level, level_info = 1.0, None
        if use_bayes:
            upd_dates = set(pc._date_range(*price_rng))
            if use_target and window is not None:
                upd_dates |= set(pc._date_range(*window))
            alpha = alpha_user if alpha_user is not None else m.alpha_hist
            level_info = pc.update_level(m, data["holidays"], data["otb"], today, sorted(upd_dates), alpha)
            if level_info is None:
                st.warning("「目前訂房」裡沒有這段期間已經開始銷售的資料，無法做貝氏更新，改用歷史模型。")
            else:
                level = level_info["u"]
        with st.spinner("計算中…"):
            ss.result = pc.price_range(m, data["holidays"], price_rng[0], price_rng[1], today, data["otb"],
                                       step, costs, bounds, target=target if use_target else None,
                                       buffer=buffer or 0, window=window,
                                       level=level, level_info=level_info)
    except pc.InputError as e:
        st.error(str(e))
        ss.pop("result", None)

if "result" in ss:
    res = ss.result
    s = res["summary"]
    fmt = lambda v: "—" if v is None else f"{v:,.0f}"
    k1, k2 = st.columns(2)
    k1.metric("定價區間：預期總淨收益", f"{fmt(s['定價區間：預期總淨收益'])} 元",
              (f"{s['定價區間：預期總淨收益'] - s['不設限時：定價區間預期總淨收益']:+,.0f} 元（相對不設限）"
               if res["use_target"] else f"營收 {fmt(s['定價區間：預期總營收'])} 元"), delta_color="off")
    k2.metric("定價區間：預期平均淨收益", f"{fmt(s['定價區間：預期平均淨收益（每間夜）'])} 元",
              f"平均成交價 {fmt(s['定價區間：預期平均成交價'])} 元，已成交 {s['定價區間：已成交房數']:,} 間",
              delta_color="off")
    if res["use_target"]:
        k3, k4 = st.columns(2)
        k3.metric("計算區間：預期平均淨收益", f"{fmt(s['計算區間：預期平均淨收益'])} 元",
                  f"不設限時 {fmt(s['不設限時：計算區間預期平均淨收益'])} 元", delta_color="off")
        th = s["θ（每間房的虛擬成本）"]
        k4.metric("θ（虛擬成本）", f"{fmt(th)} 元", "目標吃緊" if th > 0 else "目標沒有發揮作用",
                  delta_color="off")
    if "需求水準 u（後驗平均）" in s:
        u = s["需求水準 u（後驗平均）"]
        st.info(f"**貝氏更新：需求水準 u = {u:.2f}**（80% 區間 {s['u 的 80% 區間']}）。"
                f"目前訂房實際已售 {s['目前訂房：實際已售']:,.0f} 間，歷史模型預期 {s['目前訂房：歷史模型預期']:,.0f} 間；"
                f"先驗強度 α = {s['先驗強度 α']:.0f}，資料的權重 {s['資料的權重']}。"
                f"未來需求的預測已乘上 {u:.2f}（{'比' if u >= 1 else '比'}歷史模型{'高' if u >= 1 else '低'} {abs(u - 1):.0%}）。",
                icon="📈")
    for w in res["warnings"]:
        st.warning(w)

    st.subheader("建議價格（元）")
    st.dataframe(res["price_table"], hide_index=True, use_container_width=True, height=600)
    with st.expander("預測明細"):
        st.dataframe(res["detail"], hide_index=True, use_container_width=True)

    st.download_button("下載結果 Excel", pc.to_excel(res, m),
                       file_name=f"建議價格_{res['start']}_{res['end']}_計算日{res['today']}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       type="primary")
    st.caption("建議價格是「從今天到下次重算前」要掛的價格。建議每週（或更頻繁）用最新的訂房資料重新上傳計算。")
