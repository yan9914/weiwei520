# -*- coding: utf-8 -*-
"""
產生範例 Excel（sample_data.xlsx）：完全是模擬資料，用來示範上傳格式與試用網站。
執行：python make_sample.py
"""
import datetime as dt

import numpy as np
import pandas as pd

from pricing_core import Calendar, LMAX

rng = np.random.default_rng(7)
names = ["標準雙人", "高級雙人", "豪華雙人", "家庭四人", "景觀套房", "總統套房"]
cap = np.array([20, 15, 12, 8, 5, 2])
p_bar = np.array([4.0, 5.0, 6.5, 7.5, 9.0, 12.0])            # 千元


def days(a, b):
    a, b = dt.date.fromisoformat(a), dt.date.fromisoformat(b)
    return [a + dt.timedelta(k) for k in range((b - a).days + 1)]


# ⚠ 模擬用的放假日，不保證與實際公告一致
HOLIDAYS = sorted(set(
    days("2024-01-01", "2024-01-01") + days("2024-02-08", "2024-02-14") + days("2024-02-28", "2024-02-28")
    + days("2024-04-04", "2024-04-07") + days("2024-06-10", "2024-06-10") + days("2024-09-17", "2024-09-17")
    + days("2024-10-10", "2024-10-10")
    + days("2025-01-01", "2025-01-01") + days("2025-01-25", "2025-02-02") + days("2025-02-28", "2025-02-28")
    + days("2025-04-03", "2025-04-06") + days("2025-05-30", "2025-05-30") + days("2025-10-06", "2025-10-06")
    + days("2025-10-10", "2025-10-10")
    + days("2026-01-01", "2026-01-01") + days("2026-02-14", "2026-02-22") + days("2026-02-27", "2026-02-28")
    + days("2026-04-03", "2026-04-06") + days("2026-06-19", "2026-06-19") + days("2026-09-25", "2026-09-25")
    + days("2026-10-09", "2026-10-10")))
cal = Calendar(HOLIDAYS)

leads = np.arange(LMAX + 1)
w = np.exp(-leads / 25) + 0.6 * np.exp(-leads / 3)
true_h = np.log(w / w.sum())
true_a = np.log(0.9 * cap)
true_b1 = -1 / p_bar
true_b2 = np.array([-0.02, -0.01, 0.0, -0.005, 0.003, 0.0])
TRUE_MONTH = np.array([0.0, 0.1, -0.5, 0.0, -0.2, -0.1, 0.3, 0.35, -0.15, 0.05, -0.4, 0.1])
TRUE_NIGHT = np.array([0.0, 0.4, 0.5, 0.55])


# 每個「年-月」的市場需求衝擊（歷史模型看不到）：歷史月份隨機 ±15% 左右；
# 2026 年 3 月設定為比平常熱 25%，用來示範貝氏更新能從目前訂房偵測到這個變化
SHOCK = {(y, m): float(np.exp(rng.normal(0, 0.15))) for y in (2024, 2025) for m in range(1, 13)}
SHOCK[(2026, 3)] = 1.25


def rate(i, p, d, L):
    pc = p - p_bar[i]
    z = TRUE_MONTH[d.month - 1] + TRUE_NIGHT[cal.night_type(d)]
    return SHOCK.get((d.year, d.month), 1.0) * np.exp(true_a[i] + true_b1[i] * pc + true_b2[i] * pc**2
                                                       + true_h[L] + z)


# 歷史時段：2024–2025，依訂房速度動態調價
spell_def = [(90, 61), (60, 31), (30, 15), (14, 8), (7, 4), (3, 0)]
rows = []
for d in days("2024-01-01", "2025-12-31"):
    for i in range(len(names)):
        p = p_bar[i] * rng.uniform(0.75, 1.25)
        booked = 0
        for s, (Ls, Le) in enumerate(spell_def):
            if s > 0:
                target = 0.85 * cap[i] * w[Ls + 1:].sum() / w.sum()
                ratio = booked / max(target, 0.5)
                p *= rng.uniform(1.05, 1.15) if ratio > 1.15 else rng.uniform(0.88, 0.97) if ratio < 0.85 \
                    else rng.uniform(0.97, 1.03)
                p = np.clip(p, 0.6 * p_bar[i], 1.8 * p_bar[i])
            p = round(p * 10) / 10
            R = cap[i] - booked
            if R == 0:
                break
            y = min(rng.poisson(rate(i, p, d, np.arange(Le, Ls + 1)).sum()), R)
            rows.append((names[i], d, int(p * 1000), Ls, Le, y, R))
            booked += y
hist = pd.DataFrame(rows, columns=["房型", "入住日", "牌價", "開始提前天數", "結束提前天數", "訂房數", "時段開始剩餘房數"])

# 目前時段：假設今天是 2026-02-15，要定 2026 年 3 月的價格。
# 從開賣到今天，同樣依訂房速度動態調價；最後一段的結束提前天數 = 今天距入住日的天數
today = dt.date(2026, 2, 15)
cur_rows = []
for d in days("2026-03-01", "2026-03-31"):
    L_now = (d - today).days
    for i in range(len(names)):
        p = p_bar[i] * rng.uniform(0.85, 1.35)
        booked = 0
        for s_, (Ls, Le) in enumerate(spell_def):
            if Ls < L_now:                                     # 這段還沒開始
                break
            if s_ > 0:
                target = 0.85 * cap[i] * w[Ls + 1:].sum() / w.sum()
                ratio = booked / max(target, 0.5)
                p *= rng.uniform(1.05, 1.15) if ratio > 1.15 else rng.uniform(0.80, 0.92) if ratio < 0.85 \
                    else rng.uniform(0.97, 1.03)
                p = np.clip(p, 0.6 * p_bar[i], 1.8 * p_bar[i])
            p = round(p * 10) / 10
            R = cap[i] - booked
            if R == 0:
                break
            Le_ = max(Le, L_now)
            y = min(rng.poisson(rate(i, p, d, np.arange(Le_, Ls + 1)).sum()), R)
            cur_rows.append((names[i], d, int(p * 1000), Ls, Le_, y, R))
            booked += y
cur = pd.DataFrame(cur_rows, columns=["房型", "入住日", "牌價", "開始提前天數", "結束提前天數", "訂房數", "時段開始剩餘房數"])

# 目前訂房：由目前時段彙總（實務上已售營收建議用訂房系統的實際金額）
cur["營收"] = cur["訂房數"] * cur["牌價"]
g = cur.groupby(["入住日", "房型"], sort=False)
otb = g.agg(已售房數=("訂房數", "sum"), 已售營收=("營收", "sum")).reset_index()
otb["目前牌價"] = g["牌價"].last().to_numpy()
cur = cur.drop(columns="營收")

rooms = pd.DataFrame({"房型": names, "房量": cap, "成本": [1200, 1400, 1700, 2000, 2300, 3500],
                      "最低價": np.nan, "最高價": np.nan})
readme = pd.DataFrame({"說明": [
    "這是模擬資料，只用來示範格式。",
    "房型：房型名稱、房量、成本（每賣出一間夜的變動成本，元，例如清潔、備品、水電）；",
    "  最低價、最高價可留空（留空 = 用歷史價格範圍）。",
    "  淨收益 = (成交價 - 成本) × 房數；工具會最大化總淨收益。",
    "歷史時段：每列是某入住日、某房型、一段固定價格的期間。牌價以元為單位。",
    "  開始提前天數 / 結束提前天數：這個價格從入住前幾天掛到入住前幾天（含），例如 30 和 15。",
    "  訂房數：這段期間收到幾間訂房；時段開始剩餘房數：這段開始時還剩幾間可賣。",
    "  賣光之後的時段不要列出來（不能當成需求是 0）。",
    "放假日：國定假日與補假日期（週六日不用列）。⚠ 範例裡的日期是模擬的，請換成官方行事曆。",
    "目前時段（建議提供）：格式和「歷史時段」完全相同，換成要定價的入住日，記錄從開賣到今天的每一段牌價與訂房數。",
    "  最後一段的結束提前天數 = 今天距入住日的天數。沒有賣出的入住日也要列出它的時段（訂房數 = 0）。",
    "  有這張表時，貝氏更新會逐段精確計算；也可以由它推算已售房數、已售營收（訂房數 × 牌價）、目前牌價。",
    "目前訂房（選填）：要定價的期間裡，每晚每個房型的已售房數、已售營收（元）、目前牌價（元）。",
    "  有提供時，已售房數與已售營收以這張表為準（例如有折扣、不同通路價、取消時，請用實際金額）。",
    "  沒有賣出的格子也要列出來（已售房數 = 0），並填目前牌價：貝氏更新需要知道這些格子掛的價格。",
    "  已售營收：實際成交金額，用來計算已成交的淨收益；沒有填時，會用「已售房數 × 目前牌價」估算。",
    "  已售營收、目前牌價至少要有一欄，建議兩欄都填。",
]})

with pd.ExcelWriter("sample_data.xlsx", engine="openpyxl") as xw:
    readme.to_excel(xw, sheet_name="說明", index=False)
    rooms.to_excel(xw, sheet_name="房型", index=False)
    hist.to_excel(xw, sheet_name="歷史時段", index=False)
    pd.DataFrame({"日期": HOLIDAYS}).to_excel(xw, sheet_name="放假日", index=False)
    cur.to_excel(xw, sheet_name="目前時段", index=False)
    otb.to_excel(xw, sheet_name="目前訂房", index=False)
print(f"sample_data.xlsx：歷史時段 {len(hist):,} 列，目前時段 {len(cur)} 列，目前訂房 {len(otb)} 列（今天 = {today}）")
