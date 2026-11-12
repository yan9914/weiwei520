# -*- coding: utf-8 -*-
"""
定價工具核心（不讀寫任何檔案，所有資料只在記憶體中處理）
  1. read_workbook：讀取並檢查上傳的 Excel
  2. fit_model    ：設限卜瓦松迴歸（依價格時段）估計需求
  3. price_month  ：整月平均價限制下的滾動最佳化，產生建議價格
  4. to_excel     ：把結果轉成 Excel（bytes，給使用者下載）
價格在內部以「千元」計算，輸入輸出以「元」呈現。
"""
import datetime as dt
import io
from dataclasses import dataclass, field

import numpy as np
import pandas as pd
from scipy import optimize, stats
from scipy.special import gammaln

LMAX = 90
LEADS = np.arange(LMAX + 1)
_XG = np.log1p(LEADS)
X_MEAN, X_STD = _XG.mean(), _XG.std()
NIGHT_TYPES = ["一般平日", "一般週末", "平日變假日", "假日週末"]
WEEKDAYS = "一二三四五六日"
ONE = dt.timedelta(1)


class InputError(Exception):
    """上傳資料有問題時拋出，訊息會直接顯示給使用者。"""


def h_basis(L):
    x = (np.log1p(np.minimum(L, LMAX)) - X_MEAN) / X_STD
    return np.column_stack([x, x**2, x**3])


XH = h_basis(LEADS)                                    # 91 × 3


# ---------------------------------------------------------------
# 行事曆：夜晚類型 = 週末(週五/六晚) × 假日夜(隔天放假，且該段連假含國定假日)
# ---------------------------------------------------------------
class Calendar:
    def __init__(self, holidays):
        self.holidays = set(holidays)
        self._cache = {}

    def is_off(self, d):
        return d.weekday() >= 5 or d in self.holidays

    def night_type(self, d):
        if d in self._cache:
            return self._cache[d]
        weekend = d.weekday() in (4, 5)
        nxt = d + ONE
        holiday = False
        if self.is_off(nxt):
            s, e = nxt, nxt
            while self.is_off(s - ONE):
                s -= ONE
            while self.is_off(e + ONE):
                e += ONE
            holiday = any(s + k * ONE in self.holidays for k in range((e - s).days + 1))
        self._cache[d] = int(weekend) + 2 * int(holiday)
        return self._cache[d]

    def features(self, dates):
        return np.array([d.month for d in dates]), np.array([self.night_type(d) for d in dates])


# ---------------------------------------------------------------
# 1. 讀取與檢查上傳的 Excel
# ---------------------------------------------------------------
SHEETS = {
    "房型": ["房型", "房量"],
    "歷史時段": ["房型", "入住日", "牌價", "開始提前天數", "結束提前天數", "訂房數", "時段開始剩餘房數"],
    "放假日": ["日期"],
    "目前訂房": ["入住日", "房型", "已售房數", "已售營收"],
}


def _bad_rows(mask, sheet, msg, errors):
    idx = np.flatnonzero(np.asarray(mask))
    if len(idx):
        rows = "、".join(str(i + 2) for i in idx[:5]) + (" 等" if len(idx) > 5 else "")
        errors.append(f"「{sheet}」第 {rows} 列：{msg}（共 {len(idx)} 列）")


def _to_date(s):
    return pd.to_datetime(s, errors="coerce").dt.date


def read_workbook(file):
    """回傳 (data, warnings)。有錯誤時拋出 InputError。"""
    try:
        xls = pd.ExcelFile(file)
    except Exception as e:
        raise InputError(f"無法讀取 Excel 檔：{e}")
    errors, warnings = [], []
    dfs = {}
    for sheet, cols in SHEETS.items():
        if sheet not in xls.sheet_names:
            if sheet == "目前訂房":
                warnings.append("沒有「目前訂房」工作表：產生價格時會當作目標月份還沒有任何訂房。")
                continue
            if sheet == "放假日":
                warnings.append("沒有「放假日」工作表：所有國定假日都不會被辨識，夜晚類型只會分平日、週末。")
                continue
            errors.append(f"缺少工作表「{sheet}」")
            continue
        df = pd.read_excel(xls, sheet_name=sheet)
        df.columns = [str(c).strip() for c in df.columns]
        miss = [c for c in cols if c not in df.columns]
        if miss:
            errors.append(f"「{sheet}」缺少欄位：{'、'.join(miss)}")
            continue
        dfs[sheet] = df.dropna(how="all")
    if errors:
        raise InputError("\n".join(errors))

    # 房型
    rooms = dfs["房型"].copy()
    rooms["房型"] = rooms["房型"].astype(str).str.strip()
    rooms["房量"] = pd.to_numeric(rooms["房量"], errors="coerce")
    _bad_rows(rooms["房量"].isna() | (rooms["房量"] < 1) | (rooms["房量"] % 1 != 0), "房型", "房量必須是正整數", errors)
    if rooms["房型"].duplicated().any():
        errors.append("「房型」工作表的房型名稱重複")
    for c in ["最低價", "最高價"]:
        if c in rooms.columns:
            rooms[c] = pd.to_numeric(rooms[c], errors="coerce")
    names = list(rooms["房型"])

    # 歷史時段
    h = dfs["歷史時段"].copy()
    h["房型"] = h["房型"].astype(str).str.strip()
    h["入住日"] = _to_date(h["入住日"])
    num = ["牌價", "開始提前天數", "結束提前天數", "訂房數", "時段開始剩餘房數"]
    for c in num:
        h[c] = pd.to_numeric(h[c], errors="coerce")
    _bad_rows(~h["房型"].isin(names), "歷史時段", "房型不在「房型」工作表中", errors)
    _bad_rows(h["入住日"].isna(), "歷史時段", "入住日不是有效日期", errors)
    _bad_rows(h[num].isna().any(axis=1), "歷史時段", "數值欄位空白或不是數字", errors)
    _bad_rows(h["牌價"] <= 0, "歷史時段", "牌價必須大於 0", errors)
    _bad_rows((h["結束提前天數"] < 0) | (h["開始提前天數"] < h["結束提前天數"]), "歷史時段",
              "提前天數必須 ≥ 0，且開始 ≥ 結束", errors)
    _bad_rows((h["訂房數"] < 0) | (h["訂房數"] > h["時段開始剩餘房數"]), "歷史時段",
              "訂房數必須介於 0 和剩餘房數之間", errors)
    capmap = dict(zip(rooms["房型"], rooms["房量"]))
    _bad_rows((h["時段開始剩餘房數"] < 1) | (h["時段開始剩餘房數"] > h["房型"].map(capmap)), "歷史時段",
              "剩餘房數必須介於 1 和房量之間", errors)
    if (h["開始提前天數"] > LMAX).any():
        warnings.append(f"有時段的開始提前天數超過 {LMAX} 天，超過的部分會視為第 {LMAX} 天。")

    # 放假日
    holidays = []
    if "放假日" in dfs:
        hd = _to_date(dfs["放假日"]["日期"])
        _bad_rows(hd.isna(), "放假日", "不是有效日期", errors)
        holidays = [d for d in hd if pd.notna(d)]

    # 目前訂房
    otb = None
    if "目前訂房" in dfs:
        otb = dfs["目前訂房"].copy()
        otb["房型"] = otb["房型"].astype(str).str.strip()
        otb["入住日"] = _to_date(otb["入住日"])
        for c in ["已售房數", "已售營收"]:
            otb[c] = pd.to_numeric(otb[c], errors="coerce")
        _bad_rows(~otb["房型"].isin(names), "目前訂房", "房型不在「房型」工作表中", errors)
        _bad_rows(otb["入住日"].isna(), "目前訂房", "入住日不是有效日期", errors)
        _bad_rows(otb[["已售房數", "已售營收"]].isna().any(axis=1) | (otb["已售房數"] < 0), "目前訂房",
                  "已售房數、已售營收必須是 ≥ 0 的數字", errors)
        _bad_rows(otb["已售房數"] > otb["房型"].map(capmap), "目前訂房", "已售房數超過房量", errors)
        if otb.duplicated(["入住日", "房型"]).any():
            errors.append("「目前訂房」同一個入住日、房型出現不只一列")

    if errors:
        raise InputError("\n".join(errors))
    return dict(rooms=rooms, hist=h, holidays=holidays, otb=otb), warnings


# ---------------------------------------------------------------
# 2. 估計需求模型
#    log(每日訂房速度) = a_i + b1_i·(p-p̄_i) + b2_i·(p-p̄_i)² + h_i(L) + 月份 + 夜晚類型
#    μ_j = exp(a + b1·pc + b2·pc² + 控制) × Σ_{L∈時段 j} exp(h(L))
#    賣光的時段用 P(D ≥ R)（設限）
# ---------------------------------------------------------------
@dataclass
class Model:
    names: list
    cap: np.ndarray
    p_bar: np.ndarray
    p_lo: np.ndarray
    p_hi: np.ndarray
    A: np.ndarray
    B1: np.ndarray
    B2: np.ndarray
    H: np.ndarray
    month_eff: dict
    night_eff: np.ndarray
    interact: bool
    p_int: float
    quad: np.ndarray
    pvals: list
    checks: list
    n_spells: int
    cens_rate: float
    hist_range: tuple
    notes: list = field(default_factory=list)

    def ctrl(self, months, ntypes):
        miss = sorted({int(m) for m in months} - set(self.month_eff))
        if miss:
            raise InputError(f"歷史資料裡沒有 {'、'.join(map(str, miss))} 月，無法預測這個月份的需求。")
        return np.array([self.month_eff[int(m)] for m in months]) + self.night_eff[np.asarray(ntypes)]


def _design(months, ntypes, month_levels, night_cols):
    cols = [(months == m).astype(float) for m in month_levels[1:]]
    for c in night_cols:
        if c == "週末":
            cols.append(((ntypes == 1) | (ntypes == 3)).astype(float))
        elif c == "假日夜":
            cols.append((ntypes >= 2).astype(float))
        else:
            cols.append((ntypes == c).astype(float))
    return np.column_stack(cols) if cols else np.zeros((len(months), 0))


def _negll_grad(x, P, free, base):
    K = P["K"]
    full = base.copy()
    full[free] = x
    a, b1, b2 = full[0:K], full[K:2 * K], full[2 * K:3 * K]
    h, gam = full[3 * K:6 * K].reshape(K, 3), full[6 * K:]
    r, pc, Ls, Le, y, R, cens = P["r"], P["pc"], P["Ls"], P["Le"], P["y"], P["R"], P["cens"]

    E = np.exp(np.clip(XH @ h.T, -30, 30)).T
    CUM0 = np.concatenate([np.zeros((K, 1)), np.cumsum(E, 1)], 1)
    CUMX0 = np.concatenate([np.zeros((K, 1, 3)), np.cumsum(E[:, :, None] * XH[None], 1)], 1)
    HS = CUM0[r, Ls + 1] - CUM0[r, Le]

    eta = np.clip(a[r] + b1[r] * pc + b2[r] * pc**2 + P["Z"] @ gam, -30, 30)
    ee = np.exp(eta)
    mu = np.maximum(ee * HS, 1e-300)
    lsf = stats.poisson.logsf(R - 1, mu)
    ll = np.where(cens, lsf, y * np.log(mu) - mu - gammaln(y + 1))
    s = np.where(cens, np.exp(stats.poisson.logpmf(R - 1, mu) - lsf), y / mu - 1)
    sm = s * mu

    g = np.zeros_like(full)
    g[0:K] = np.bincount(r, sm, K)
    g[K:2 * K] = np.bincount(r, sm * pc, K)
    g[2 * K:3 * K] = np.bincount(r, sm * pc**2, K)
    HX = (CUMX0[r, Ls + 1] - CUMX0[r, Le]) * (s * ee)[:, None]
    g[3 * K:6 * K] = np.column_stack([np.bincount(r, HX[:, k], K) for k in range(3)]).ravel()
    g[6 * K:] = P["Z"].T @ sm
    f = -ll.sum()
    if not (np.isfinite(f) and np.all(np.isfinite(g))):
        return 1e15, np.zeros(free.sum())
    return f, -g[free]


def _fit(P, quad, cap, start=None):
    K = P["K"]
    if start is None or len(start) != 6 * K + P["Z"].shape[1]:
        start = np.zeros(6 * K + P["Z"].shape[1])
        start[0:K] = np.log(cap) - 3
        start[K:2 * K] = -0.1
    base = start.copy()
    base[2 * K:3 * K] = np.where(quad, base[2 * K:3 * K], 0.0)
    free = np.ones(len(base), bool)
    free[2 * K:3 * K] = quad
    r = optimize.minimize(_negll_grad, base[free], args=(P, free, base), jac=True,
                          method="BFGS", options={"gtol": 1e-6, "maxiter": 5000})
    full = base.copy()
    full[free] = r.x
    return full, r.fun


def fit_model(data, progress=None):
    step = progress or (lambda frac, msg: None)
    rooms, h = data["rooms"], data["hist"]
    cal = Calendar(data["holidays"])
    names = list(rooms["房型"])
    K = len(names)
    cap = rooms["房量"].to_numpy(float)
    idx = {n: i for i, n in enumerate(names)}

    r = h["房型"].map(idx).to_numpy(int)
    missing = [n for i, n in enumerate(names) if not np.any(r == i)]
    if missing:
        raise InputError(f"這些房型沒有歷史資料：{'、'.join(missing)}")
    price = h["牌價"].to_numpy(float) / 1000
    Ls = np.minimum(h["開始提前天數"].to_numpy(int), LMAX)
    Le = np.minimum(h["結束提前天數"].to_numpy(int), Ls)
    y, R = h["訂房數"].to_numpy(float), h["時段開始剩餘房數"].to_numpy(float)
    udates = sorted(set(h["入住日"]))
    m_u, n_u = cal.features(udates)
    dmap = {d: k for k, d in enumerate(udates)}
    di = h["入住日"].map(dmap).to_numpy(int)
    months, ntypes = m_u[di], n_u[di]

    p_bar = np.array([np.round(price[r == i].mean(), 1) for i in range(K)])
    p_lo = np.array([price[r == i].min() for i in range(K)])
    p_hi = np.array([price[r == i].max() for i in range(K)])
    notes = []
    for i in range(K):
        if price[r == i].std() < 0.05:
            notes.append(f"{names[i]} 的歷史價格幾乎沒有變動，價格效果會估不準。")

    month_levels = sorted(set(months))
    if 1 in month_levels:
        month_levels.remove(1)
        month_levels = [1] + month_levels
    present = set(ntypes)

    base_P = dict(K=K, r=r, pc=price - p_bar[r], Ls=Ls, Le=Le, y=y, R=R, cens=y >= R)
    allq = np.ones(K, bool)

    # (1) 夜晚類型：交互作用檢定（需要四種夜晚都出現過）
    add_cols = [c for c, ok in [("週末", present & {1, 3}), ("假日夜", present & {2, 3})] if ok]
    P_add = dict(base_P, Z=_design(months, ntypes, month_levels, add_cols))
    step(0.1, "估計模型（沒有交互作用）…")
    full_add, nll_add = _fit(P_add, allq, cap)
    if {1, 2, 3} <= present:
        P_int = dict(base_P, Z=_design(months, ntypes, month_levels, [1, 2, 3]))
        step(0.3, "估計模型（有交互作用）…")
        full_int, nll_int = _fit(P_int, allq, cap)
        p_int = float(stats.chi2.sf(max(2 * (nll_add - nll_int), 0), 1))
        interact = p_int < 0.05
    else:
        p_int, interact = float("nan"), False
        notes.append("歷史資料沒有涵蓋全部四種夜晚類型，無法檢定交互作用，改用「週末 + 假日夜」。")
    P_use, full0, nll0 = (P_int, full_int, nll_int) if interact else (P_add, full_add, nll_add)
    night_cols = [1, 2, 3] if interact else add_cols

    # (2) 每個房型：一次式 vs 二次式
    quad = np.zeros(K, bool)
    pvals, checks = [], []
    for i in range(K):
        step(0.4 + 0.5 * i / K, f"檢定 {names[i]} 的價格曲線…")
        q = allq.copy()
        q[i] = False
        _, nll_i = _fit(P_use, q, cap, full0)
        pv = float(stats.chi2.sf(max(2 * (nll_i - nll0), 0), 1))
        b1, b2 = full0[K + i], full0[2 * K + i]
        if b2 <= 0:
            check = "OK（b2 ≤ 0）"
        else:
            vertex = p_bar[i] - b1 / (2 * b2)
            check = (f"OK（頂點 {vertex*1000:,.0f} 在範圍外）" if not (p_lo[i] <= vertex <= p_hi[i])
                     else f"頂點 {vertex*1000:,.0f} 在範圍內，改用一次式")
        quad[i] = pv < 0.05 and check.startswith("OK")
        pvals.append(pv)
        checks.append(check)

    step(0.95, "估計最終模型…")
    params, _ = _fit(P_use, quad, cap, full0)
    gam = params[6 * K:]
    nm = len(month_levels) - 1
    month_eff = {month_levels[0]: 0.0}
    month_eff.update({m: gam[k] for k, m in enumerate(month_levels[1:])})
    ng = dict(zip(night_cols, gam[nm:]))
    if interact:
        night_eff = np.array([0.0, ng[1], ng[2], ng[3]])
    else:
        w, hol = ng.get("週末", 0.0), ng.get("假日夜", 0.0)
        night_eff = np.array([0.0, w, hol, w + hol])
    for k in (1, 2, 3):
        if k not in present:
            notes.append(f"歷史資料沒有「{NIGHT_TYPES[k]}」，這類晚上的效果是推估的。")
    B1 = params[K:2 * K]
    for i in range(K):
        if B1[i] >= 0:
            notes.append(f"{names[i]} 估出來的價格係數不是負的（漲價需求反而增加），"
                         "通常代表旺季、假日等因素沒有控制好，這個房型的建議價格不可信。")

    step(1.0, "完成")
    return Model(names=names, cap=cap, p_bar=p_bar, p_lo=p_lo, p_hi=p_hi,
                 A=params[0:K], B1=B1, B2=params[2 * K:3 * K], H=params[3 * K:6 * K].reshape(K, 3),
                 month_eff=month_eff, night_eff=night_eff, interact=interact, p_int=p_int,
                 quad=quad, pvals=pvals, checks=checks, n_spells=len(h),
                 cens_rate=float(np.mean(y >= R)), hist_range=(min(udates), max(udates)), notes=notes)


def model_tables(m):
    """把模型估計結果整理成給人看的表格。"""
    price = pd.DataFrame({
        "房型": m.names,
        "房量": m.cap.astype(int),
        "歷史價格範圍": [f"{lo*1000:,.0f}–{hi*1000:,.0f}" for lo, hi in zip(m.p_lo, m.p_hi)],
        "置中價格 p̄": [f"{v*1000:,.0f}" for v in m.p_bar],
        "b1（每千元）": np.round(m.B1, 3),
        "b2": np.round(m.B2, 4),
        "二次項檢定 p 值": np.round(m.pvals, 3),
        "採用": ["二次" if q else "一次" for q in m.quad],
        "b2 檢查": m.checks,
    })
    night = pd.DataFrame({"夜晚類型": NIGHT_TYPES,
                          "需求是一般平日的幾倍": np.round(np.exp(m.night_eff), 2)})
    base = min(m.month_eff, key=lambda k: abs(m.month_eff[k]))
    month = pd.DataFrame({"月份": [f"{k} 月" for k in sorted(m.month_eff)],
                          f"需求是 {base} 月的幾倍": [round(float(np.exp(m.month_eff[k])), 2)
                                                for k in sorted(m.month_eff)]})
    return price, night, month


# ---------------------------------------------------------------
# 3. 整月建議價格
#    Lagrange：Σ p·S + μ·[Σ (p-T)·S + B] = (1+μ)·Σ (p-θ)·S + 常數，θ = T·μ/(1+μ)
#    → 整個月只有一個 θ；每個 (入住日, 房型) 各自最大化 (p-θ)·S(p)
#      再調整 θ 讓 Σ (p-T)·S剩餘 + B ≥ 0，B = Σ已成交 (成交價 - T)
# ---------------------------------------------------------------
def expected_min(lam, R, rmax):
    """E[min(D, R)] = Σ_{k=0}^{R-1} P(D > k)。"""
    pmf = np.exp(-lam)
    cdf = pmf.copy()
    out = np.zeros_like(lam)
    for k in range(int(rmax)):
        out += np.where(k < R, 1.0 - cdf, 0.0)
        pmf = pmf * lam / (k + 1)
        cdf = cdf + pmf
    return out


def price_month(model, holidays, year, month, today, otb, target, buffer, step, bounds=None):
    """回傳 dict：price_table, detail, summary, warnings。金額單位：元。"""
    K, names, cap = len(model.names), model.names, model.cap
    cal = Calendar(holidays)
    first = dt.date(year, month, 1)
    nxt = dt.date(year + (month == 12), month % 12 + 1, 1)
    dates = [first + k * ONE for k in range((nxt - first).days)]
    D = len(dates)
    months, ntypes = cal.features(dates)
    z = model.ctrl(months, ntypes)
    warnings = []

    sold, rev = np.zeros((D, K)), np.zeros((D, K))
    if otb is not None:
        o = otb[otb["入住日"].isin(dates)]
        di = {d: k for k, d in enumerate(dates)}
        ri = {n: i for i, n in enumerate(names)}
        for _, row in o.iterrows():
            sold[di[row["入住日"]], ri[row["房型"]]] = row["已售房數"]
            rev[di[row["入住日"]], ri[row["房型"]]] = row["已售營收"] / 1000
        if len(o) == 0:
            warnings.append("「目前訂房」裡沒有這個月份的資料，視為還沒有任何訂房。")
    remaining = np.maximum(cap[None] - sold, 0)
    L = np.array([(d - today).days for d in dates])
    past = L < 0
    open_ = (~past)[:, None] & (remaining > 0)
    if (L > LMAX).any():
        warnings.append(f"有 {int((L > LMAX).sum())} 個入住日距今超過 {LMAX} 天，視為整個預訂期都還沒開始。")

    # 價格網格
    bounds = bounds or {}
    grids = []
    for i, n in enumerate(names):
        lo, hi = bounds.get(n, (None, None))
        lo = model.p_lo[i] * 1000 if lo is None or np.isnan(lo) else lo
        hi = model.p_hi[i] * 1000 if hi is None or np.isnan(hi) else hi
        if hi < lo:
            raise InputError(f"{n} 的最高價低於最低價。")
        if lo < model.p_lo[i] * 1000 - 1 or hi > model.p_hi[i] * 1000 + 1:
            warnings.append(f"{n} 的價格範圍超出歷史價格（{model.p_lo[i]*1000:,.0f}–{model.p_hi[i]*1000:,.0f}），"
                            "超出的部分是模型外推，可信度較低。")
        g = np.arange(np.ceil(lo / step) * step, hi + 1e-6, step)
        if len(g) == 0:
            g = np.array([lo])
        grids.append(g / 1000)
    NG = max(len(g) for g in grids)
    GRID = np.full((K, NG), np.nan)
    for i, g in enumerate(grids):
        GRID[i, :len(g)] = g
    PC = GRID - model.p_bar[:, None]
    LEVEL = np.exp(model.A[:, None] + model.B1[:, None] * PC + model.B2[:, None] * PC**2)
    CUM = np.cumsum(np.exp(XH @ model.H.T).T, axis=1)                 # K × 91

    Lc = np.clip(L, 0, LMAX)
    lam = LEVEL[None] * np.exp(z)[:, None, None] * CUM[:, Lc].T[:, :, None]
    S = expected_min(np.nan_to_num(lam), remaining[:, :, None], cap.max())
    S[~open_] = 0.0
    S = np.where(np.isnan(GRID)[None], np.nan, S)

    T = (target + buffer) / 1000
    B = rev.sum() - T * sold.sum()

    def choose(th):
        obj = np.where(np.isnan(S), -np.inf, (GRID[None] - th) * np.nan_to_num(S))
        j = obj.argmax(-1)
        p = np.take_along_axis(np.broadcast_to(GRID, S.shape), j[..., None], -1)[..., 0]
        s = np.take_along_axis(np.nan_to_num(S), j[..., None], -1)[..., 0]
        return p, s

    def gap(th):
        p, s = choose(th)
        return ((p - T) * s).sum() + B

    feasible = True
    if gap(0.0) >= 0:
        theta = 0.0
    elif gap(T) < 0:
        theta, feasible = T, False
        warnings.append("在目前的價格範圍內，預期全月平均價追不到目標。建議價格已經盡量拉高平均價；"
                        "可以考慮放寬價格上限，或重新檢視目標。")
    else:
        lo_, hi_ = 0.0, T
        for _ in range(25):
            mid = (lo_ + hi_) / 2
            lo_, hi_ = (mid, hi_) if gap(mid) < 0 else (lo_, mid)
        theta = hi_
    P, Sx = choose(theta)
    P0, S0 = choose(0.0)

    # 輸出表格
    at_lo = open_ & np.isclose(P, GRID[:, 0][None]) & (np.array([len(g) for g in grids])[None] > 1)
    at_hi = open_ & np.isclose(P, np.array([g[-1] for g in grids])[None]) & (np.array([len(g) for g in grids])[None] > 1)
    if at_hi.any():
        cells = [f"{dates[d].strftime('%m/%d')} {names[i]}" for d, i in zip(*np.nonzero(at_hi))]
        warnings.append(f"{len(cells)} 格的建議價碰到價格上限（例如 {'、'.join(cells[:3])}），"
                        "代表模型認為還可以再漲，但已超出可信範圍，請人工確認。")
    if at_lo.any():
        cells = [f"{dates[d].strftime('%m/%d')} {names[i]}" for d, i in zip(*np.nonzero(at_lo))]
        warnings.append(f"{len(cells)} 格的建議價碰到價格下限（例如 {'、'.join(cells[:3])}），請人工確認。")

    table = pd.DataFrame({"入住日": [d.isoformat() for d in dates],
                          "星期": [WEEKDAYS[d.weekday()] for d in dates],
                          "夜晚類型": [NIGHT_TYPES[t] for t in ntypes]})
    for i, n in enumerate(names):
        col = []
        for d in range(D):
            if past[d]:
                col.append("已過")
            elif remaining[d, i] == 0:
                col.append("售完")
            else:
                col.append(int(round(P[d, i] * 1000)))
        table[n] = col

    rows = []
    for d in range(D):
        for i, n in enumerate(names):
            rows.append({"入住日": dates[d].isoformat(), "房型": n, "距入住天數": int(L[d]),
                         "已售房數": int(sold[d, i]), "剩餘房數": int(remaining[d, i]),
                         "建議價": int(round(P[d, i] * 1000)) if open_[d, i] else None,
                         "預期再售出": round(float(Sx[d, i]), 2),
                         "預期新增營收": int(round(P[d, i] * Sx[d, i] * 1000)) if open_[d, i] else 0})
    detail = pd.DataFrame(rows)

    otb_rooms, otb_rev = sold.sum(), rev.sum() * 1000
    exp_rooms = otb_rooms + Sx.sum()
    exp_rev = otb_rev + (P * Sx).sum() * 1000
    free_rooms = otb_rooms + S0.sum()
    free_rev = otb_rev + (P0 * S0).sum() * 1000
    summary = {
        "已成交房數": int(otb_rooms),
        "已成交平均價": otb_rev / otb_rooms if otb_rooms else None,
        "目前盈餘 B": B * 1000,
        "θ（每間房的虛擬成本）": theta * 1000,
        "預期全月房數": exp_rooms,
        "預期全月營收": exp_rev,
        "預期全月平均價": exp_rev / exp_rooms if exp_rooms else None,
        "不設限時預期營收": free_rev,
        "不設限時預期平均價": free_rev / free_rooms if free_rooms else None,
        "計算用的平均價目標": target + buffer,
        "可達成": feasible,
    }
    return dict(price_table=table, detail=detail, summary=summary, warnings=warnings,
                month=f"{year}-{month:02d}", today=today.isoformat())


# ---------------------------------------------------------------
# 4. 輸出 Excel（在記憶體中產生，不寫檔）
# ---------------------------------------------------------------
def to_excel(result, model):
    from openpyxl.styles import Font

    s = result["summary"]
    fmt = lambda v: "—" if v is None else (f"{v:,.0f}" if isinstance(v, (int, float, np.floating)) and not isinstance(v, bool) else v)
    summ = pd.DataFrame({"項目": ["定價月份", "計算日期"] + list(s),
                         "數值": [result["month"], result["today"]]
                         + [("是" if v else "否") if isinstance(v, bool) else fmt(v) for v in s.values()]})
    warn = pd.DataFrame({"提醒": result["warnings"] + model.notes or ["（無）"]})
    price, night, month = model_tables(model)

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        result["price_table"].to_excel(xw, sheet_name="建議價格", index=False)
        result["detail"].to_excel(xw, sheet_name="預測明細", index=False)
        summ.to_excel(xw, sheet_name="月度摘要", index=False)
        warn.to_excel(xw, sheet_name="提醒", index=False)
        price.to_excel(xw, sheet_name="模型_價格效果", index=False)
        night.to_excel(xw, sheet_name="模型_夜晚類型", index=False)
        month.to_excel(xw, sheet_name="模型_月份", index=False)
        for ws in xw.book.worksheets:
            for row in ws.iter_rows():
                for c in row:
                    c.font = Font(name="微軟正黑體", bold=c.row == 1)
                    if isinstance(c.value, (int, float)) and c.row > 1:
                        c.number_format = "#,##0.00" if isinstance(c.value, float) and abs(c.value) < 100 else "#,##0"
            for col in ws.columns:
                width = max(len(str(c.value)) if c.value is not None else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(max(10, width * 1.6), 60)
            ws.freeze_panes = "A2"
    return buf.getvalue()
