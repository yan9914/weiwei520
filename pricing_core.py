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
# 3. 建議價格（日期區間可跨月；平均價限制可有可無，計算區間可自訂）
#    Lagrange：Σ p·S + μ·[Σ_{計算區間} (p-T)·S + B]
#    → 計算區間內的晚上：最大化 (p-θ)·S(p)，θ = T·μ/(1+μ)（所有晚上共用一個 θ）
#      計算區間外的晚上：最大化 p·S(p)（等於 θ = 0）
#      調整 θ 讓 Σ_{計算區間} (p-T)·S剩餘 + B ≥ 0，B = Σ_{計算區間已成交} (成交價 - T)
# ---------------------------------------------------------------
MAX_DAYS = 366
MAX_CELLS = 10_000_000          # 日期數 × 所有房型價格網格點數的上限（避免雲端記憶體不足）


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


def _date_range(a, b):
    return [a + k * ONE for k in range((b - a).days + 1)]


def price_range(model, holidays, start, end, today, otb, step, bounds=None,
                target=None, buffer=0, window=None):
    """
    start, end : 定價日期區間（含頭尾）
    target     : 平均成交價目標（元）；None = 不設限，只做營收最大化
    window     : 平均價計算區間 (起, 迄)；None = 與定價區間相同
    回傳 dict：price_table, detail, summary, warnings。金額單位：元。
    """
    if end < start:
        raise InputError("定價區間的結束日早於開始日。")
    K, names, cap = len(model.names), model.names, model.cap
    use_target = target is not None
    P_dates = _date_range(start, end)
    W_dates = _date_range(*window) if (use_target and window) else (P_dates if use_target else [])
    if use_target and window and window[1] < window[0]:
        raise InputError("平均價計算區間的結束日早於開始日。")
    dates = sorted(set(P_dates) | set(W_dates))
    if (dates[-1] - dates[0]).days + 1 > MAX_DAYS:
        raise InputError(f"定價區間與平均價計算區間合計跨越超過 {MAX_DAYS} 天，請縮短區間。")
    D = len(dates)
    in_P = np.isin(dates, P_dates)
    in_W = np.isin(dates, W_dates)
    cal = Calendar(holidays)
    months, ntypes = cal.features(dates)
    z = model.ctrl(months, ntypes)
    warnings = []

    # 目前訂房
    sold, rev = np.zeros((D, K)), np.zeros((D, K))
    if otb is not None:
        o = otb[otb["入住日"].isin(dates)]
        di = {d: k for k, d in enumerate(dates)}
        ri = {n: i for i, n in enumerate(names)}
        for _, row in o.iterrows():
            sold[di[row["入住日"]], ri[row["房型"]]] = row["已售房數"]
            rev[di[row["入住日"]], ri[row["房型"]]] = row["已售營收"] / 1000
        if len(o) == 0:
            warnings.append("「目前訂房」裡沒有這些日期的資料，視為還沒有任何訂房。")
    remaining = np.maximum(cap[None] - sold, 0)
    L = np.array([(d - today).days for d in dates])
    past = L < 0
    open_ = (~past)[:, None] & (remaining > 0)
    if (L > LMAX).any():
        warnings.append(f"有 {int((L > LMAX).sum())} 個入住日距今超過 {LMAX} 天，視為整個預訂期都還沒開始。")
    if use_target and (in_W & past).any() and otb is None:
        warnings.append("平均價計算區間包含已經過去的日期，但沒有「目前訂房」資料，這些日期會被當成沒有成交。")
    w_only = in_W & ~in_P & ~past
    if w_only.any():
        warnings.append(f"平均價計算區間有 {int(w_only.sum())} 個未來日期不在定價區間內。"
                        "為了估計整段區間的平均價，這些日期也一起規劃了價格（表中標示為「只計入平均價」）。")

    # 價格網格（每個房型分開，元 → 千元）
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
        grids.append((g if len(g) else np.array([lo])) / 1000)
    cells = D * sum(len(g) for g in grids)
    if cells > MAX_CELLS:
        raise InputError(f"計算量太大（{cells:,} 個價格組合）。請加大價格級距或縮短日期區間。")

    # 每個房型：每晚 × 每個候選價格 的預期剩餘成交數
    CUM = np.cumsum(np.exp(XH @ model.H.T).T, axis=1)                 # K × 91
    Lc = np.clip(L, 0, LMAX)
    S_list = []
    for i in range(K):
        pc = grids[i] - model.p_bar[i]
        level = np.exp(model.A[i] + model.B1[i] * pc + model.B2[i] * pc**2)
        lam = level[None] * (np.exp(z) * CUM[i, Lc])[:, None]
        Si = expected_min(lam, remaining[:, i][:, None], cap[i])
        Si[~open_[:, i]] = 0.0
        S_list.append(Si)

    T = (target + buffer) / 1000 if use_target else 0.0
    B = (rev[in_W].sum() - T * sold[in_W].sum()) if use_target else 0.0
    rows_ = np.arange(D)

    def choose(th):
        th_d = np.where(in_W, th, 0.0)[:, None]
        p, s = np.zeros((D, K)), np.zeros((D, K))
        for i in range(K):
            j = ((grids[i][None] - th_d) * S_list[i]).argmax(1)
            p[:, i], s[:, i] = grids[i][j], S_list[i][rows_, j]
        return p, s

    def gap(th):
        p, s = choose(th)
        return ((p - T) * s)[in_W].sum() + B

    feasible, theta = True, 0.0
    if use_target and gap(0.0) < 0:
        if gap(T) < 0:
            theta, feasible = T, False
            warnings.append("在目前的價格範圍內，預期平均價追不到目標。建議價格已經盡量拉高平均價；"
                            "可以考慮放寬價格上限，或重新檢視目標。")
        else:
            lo_, hi_ = 0.0, T
            for _ in range(30):
                mid = (lo_ + hi_) / 2
                lo_, hi_ = (mid, hi_) if gap(mid) < 0 else (lo_, mid)
            theta = hi_
    P, Sx = choose(theta)
    P0, S0 = choose(0.0)

    # 碰到價格邊界的提醒（只看定價區間）
    multi = np.array([len(g) > 1 for g in grids])[None]
    show = open_ & in_P[:, None] & multi
    for mask, word in [(show & np.isclose(P, np.array([g[-1] for g in grids])[None]), "上限"),
                       (show & np.isclose(P, np.array([g[0] for g in grids])[None]), "下限")]:
        if mask.any():
            ex = [f"{dates[d].strftime('%m/%d')} {names[i]}" for d, i in zip(*np.nonzero(mask))]
            tail = "代表模型認為還可以再漲，但已超出可信範圍，" if word == "上限" else ""
            warnings.append(f"{len(ex)} 格的建議價碰到價格{word}（例如 {'、'.join(ex[:3])}），{tail}請人工確認。")

    # 建議價格表
    usage = []
    for d in range(D):
        if in_P[d] and in_W[d]:
            usage.append("定價 + 計入平均價")
        elif in_P[d]:
            usage.append("定價")
        else:
            usage.append("只計入平均價")
    table = pd.DataFrame({"入住日": [d.isoformat() for d in dates],
                          "星期": [WEEKDAYS[d.weekday()] for d in dates],
                          "夜晚類型": [NIGHT_TYPES[t] for t in ntypes]})
    if use_target:
        table["用途"] = usage
    for i, n in enumerate(names):
        table[n] = ["已過" if past[d] else "售完" if remaining[d, i] == 0 else int(round(P[d, i] * 1000))
                    for d in range(D)]

    detail = pd.DataFrame([{"入住日": dates[d].isoformat(), "房型": n, "用途": usage[d] if use_target else "定價",
                            "距入住天數": int(L[d]), "已售房數": int(sold[d, i]), "剩餘房數": int(remaining[d, i]),
                            "建議價": int(round(P[d, i] * 1000)) if open_[d, i] else None,
                            "預期再售出": round(float(Sx[d, i]), 2),
                            "預期新增營收": int(round(P[d, i] * Sx[d, i] * 1000)) if open_[d, i] else 0}
                           for d in range(D) for i, n in enumerate(names)])

    def totals(mask, p, s):
        rooms = sold[mask].sum() + s[mask].sum()
        money = (rev[mask].sum() + (p * s)[mask].sum()) * 1000
        return rooms, money, (money / rooms if rooms else None)

    pr, pm, pa = totals(in_P, P, Sx)
    fr, fm, fa = totals(in_P, P0, S0)
    summary = {"定價區間": f"{start} ～ {end}（{len(P_dates)} 晚）",
               "定價區間：已成交房數": int(sold[in_P].sum()),
               "定價區間：預期總房數": pr,
               "定價區間：預期總營收": pm,
               "定價區間：預期平均價": pa}
    if use_target:
        wr, wm, wa = totals(in_W, P, Sx)
        _, _, wfa = totals(in_W, P0, S0)
        ws = sold[in_W].sum()
        summary.update({
            "平均價計算區間": f"{W_dates[0]} ～ {W_dates[-1]}（{len(W_dates)} 晚）",
            "平均價目標": target,
            "計算用目標（含緩衝）": target + buffer,
            "計算區間：已成交房數": int(ws),
            "計算區間：已成交平均價": rev[in_W].sum() * 1000 / ws if ws else None,
            "目前盈餘 B": B * 1000,
            "θ（每間房的虛擬成本）": theta * 1000,
            "計算區間：預期平均價": wa,
            "不設限時：計算區間預期平均價": wfa,
            "不設限時：定價區間預期營收": fm,
            "可達成": feasible,
        })
    return dict(price_table=table, detail=detail, summary=summary, warnings=warnings,
                start=start.isoformat(), end=end.isoformat(), today=today.isoformat(), use_target=use_target)


# ---------------------------------------------------------------
# 4. 輸出 Excel（在記憶體中產生，不寫檔）
# ---------------------------------------------------------------
def to_excel(result, model):
    from openpyxl.styles import Font

    s = result["summary"]
    fmt = lambda v: "—" if v is None else (f"{v:,.0f}" if isinstance(v, (int, float, np.floating)) and not isinstance(v, bool) else v)
    summ = pd.DataFrame({"項目": ["計算日期", "是否設定平均價目標"] + list(s),
                         "數值": [result["today"], "是" if result["use_target"] else "否（只做營收最大化）"]
                         + [("是" if v else "否") if isinstance(v, bool) else fmt(v) for v in s.values()]})
    warn = pd.DataFrame({"提醒": result["warnings"] + model.notes or ["（無）"]})
    price, night, month = model_tables(model)

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        result["price_table"].to_excel(xw, sheet_name="建議價格", index=False)
        result["detail"].to_excel(xw, sheet_name="預測明細", index=False)
        summ.to_excel(xw, sheet_name="摘要", index=False)
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
