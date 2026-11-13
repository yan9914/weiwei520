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
    "房型": ["房型", "房量", "成本"],
    "歷史時段": ["房型", "入住日", "牌價", "開始提前天數", "結束提前天數", "訂房數", "時段開始剩餘房數"],
    "放假日": ["日期"],
    "目前訂房": ["入住日", "房型", "已售房數"],     # 另需「已售營收」或「目前牌價」至少一欄（建議兩欄都有）
}


def _bad_rows(mask, sheet, msg, errors):
    idx = np.flatnonzero(np.asarray(mask))
    if len(idx):
        rows = "、".join(str(i + 2) for i in idx[:5]) + (" 等" if len(idx) > 5 else "")
        errors.append(f"「{sheet}」第 {rows} 列：{msg}（共 {len(idx)} 列）")


def _to_date(s):
    return pd.to_datetime(s, errors="coerce").dt.date


SPELL_COLS = ["牌價", "開始提前天數", "結束提前天數", "訂房數", "時段開始剩餘房數"]


def _parse_spells(df, sheet, names, capmap, errors, warnings):
    """檢查「歷史時段」或「目前時段」（兩者格式相同）。"""
    h = df.copy()
    h["房型"] = h["房型"].astype(str).str.strip()
    h["入住日"] = _to_date(h["入住日"])
    for c in SPELL_COLS:
        h[c] = pd.to_numeric(h[c], errors="coerce")
    _bad_rows(~h["房型"].isin(names), sheet, "房型不在「房型」工作表中", errors)
    _bad_rows(h["入住日"].isna(), sheet, "入住日不是有效日期", errors)
    _bad_rows(h[SPELL_COLS].isna().any(axis=1), sheet, "數值欄位空白或不是數字", errors)
    _bad_rows(h["牌價"] <= 0, sheet, "牌價必須大於 0", errors)
    _bad_rows((h["結束提前天數"] < 0) | (h["開始提前天數"] < h["結束提前天數"]), sheet,
              "提前天數必須 ≥ 0，且開始 ≥ 結束", errors)
    _bad_rows((h["訂房數"] < 0) | (h["訂房數"] > h["時段開始剩餘房數"]), sheet,
              "訂房數必須介於 0 和剩餘房數之間", errors)
    _bad_rows((h["時段開始剩餘房數"] < 1) | (h["時段開始剩餘房數"] > h["房型"].map(capmap)), sheet,
              "剩餘房數必須介於 1 和房量之間", errors)
    if (h["開始提前天數"] > LMAX).any():
        warnings.append(f"「{sheet}」有時段的開始提前天數超過 {LMAX} 天，超過的部分會視為第 {LMAX} 天。")
    return h


def otb_from_spells(cur):
    """由「目前時段」推出每晚每房型的已售房數、已售營收（訂房數 × 牌價）、目前牌價（最近一段的牌價）。"""
    c = cur.copy()
    c["營收"] = c["訂房數"] * c["牌價"]
    g = c.groupby(["入住日", "房型"])
    last = c.loc[g["結束提前天數"].idxmin(), ["入住日", "房型", "牌價"]].rename(columns={"牌價": "目前牌價"})
    out = g.agg(已售房數=("訂房數", "sum"), 已售營收=("營收", "sum")).reset_index()
    return out.merge(last, on=["入住日", "房型"])


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
    rooms["成本"] = pd.to_numeric(rooms["成本"], errors="coerce")
    _bad_rows(rooms["成本"].isna() | (rooms["成本"] < 0), "房型", "成本必須是 ≥ 0 的數字（每間夜，元）", errors)
    for c in ["最低價", "最高價"]:
        if c in rooms.columns:
            rooms[c] = pd.to_numeric(rooms[c], errors="coerce")
    names = list(rooms["房型"])

    # 歷史時段
    capmap = dict(zip(rooms["房型"], rooms["房量"]))
    h = _parse_spells(dfs["歷史時段"], "歷史時段", names, capmap, errors, warnings)

    # 目前時段（選填）：要定價的入住日，從開賣到今天的每段牌價與訂房數，格式同「歷史時段」
    cur = None
    if "目前時段" in xls.sheet_names:
        cdf = pd.read_excel(xls, sheet_name="目前時段")
        cdf.columns = [str(c).strip() for c in cdf.columns]
        miss = [c for c in SHEETS["歷史時段"] if c not in cdf.columns]
        if miss:
            errors.append(f"「目前時段」缺少欄位：{'、'.join(miss)}")
        else:
            cur = _parse_spells(cdf.dropna(how="all"), "目前時段", names, capmap, errors, warnings)

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
        has_rev, has_px = "已售營收" in otb.columns, "目前牌價" in otb.columns
        if not (has_rev or has_px):
            errors.append("「目前訂房」至少要有「已售營收」或「目前牌價」其中一欄")
        for c in ["已售房數", "已售營收", "目前牌價"]:
            otb[c] = pd.to_numeric(otb[c], errors="coerce") if c in otb.columns else np.nan
        _bad_rows(~otb["房型"].isin(names), "目前訂房", "房型不在「房型」工作表中", errors)
        _bad_rows(otb["入住日"].isna(), "目前訂房", "入住日不是有效日期", errors)
        _bad_rows(otb["已售房數"].isna() | (otb["已售房數"] < 0), "目前訂房", "已售房數必須是 ≥ 0 的數字", errors)
        _bad_rows(otb["已售營收"] < 0, "目前訂房", "已售營收不能是負數", errors)
        _bad_rows(otb["目前牌價"] <= 0, "目前訂房", "目前牌價必須大於 0", errors)
        if not has_px:
            warnings.append("「目前訂房」沒有「目前牌價」欄：貝氏更新時，還沒賣出的格子只能用歷史平均價近似，"
                            "需求水準 u 可能偏低。")
        if has_rev:
            miss = otb["已售營收"].isna() & (otb["已售房數"] > 0)
        else:
            miss = otb["已售房數"] > 0
        if miss.any():
            fill = otb["已售房數"] * otb["目前牌價"]
            _bad_rows(miss & fill.isna(), "目前訂房", "有已售房數，但已售營收和目前牌價都是空白", errors)
            otb.loc[miss, "已售營收"] = fill[miss]
            warnings.append(f"「目前訂房」有 {int(miss.sum())} 列沒有已售營收，暫以「已售房數 × 目前牌價」估算。"
                            "如果這些訂房是用其他價格成交的，已成交的淨收益與平均淨收益目標的計算會有誤差。")
        otb["已售營收"] = otb["已售營收"].fillna(0.0)
        _bad_rows(otb["已售房數"] > otb["房型"].map(capmap), "目前訂房", "已售房數超過房量", errors)
        if otb.duplicated(["入住日", "房型"]).any():
            errors.append("「目前訂房」同一個入住日、房型出現不只一列")

    if errors:
        raise InputError("\n".join(errors))

    if cur is not None and len(cur):
        derived = otb_from_spells(cur)
        if otb is None:
            otb = derived
            warnings.append("沒有「目前訂房」工作表：已售房數、已售營收、目前牌價改由「目前時段」推算"
                            "（已售營收 = Σ 訂房數 × 牌價）。如果有折扣、不同通路價或取消，"
                            "建議另外提供「目前訂房」的實際已售營收。")
        else:
            chk = otb.merge(derived[["入住日", "房型", "已售房數"]], on=["入住日", "房型"], how="inner",
                            suffixes=("", "_時段"))
            diff = chk[chk["已售房數"] != chk["已售房數_時段"]]
            if len(diff):
                ex = "、".join(f"{r.入住日} {r.房型}" for r in diff.head(3).itertuples())
                warnings.append(f"「目前訂房」與「目前時段」有 {len(diff)} 格的已售房數不一致（例如 {ex}）。"
                                "已售房數與營收以「目前訂房」為準；貝氏更新使用「目前時段」。"
                                "常見原因是取消訂房，或兩張表的截止日期不同。")
    elif otb is None:
        warnings.append("沒有「目前訂房」或「目前時段」：產生價格時會當作還沒有任何訂房。")
    return dict(rooms=rooms, hist=h, holidays=holidays, otb=otb, cur=cur), warnings


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
    alpha_hist: float = 400.0                  # 從歷史資料估的貝氏先驗強度
    level_hist: pd.DataFrame = None            # 歷史各年月的「實際 ÷ 模型」比例

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


def _spell_mu(full, P):
    """每個時段的預期需求 μ_j（和 _negll_grad 同一套公式）。"""
    K = P["K"]
    a, b1, b2 = full[0:K], full[K:2 * K], full[2 * K:3 * K]
    h, gam = full[3 * K:6 * K].reshape(K, 3), full[6 * K:]
    r = P["r"]
    E = np.exp(np.clip(XH @ h.T, -30, 30)).T
    CUM0 = np.concatenate([np.zeros((K, 1)), np.cumsum(E, 1)], 1)
    HS = CUM0[r, P["Ls"] + 1] - CUM0[r, P["Le"]]
    eta = np.clip(a[r] + b1[r] * P["pc"] + b2[r] * P["pc"]**2 + P["Z"] @ gam, -30, 30)
    return np.exp(eta) * HS


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

    # 貝氏先驗強度 α：歷史上每個「年-月」的實際訂房 ÷ 模型預期，看它們離 1 多遠
    #   比例的變異數 ≈ 真實的月度需求波動 + 卜瓦松隨機誤差（≈ 1/預期數），扣掉後者，α = 1 / 前者
    mu = _spell_mu(params, P_use)
    exp_obs = expected_min(mu, R, cap.max())                  # 考慮剩餘房數上限後的預期成交數
    ym = np.array([d.year * 100 + d.month for d in udates])[di]
    keys = sorted(set(ym))
    obs = np.array([y[ym == k].sum() for k in keys])
    exp_ = np.array([exp_obs[ym == k].sum() for k in keys])
    ratios = obs / exp_
    var_true = np.var(ratios, ddof=1) - np.mean(1 / exp_) if len(keys) > 2 else np.nan
    alpha_hist = float(np.clip(1 / var_true, 5, 400)) if var_true == var_true and var_true > 0 else 400.0
    level_hist = pd.DataFrame({"年月": [f"{k // 100}-{k % 100:02d}" for k in keys],
                               "實際訂房 ÷ 模型預期": np.round(ratios, 3)})

    step(1.0, "完成")
    return Model(names=names, cap=cap, p_bar=p_bar, p_lo=p_lo, p_hi=p_hi,
                 A=params[0:K], B1=B1, B2=params[2 * K:3 * K], H=params[3 * K:6 * K].reshape(K, 3),
                 month_eff=month_eff, night_eff=night_eff, interact=interact, p_int=p_int,
                 quad=quad, pvals=pvals, checks=checks, n_spells=len(h),
                 cens_rate=float(np.mean(y >= R)), hist_range=(min(udates), max(udates)), notes=notes,
                 alpha_hist=alpha_hist, level_hist=level_hist)


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
# 3. 建議價格：最大化總淨收益（日期區間可跨月；平均淨收益限制可有可無，計算區間可自訂）
#    淨收益 = (價格 - 房型成本 c_i) × 成交房數
#    Lagrange：Σ (p-c)·S + μ·[Σ_{計算區間} (p-c-T)·S + B]
#    → 計算區間內的晚上：最大化 (p-c-θ)·S(p)，θ = T·μ/(1+μ)（所有晚上共用一個 θ）
#      計算區間外的晚上：最大化 (p-c)·S(p)（等於 θ = 0）
#      調整 θ 讓 Σ_{計算區間} (p-c-T)·S剩餘 + B ≥ 0，
#      B = Σ_{計算區間已成交} (成交價 - c - T)：已成交訂單的「平均淨收益盈餘」
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


def _posterior(n, E, cens, alpha):
    """u 的後驗（1 維網格）：先驗 Gamma(α, α) × 卜瓦松概似；設限的觀察用 P(D ≥ n)。"""
    u = np.linspace(0.02, 5, 2500)
    logp = (alpha - 1) * np.log(u) - alpha * u
    logp += (n[~cens].sum() * np.log(u)) - u * E[~cens].sum()
    if cens.any():
        logp += stats.poisson.logsf(n[cens][None] - 1, u[:, None] * E[cens][None]).sum(1)
    w = np.exp(logp - logp.max())
    w /= w.sum()
    cdf = np.cumsum(w)
    return dict(u=float((u * w).sum()), lo=float(u[np.searchsorted(cdf, 0.1)]),
                hi=float(u[np.searchsorted(cdf, 0.9)]), n=float(n.sum()), E=float(E.sum()),
                censored=int(cens.sum()), alpha=float(alpha), weight=float(E.sum() / (E.sum() + alpha)))


def update_level(model, holidays, otb, today, dates, alpha, spells=None):
    """
    貝氏更新「需求水準 u」：實際需求 = u × 歷史模型需求
      先驗 u ~ Gamma(α, α)（平均 1）
      資料：到今天為止實際的訂房數 n，與歷史模型預期 E 比較
      後驗用 1 維網格計算（沒有賣光時等於 Gamma(α+n, α+E)）
    兩種資料來源：
      1. 有「目前時段」（spells）：每一段價格分開算 E，和估計歷史模型的方法完全相同（精確）
      2. 只有「目前訂房」：每格用一個價格近似 —— 有成交用平均成交價，沒成交用目前牌價，
         都沒有才用歷史平均價 p̄（近似；中途調過價的格子會有誤差）
    """
    dates = list(dates)
    K, names, cap = len(model.names), model.names, model.cap
    cal = Calendar(holidays)
    ri = {n: i for i, n in enumerate(names)}
    CUM = np.cumsum(np.exp(XH @ model.H.T).T, axis=1)                  # K × 91，CUM[i, L] = Σ_{l≤L}
    CUM0 = np.concatenate([np.zeros((K, 1)), CUM], 1)                   # CUM0[i, L+1] = Σ_{l≤L}

    # ---------- 1. 用「目前時段」 ----------
    if spells is not None:
        sp = spells[spells["入住日"].isin(dates)]
        if len(sp):
            sd = list(sp["入住日"])
            months, ntypes = cal.features(sd)
            z = model.ctrl(months, ntypes)
            r = sp["房型"].map(ri).to_numpy(int)
            L_now = np.array([(d - today).days for d in sd])
            Ls = np.minimum(sp["開始提前天數"].to_numpy(int), LMAX)
            Le_raw = sp["結束提前天數"].to_numpy(int)
            Le = np.maximum(Le_raw, np.maximum(L_now, 0))                   # 不能超過今天
            n_clip = int((Le_raw < L_now).sum())
            ok = Ls >= Le
            pc = sp["牌價"].to_numpy(float) / 1000 - model.p_bar[r]
            mu = (np.exp(model.A[r] + model.B1[r] * pc + model.B2[r] * pc**2 + z)
                  * (CUM0[r, Ls + 1] - CUM0[r, np.minimum(Le, LMAX)]))
            y = sp["訂房數"].to_numpy(float)
            R = sp["時段開始剩餘房數"].to_numpy(float)
            ok &= mu > 1e-12
            if ok.any():
                out = _posterior(y[ok], mu[ok], y[ok] >= R[ok], alpha)
                out.update(source="目前時段（逐段精確計算）", cells=int(ok.sum()), n_pbar=0, n_clip=n_clip)
                return out

    # ---------- 2. 用「目前訂房」（近似） ----------
    if otb is None:
        return None
    o = otb[otb["入住日"].isin(dates)]
    if len(o) == 0:
        return None
    D = len(dates)
    sold, rev, px = np.zeros((D, K)), np.zeros((D, K)), np.full((D, K), np.nan)
    di = {d: k for k, d in enumerate(dates)}
    for _, row in o.iterrows():
        sold[di[row["入住日"]], ri[row["房型"]]] = row["已售房數"]
        rev[di[row["入住日"]], ri[row["房型"]]] = row["已售營收"] / 1000
        if "目前牌價" in row.index and pd.notna(row["目前牌價"]):
            px[di[row["入住日"]], ri[row["房型"]]] = row["目前牌價"] / 1000
    months, ntypes = cal.features(dates)
    z = model.ctrl(months, ntypes)
    L_now = np.array([(d - today).days for d in dates])
    # 已經過去的提前天數：L_now+1 … 90（入住日已過則為 0 … 90）
    start_L = np.clip(L_now + 1, 0, LMAX + 1)
    elapsed = CUM[:, LMAX][None] - np.where(start_L[:, None] > 0, CUM[:, np.clip(start_L - 1, 0, LMAX)].T, 0.0)
    fallback = np.where(np.isnan(px), model.p_bar[None], px)
    price = np.where(sold > 0, rev / np.maximum(sold, 1), fallback)
    n_pbar = int(((sold == 0) & np.isnan(px)).sum())                    # 用歷史平均價近似的格子數
    pc = price - model.p_bar[None]
    E = np.exp(model.A[None] + model.B1[None] * pc + model.B2[None] * pc**2 + z[:, None]) * elapsed
    use = E > 1e-9
    if not use.any():
        return None
    n, Eu, capu = sold[use], E[use], np.broadcast_to(cap[None], sold.shape)[use]
    out = _posterior(n, Eu, n >= capu, alpha)
    out.update(source="目前訂房（每格一個價格，近似）", cells=int(use.sum()), n_pbar=n_pbar, n_clip=0)
    return out


def price_range(model, holidays, start, end, today, otb, step, costs, bounds=None,
                target=None, buffer=0, window=None, level=1.0, level_info=None):
    """
    start, end : 定價日期區間（含頭尾）
    costs      : {房型: 每間夜成本（元）}
    target     : 平均淨收益目標（元／間夜）；None = 不設限，只做總淨收益最大化
    window     : 平均淨收益計算區間 (起, 迄)；None = 與定價區間相同
    回傳 dict：price_table, detail, summary, warnings。金額單位：元。
    """
    if end < start:
        raise InputError("定價區間的結束日早於開始日。")
    K, names, cap = len(model.names), model.names, model.cap
    c = np.array([float(costs.get(n, 0.0)) for n in names]) / 1000          # 千元
    use_target = target is not None
    P_dates = _date_range(start, end)
    W_dates = _date_range(*window) if (use_target and window) else (P_dates if use_target else [])
    if use_target and window and window[1] < window[0]:
        raise InputError("平均淨收益計算區間的結束日早於開始日。")
    dates = sorted(set(P_dates) | set(W_dates))
    if (dates[-1] - dates[0]).days + 1 > MAX_DAYS:
        raise InputError(f"定價區間與計算區間合計跨越超過 {MAX_DAYS} 天，請縮短區間。")
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
        warnings.append("計算區間包含已經過去的日期，但沒有「目前訂房」資料，這些日期會被當成沒有成交。")
    w_only = in_W & ~in_P & ~past
    if w_only.any():
        warnings.append(f"平均淨收益計算區間有 {int(w_only.sum())} 個未來日期不在定價區間內。"
                        "為了估計整段區間的平均淨收益，這些日期也一起規劃了價格（表中標示為「只計入平均」）。")

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
        if c[i] * 1000 >= hi:
            warnings.append(f"{n} 的成本（{c[i]*1000:,.0f} 元）不低於可設定的最高價，賣出一定虧損，請確認成本。")
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
        lev = np.exp(model.A[i] + model.B1[i] * pc + model.B2[i] * pc**2)
        lam = level * lev[None] * (np.exp(z) * CUM[i, Lc])[:, None]
        Si = expected_min(lam, remaining[:, i][:, None], cap[i])
        Si[~open_[:, i]] = 0.0
        S_list.append(Si)

    T = (target + buffer) / 1000 if use_target else 0.0
    B = (rev[in_W].sum() - ((c[None] + T) * sold)[in_W].sum()) if use_target else 0.0
    rows_ = np.arange(D)

    def choose(th):
        th_d = np.where(in_W, th, 0.0)[:, None]
        p, s = np.zeros((D, K)), np.zeros((D, K))
        for i in range(K):
            j = ((grids[i][None] - c[i] - th_d) * S_list[i]).argmax(1)
            p[:, i], s[:, i] = grids[i][j], S_list[i][rows_, j]
        return p, s

    def gap(th):
        p, s = choose(th)
        return ((p - c[None] - T) * s)[in_W].sum() + B

    feasible, theta = True, 0.0
    if use_target and gap(0.0) < 0:
        if T <= 0 or gap(T) < 0:
            theta, feasible = max(T, 0.0), False
            warnings.append("在目前的價格範圍內，預期平均淨收益追不到目標。建議價格已經盡量拉高平均淨收益；"
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
            usage.append("定價 + 計入平均")
        elif in_P[d]:
            usage.append("定價")
        else:
            usage.append("只計入平均")
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
                            "成本": int(round(c[i] * 1000)),
                            "預期再售出": round(float(Sx[d, i]), 2),
                            "預期新增營收": int(round(P[d, i] * Sx[d, i] * 1000)) if open_[d, i] else 0,
                            "預期新增淨收益": int(round((P[d, i] - c[i]) * Sx[d, i] * 1000)) if open_[d, i] else 0}
                           for d in range(D) for i, n in enumerate(names)])

    def totals(mask, p, s):
        rooms_ = sold[mask].sum() + s[mask].sum()
        money = (rev[mask].sum() + (p * s)[mask].sum()) * 1000
        net = money - ((sold + s) * c[None])[mask].sum() * 1000
        per = lambda v: v / rooms_ if rooms_ else None
        return dict(rooms=rooms_, rev=money, net=net, adr=per(money), net_avg=per(net))

    tp, tf = totals(in_P, P, Sx), totals(in_P, P0, S0)
    summary = {"需求預測方式": "貝氏更新" if level_info else "只用歷史模型"}
    if level_info:
        summary.update({"需求水準 u（後驗平均）": round(level_info["u"], 3),
                        "u 的 80% 區間": f"{level_info['lo']:.2f} ～ {level_info['hi']:.2f}",
                        "先驗強度 α": round(level_info["alpha"], 1),
                        "目前訂房：實際已售": level_info["n"],
                        "目前訂房：歷史模型預期": round(level_info["E"], 1),
                        "資料的權重": f"{level_info['weight']:.0%}",
                        "用歷史平均價近似的格子數": level_info.get("n_pbar", 0),
                        "貝氏更新的資料來源": level_info.get("source", "")})
    summary.update({"定價區間": f"{start} ～ {end}（{len(P_dates)} 晚）"})
    summary.update({
               "定價區間：已成交房數": int(sold[in_P].sum()),
               "定價區間：預期總房數": tp["rooms"],
               "定價區間：預期總淨收益": tp["net"],
               "定價區間：預期總營收": tp["rev"],
               "定價區間：預期平均淨收益（每間夜）": tp["net_avg"],
               "定價區間：預期平均成交價": tp["adr"]})
    if use_target:
        tw, twf = totals(in_W, P, Sx), totals(in_W, P0, S0)
        ws = sold[in_W].sum()
        wnet = (rev[in_W].sum() - (sold * c[None])[in_W].sum()) * 1000
        summary.update({
            "平均淨收益計算區間": f"{W_dates[0]} ～ {W_dates[-1]}（{len(W_dates)} 晚）",
            "平均淨收益目標": target,
            "計算用目標（含緩衝）": target + buffer,
            "計算區間：已成交房數": int(ws),
            "計算區間：已成交平均淨收益": wnet / ws if ws else None,
            "目前盈餘 B": B * 1000,
            "θ（每間房的虛擬成本）": theta * 1000,
            "計算區間：預期平均淨收益": tw["net_avg"],
            "不設限時：計算區間預期平均淨收益": twf["net_avg"],
            "不設限時：定價區間預期總淨收益": tf["net"],
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
    summ = pd.DataFrame({"項目": ["計算日期", "是否設定平均淨收益目標"] + list(s),
                         "數值": [result["today"], "是" if result["use_target"] else "否（只做總淨收益最大化）"]
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
