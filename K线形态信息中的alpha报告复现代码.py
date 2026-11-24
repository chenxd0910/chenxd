import json
import os
import time
import warnings

import numpy as np
import pandas as pd
from scipy import stats

import matplotlib
matplotlib.use("Agg")
import matplotlib.dates as mdates
import matplotlib.pyplot as plt
from matplotlib.ticker import PercentFormatter

warnings.filterwarnings("ignore")

ROOT = os.path.dirname(os.path.abspath(__file__))
RAW_DIR = os.path.join(ROOT, "dataset")
OUT_DIR = os.path.join(ROOT, "数据", "处理后数据")
RES_DIR = os.path.join(ROOT, "结果")
TAB_DIR = os.path.join(RES_DIR, "tables")
SER_DIR = os.path.join(RES_DIR, "series")
FIG_DIR = os.path.join(RES_DIR, "figures")

for d in (OUT_DIR, TAB_DIR, SER_DIR, FIG_DIR):
    os.makedirs(d, exist_ok=True)

FIELDS = ["open", "high", "low", "close", "vol", "amount",
          "adj_factor", "total_mv", "pb", "pe_ttm", "is_st"]
N_BASE = 24
N_PATTERN = 41472
BODY_BINS = [0.015, 0.03]


# 1. 数据读取与预处理
class Panel:
    """行情面板：dates、codes 和 fields 字典。"""

    def __init__(self, dates, codes, fields):
        self.dates = dates
        self.codes = codes
        self.fields = fields

    def __getitem__(self, key):
        return self.fields[key]

    @property
    def shape(self):
        return len(self.dates), len(self.codes)

def read_field(name):
    """读一个字段的沪深两个文件，拼成宽表。"""
    parts = []
    for mkt in ("sh", "sz"):
        path = os.path.join(RAW_DIR, f"{name}_{mkt}.csv")
        if not os.path.exists(path):
            continue
        df = pd.read_csv(path, index_col=0)
        df.index = pd.to_datetime(df.index, errors="coerce")
        df = df[df.index.notna()].sort_index()
        df = df[~df.index.duplicated(keep="last")]
        df = df.loc[:, ~df.columns.duplicated()]
        df = df.apply(pd.to_numeric, errors="coerce")
        parts.append(df.astype("float32"))
    df = pd.concat(parts, axis=1)
    df = df.loc[:, ~df.columns.duplicated()].sort_index()
    return df

def load_panel(start="2020-01-01", end="2025-12-31", min_obs=120):
    """对齐所有字段；有缓存就直接读缓存。"""
    cache = os.path.join(OUT_DIR, "panel.npz")
    if os.path.exists(cache):
        z = np.load(cache, allow_pickle=True)
        dates = pd.DatetimeIndex(z["dates"])
        codes = pd.Index(z["codes"].astype(str))
        fields = {k[2:]: z[k] for k in z.files if k.startswith("f_")}
        return Panel(dates, codes, fields)

    data = {f: read_field(f) for f in FIELDS}
    close = data["close"].loc[start:end]
    use = ((close > 0) & close.notna()).sum() >= min_obs
    close = close.loc[:, use]
    dates, codes = close.index, close.columns

    fields = {}
    for f in FIELDS:
        fields[f] = data[f].reindex(index=dates, columns=codes).to_numpy("float32")
    for f in ("open", "high", "low", "close"):
        fields[f][~(fields[f] > 0)] = np.nan
    for f in ("amount", "vol"):
        fields[f][~(fields[f] > 0)] = np.nan

    np.savez_compressed(cache, dates=dates.values.astype("datetime64[ns]"),
                        codes=np.asarray(codes, dtype=object),
                        **{"f_" + k: v for k, v in fields.items()})
    return Panel(dates, codes, fields)

def week_ends(dates):
    """每周最后一个交易日。"""
    s = pd.Series(dates, index=dates)
    g = s.groupby([dates.isocalendar().year, dates.isocalendar().week]).max()
    return pd.DatetimeIndex(sorted(g.values))

def month_ends(dates):
    """每月最后一个交易日。"""
    s = pd.Series(dates, index=dates)
    g = s.groupby([dates.year, dates.month]).max()
    return pd.DatetimeIndex(sorted(g.values))


# 2. K 线形态与 CSP 因子
def adj_price(panel):
    adj = panel["adj_factor"]
    return {f: panel[f] * adj for f in ("open", "high", "low", "close")}

def fwd_return(close, h=20):
    t, n = close.shape
    out = np.full((t, n), np.nan, dtype="float32")
    if t > h:
        out[:-h] = close[h:] / close[:-h] - 1.0
    return out

def excess_ret(ret):
    """减去当日全市场等权平均收益。"""
    good = np.isfinite(ret)
    cnt = good.sum(axis=1, keepdims=True)
    s = np.where(good, ret, 0.0).sum(axis=1, keepdims=True)
    mean = np.divide(s, cnt, out=np.full_like(s, np.nan, dtype="float32"),
                     where=cnt > 0)
    return ret - mean

def make_bars(panel, k):
    """把日线合并成 k 日 K 线（不重叠）。"""
    T, N = panel.shape
    O, H, L, C, A = (panel[f] for f in ("open", "high", "low", "close", "amount"))
    ends = np.arange(k - 1, T, k)
    starts = ends - k + 1
    m = len(ends)

    op = O[starts]
    cp = C[ends]
    prev = np.full((m, N), np.nan, dtype="float32")
    ok = ends - k >= 1
    prev[ok] = C[ends[ok] - k]

    hi = np.full((m, N), -np.inf, dtype="float32")
    lo = np.full((m, N), np.inf, dtype="float32")
    amt = np.zeros((m, N), dtype="float32")
    cnt = np.zeros((m, N), dtype="int16")
    for o in range(k):
        i = starts + o
        hi = np.fmax(hi, H[i])
        lo = np.fmin(lo, L[i])
        a = A[i]
        good = np.isfinite(a) & (a > 0)
        amt += np.where(good, a, 0).astype("float32")
        cnt += good.astype("int16")
    hi = np.where(np.isfinite(hi), hi, np.nan)
    lo = np.where(np.isfinite(lo), lo, np.nan)
    amt[cnt < k] = np.nan

    valid = (np.isfinite(op) & np.isfinite(hi) & np.isfinite(lo) &
             np.isfinite(cp) & np.isfinite(prev) & (prev > 0) &
             np.isfinite(amt) & (hi >= lo))
    body = (cp - op) / prev
    upper = (hi - np.fmax(op, cp)) / prev
    lower = (np.fmin(op, cp) - lo) / prev
    upper = np.where(np.isfinite(upper), np.maximum(upper, 0), np.nan)
    lower = np.where(np.isfinite(lower), np.maximum(lower, 0), np.nan)
    body = np.where(valid, body, np.nan)

    body_cls = np.digitize(np.abs(body), BODY_BINS).astype("int16")
    bull = cp >= op
    base = (bull.astype("int16") * 12 + body_cls * 4 +
            (upper >= 0.02).astype("int16") * 2 + (lower >= 0.02).astype("int16"))
    base = np.where(valid, base, -1).astype("int16")

    ma = pd.DataFrame(amt).rolling(20, min_periods=10).mean().shift(1).to_numpy()
    ratio = amt / ma
    vol = np.where(ratio >= 1.2, 2, np.where(ratio <= 0.8, 0, 1)).astype("int8")
    vol = np.where(np.isfinite(ratio), vol, 1).astype("int8")

    return dict(ends=ends, open=op, high=hi, low=lo, close=cp, prev=prev,
                amount=amt, base=base, vol=vol, body=body,
                upper=upper, lower=lower)

def make_codes(bars):
    """三根 K 线 + 成交量状态 -> 组合形态编码。"""
    b = bars["base"]
    if len(b) < 3:
        return np.empty((0, b.shape[1]), dtype="int32"), np.empty(0, dtype="int64")
    b1, b2, b3 = b[:-2], b[1:-1], b[2:]
    vc = bars["vol"][2:]
    good = (b1 >= 0) & (b2 >= 0) & (b3 >= 0)
    code = ((b1.astype("int32") * 24 + b2.astype("int32")) * 24 +
            b3.astype("int32")) * 3 + vc
    code = np.where(good, code, -1).astype("int32")
    return code, bars["ends"][2:]

def score_cutoffs(dates):
    """每月末可以使用的历史样本截止位置。"""
    idx = dates.get_indexer(month_ends(dates))
    idx = idx[idx >= 120]
    cut = idx - 20
    keep = cut > 0
    return idx[keep], cut[keep]

def fit_scores(code, end_idx, exc, cut, min_n=20, kappa=50):
    """每种形态的 IR 得分（带小样本收缩）。"""
    M = len(cut)
    scores = np.full((M, N_PATTERN), np.nan, dtype="float32")
    good = (code >= 0) & np.isfinite(exc)
    if not good.any():
        return scores
    c = code[good]
    e = exc[good].astype("float64")
    b = np.broadcast_to(end_idx[:, None], code.shape)[good]
    known = np.searchsorted(cut, b, side="left")
    keep = known < M
    c, e, known = c[keep], e[keep], known[keep]
    if len(c) == 0:
        return scores

    flat = known.astype("int64") * N_PATTERN + c
    n = M * N_PATTERN
    cnt = np.bincount(flat, minlength=n).reshape(M, N_PATTERN)
    s1 = np.bincount(flat, weights=e, minlength=n).reshape(M, N_PATTERN)
    s2 = np.bincount(flat, weights=e * e, minlength=n).reshape(M, N_PATTERN)
    cnt, s1, s2 = np.cumsum(cnt, 0), np.cumsum(s1, 0), np.cumsum(s2, 0)

    mean = np.divide(s1, cnt, out=np.zeros_like(s1), where=cnt > 0)
    var = np.divide(s2, cnt, out=np.zeros_like(s2), where=cnt > 0) - mean ** 2
    ir = mean / np.sqrt(np.maximum(var, 1e-8))
    sc = np.clip(ir * cnt / (cnt + kappa), -3, 3)
    sc[cnt < min_n] = np.nan
    scores[:] = sc.astype("float32")
    return scores

def map_scores(code, end_idx, scores, refit, dates, N):
    """把形态得分映射到每个交易日（最近一根已完成 K 线的得分）。"""
    out = np.full((len(dates), N), np.nan, dtype="float32")
    r = np.searchsorted(refit, end_idx, side="right") - 1
    r = np.broadcast_to(r[:, None], code.shape)
    good = (r >= 0) & (code >= 0)
    if good.any():
        val = np.full(code.shape, np.nan, dtype="float32")
        val[good] = scores[r[good], code[good]]
        out[end_idx] = val
    return pd.DataFrame(out).ffill().to_numpy(dtype="float32")

def build_csp(panel, k, exc, cut, refit):
    """单个周期的 CSP 因子。"""
    bars = make_bars(panel, k)
    code, end_idx = make_codes(bars)
    scores = fit_scores(code, end_idx, exc[end_idx], cut)
    factor = map_scores(code, end_idx, scores, refit, panel.dates, panel.shape[1])
    return dict(factor=factor, scores=scores, n_bars=len(bars["ends"]),
                n_patterns=int(len(np.unique(code[code >= 0]))),
                n_scored=int(np.isfinite(scores[-1]).sum()),
                coverage=float(np.isfinite(factor).mean()))

def combine_csp(factors):
    """多个周期的因子 z 值等权平均。"""
    zs = []
    for f in factors:
        mu = np.nanmean(f, axis=1, keepdims=True)
        sd = np.nanstd(f, axis=1, keepdims=True)
        zs.append(np.clip((f - mu) / sd, -3, 3))
    arr = np.stack(zs)
    cnt = np.isfinite(arr).sum(0)
    out = np.nanmean(arr, axis=0)
    out[cnt < 2] = np.nan
    return out.astype("float32")

def classic_events(panel):
    """十字星、锤头线、倒锤头线 + 位置 + 放量。"""
    bars = make_bars(panel, 1)
    body, up, low = np.abs(bars["body"]), bars["upper"], bars["lower"]
    doji = (body <= 0.02) & (up >= 0.03) & (low >= 0.03)
    hammer = (body <= 0.02) & (up <= 0.02) & (low >= 0.05)
    inv = (body <= 0.02) & (up >= 0.05) & (low <= 0.02)

    hi = pd.DataFrame(bars["high"]).rolling(250, min_periods=120).max().to_numpy()
    lo = pd.DataFrame(bars["low"]).rolling(250, min_periods=120).min().to_numpy()
    top = bars["close"] >= 0.9 * hi
    bottom = bars["close"] <= 1.1 * lo
    vma = pd.DataFrame(bars["amount"]).rolling(20, min_periods=10).mean().shift(1).to_numpy()
    volup = bars["amount"] >= 1.2 * vma

    bull = bottom & (doji | hammer) & volup
    bear = top & (doji | inv) & volup
    ind = bull.astype("float32") - bear.astype("float32")
    score = pd.DataFrame(ind).rolling(20, min_periods=1).sum().to_numpy("float32")
    score = pd.DataFrame(score).ffill().to_numpy("float32")
    return dict(doji=doji, hammer=hammer, inv=inv, top=top, bottom=bottom,
                volup=volup, bull=bull, bear=bear, score=score)

def pattern_dist(panel):
    """24 种单根 K 线的样本占比。"""
    b = make_bars(panel, 1)["base"].ravel()
    b = b[b >= 0]
    cnt = np.bincount(b, minlength=24).astype(float)
    share = cnt / cnt.sum()
    names = []
    for i in range(24):
        bull = i // 12
        body = (i % 12) // 4
        up = (i % 4) // 2
        low = i % 2
        names.append(f"{'阳' if bull else '阴'}-{['小','中','大'][body]}-"
                     f"{'长上' if up else '短上'}{'长下' if low else '短下'}")
    return pd.DataFrame({"code": range(24), "name": names,
                         "count": cnt.astype(int), "share": share})


# 3. 绩效指标
def ann_ret(r):
    r = r.dropna()
    return (1 + r).prod() ** (52 / len(r)) - 1

def ann_vol(r):
    r = r.dropna()
    return r.std(ddof=1) * np.sqrt(52)

def max_dd(r):
    r = r.dropna()
    nav = (1 + r).cumprod()
    return (nav / nav.cummax() - 1).min()

def perf_row(r, name):
    r = r.dropna()
    nav = (1 + r).cumprod()
    a, v, md = float(ann_ret(r)), float(ann_vol(r)), float(max_dd(r))
    return {"名称": name, "年化收益率": a, "年化波动率": v,
            "夏普比率": float((a - 0.02) / v) if v else np.nan,
            "最大回撤": md, "卡玛比率": float(a / abs(md)) if md else np.nan,
            "周胜率": (r > 0).mean(), "累计净值": nav.iloc[-1] if len(nav) else np.nan,
            "样本周数": len(r)}

def yearly_ret(r):
    r = r.dropna()
    return r.groupby(r.index.year).apply(lambda x: (1 + x).prod() - 1)

def dd_series(r):
    r = r.dropna()
    nav = (1 + r).cumprod()
    return nav / nav.cummax() - 1

def save_table(df, path, index=False):
    df.to_csv(path, index=index, encoding="utf-8-sig", float_format="%.6f")


# 4. 周度多因子模型与回测
def weekly_data(panel, bench_close):
    close, amount = panel["close"], panel["amount"]
    st, adj = panel["is_st"], panel["adj_factor"]
    valid = np.isfinite(close) & (close > 0)
    listed = np.cumsum(valid, axis=0)
    is_st = np.isfinite(st) & (st > 0.5)

    wd = week_ends(panel.dates)
    idx = panel.dates.get_indexer(wd)
    keep = (idx >= 0) & (idx + 5 < len(panel.dates))
    wd, idx = wd[keep], idx[keep]

    adjc = close * adj
    nret = (adjc[idx + 5] / adjc[idx] - 1).astype("float32")
    ok = valid[idx] & np.isfinite(amount[idx]) & (~is_st[idx]) & (listed[idx] >= 250)

    bm = bench_close.reindex(panel.dates).ffill()
    bm_ret = pd.Series(bm.iloc[idx + 5].values / bm.iloc[idx].values - 1, index=wd)
    all_ret = pd.Series(np.nanmean(nret, axis=1), index=wd)
    return {"wd": wd, "idx": idx, "nret": nret, "bm": bm_ret,
            "all": all_ret, "ok": ok}

def ind_neutral(x, group):
    out = x.copy()
    m = np.isfinite(x) & (group >= 0)
    if not m.any():
        return out
    g, v = group[m], x[m]
    n = g.max() + 1
    cnt = np.bincount(g, minlength=n).astype(float)
    s = np.bincount(g, weights=v, minlength=n)
    mean = np.divide(s, cnt, out=np.zeros(n), where=cnt > 0)
    out[m] = v - mean[g]
    return out

def prep_factor(x, group):
    """缩尾 + 标准化 + 行业中性化。"""
    m = np.isfinite(x)
    out = np.full_like(x, np.nan, dtype=float)
    if m.sum() < 10:
        return out
    lo, hi = np.nanquantile(x[m], [0.01, 0.99])
    v = np.clip(x, lo, hi)
    mu, sd = np.nanmean(v[m]), np.nanstd(v[m])
    if not np.isfinite(sd) or sd == 0:
        out[m] = 0
        return out
    z = (v - mu) / sd
    z = ind_neutral(z, group)
    sd = np.nanstd(z[m])
    if np.isfinite(sd) and sd > 0:
        z = (z - np.nanmean(z[m])) / sd
    out[m] = z[m]
    return out

def industry_code(codes):
    info = pd.read_csv(os.path.join(RAW_DIR, "stk_company_info.csv"), encoding="utf-8-sig")
    ind = info.dropna(subset=["ts_code", "l1_name"]).drop_duplicates("ts_code")
    ind = ind.set_index("ts_code")["l1_name"]
    names = ind.reindex(codes).fillna("未知").astype(str)
    uniq = {n: i for i, n in enumerate(sorted(names.unique()))}
    return names.map(uniq).values

def control_factors(panel):
    """相关性/回归用的常用因子。"""
    close = adj_price(panel)["close"]
    ret = close / np.roll(close, 1, axis=0) - 1
    mom = np.full_like(close, np.nan)
    mom[20:] = close[20:] / close[:-20] - 1
    rev = np.full_like(close, np.nan)
    rev[5:] = close[5:] / close[:-5] - 1
    vol = pd.DataFrame(ret).rolling(20, min_periods=15).std().to_numpy("float32")
    turn = np.where(panel["total_mv"] > 0, panel["amount"] / panel["total_mv"], np.nan)
    amihud = np.where(panel["amount"] > 0, np.abs(ret) / panel["amount"] * 1e8, np.nan)
    amihud = pd.DataFrame(amihud).rolling(20, min_periods=15).mean().to_numpy("float32")
    size = np.log(panel["total_mv"])
    bp = 1 / panel["pb"]
    return {"市值": size, "BP": bp, "动量": mom, "反转": rev,
            "波动率": vol, "换手率": turn, "非流动性": amihud}

def prepare_scores(panel, wk, factors, controls, ev):
    """每个调仓日的因子 z 值与综合得分。"""
    group = industry_code(panel.codes)
    n, m = len(wk["idx"]), panel.shape[1]
    z = {k: np.full((n, m), np.nan, dtype="float32")
         for k in ("size", "bp", "csp5", "cspc", "event")}
    for step, i in enumerate(wk["idx"]):
        size = np.log(panel["total_mv"][i])
        bp = np.where(panel["pb"][i] > 0, 1 / panel["pb"][i], np.nan)
        z["size"][step] = prep_factor(size, group)
        z["bp"][step] = prep_factor(bp, group)
        z["csp5"][step] = prep_factor(factors["CSP_5"][i], group)
        z["cspc"][step] = prep_factor(factors["CSP_复合"][i], group)
        z["event"][step] = prep_factor(ev["score"][i], group)

    smooth = pd.DataFrame(z["cspc"]).rolling(4, min_periods=1).mean().to_numpy("float32")
    return {
        "bm": 0.5 * (-z["size"]) + 0.5 * z["bp"],
        "rm5": (-z["size"] + z["bp"] + z["csp5"]) / 3,
        "rmc": (-z["size"] + z["bp"] + z["cspc"]) / 3,
        "rme": (-z["size"] + z["bp"] + z["event"]) / 3,
        "rms": (-z["size"] + z["bp"] + smooth) / 3,
        "z": z, "smooth": smooth,
    }

def pick_top(score, ok, top_frac=0.1, min_n=30, prev=None, buffer=0.0):
    cand = np.flatnonzero(ok & np.isfinite(score))
    if len(cand) == 0:
        return np.empty(0, dtype=int)
    k = max(min_n, int(round(top_frac * len(cand))))
    k = min(k, len(cand))
    if buffer > 0 and prev is not None and len(prev) > 0:
        thr = np.nanquantile(score[cand], 1 - min(top_frac + buffer, 0.99))
        keep = prev[np.isfinite(score[prev]) & (score[prev] >= thr)]
        if len(keep) >= k:
            return keep[np.argsort(-score[keep])[:k]]
        rest = cand[~np.isin(cand, keep)]
        need = k - len(keep)
        if need <= 0 or len(rest) == 0:
            return keep
        add = rest[np.argpartition(-score[rest], need - 1)[:need]]
        return np.concatenate([keep, add])
    return cand[np.argpartition(-score[cand], k - 1)[:k]]

def run_strategy(wk, score, name, top_frac=0.1, fee=0.0015, reb=1, buffer=0.0):
    """周度调仓回测。score 是按周排列的得分矩阵。"""
    wd = wk["wd"]
    ok, nret = wk["ok"], wk["nret"]
    n = len(wd)
    ret = np.full(n, np.nan)
    gross = np.full(n, np.nan)
    turn = np.full(n, np.nan)
    hold_n = np.full(n, np.nan)
    holds, prev = np.empty(0, dtype=int), np.empty(0, dtype=int)

    for step in range(n):
        if step % reb == 0 or len(holds) == 0:
            pick = pick_top(score[step], ok[step], top_frac, prev=holds, buffer=buffer)
            pick = pick[np.isfinite(nret[step, pick])]
            if len(pick) == 0:
                if len(holds) == 0:
                    continue
                pick = holds
            if len(prev) == 0:
                to = 1.0
            else:
                to = 1 - np.intersect1d(prev, pick).size / max(len(pick), 1)
            turn[step] = to
            cost = 2 * fee * to
            prev, holds = pick, pick
        hold_n[step] = len(holds)
        if len(holds) == 0:
            continue
        r = nret[step, holds]
        r = r[np.isfinite(r)]
        if len(r) == 0:
            continue
        gross[step] = r.mean()
        ret[step] = gross[step] - (2 * fee * turn[step] if np.isfinite(turn[step]) else 0)

    return {"name": name, "ret": pd.Series(ret, index=wd),
            "gross": pd.Series(gross, index=wd),
            "turnover": pd.Series(turn, index=wd),
            "holdings": pd.Series(hold_n, index=wd)}

def strategy_table(results, bench):
    rows = []
    for res in results:
        s = perf_row(res["ret"], res["name"])
        r = res["ret"].dropna()
        ex = (r - bench.reindex(r.index)).dropna()
        s["年化超额(vs沪深300)"] = ex.mean() * 52
        te = ex.std(ddof=1) * np.sqrt(52) if len(ex) > 2 else np.nan
        s["跟踪误差(vs沪深300)"] = te
        s["信息比率(vs沪深300)"] = s["年化超额(vs沪深300)"] / te if te else np.nan
        s["平均换手率"] = res["turnover"].dropna().mean()
        s["平均持仓数"] = res["holdings"].dropna().mean()
        rows.append(s)
    rows.append(perf_row(bench, "沪深300"))
    return pd.DataFrame(rows)


# 5. 因子检验
def rank_ic(a, b, min_n=100):
    m = np.isfinite(a) & np.isfinite(b)
    if m.sum() < min_n:
        return np.nan
    ra, rb = stats.rankdata(a[m]), stats.rankdata(b[m])
    if ra.std() == 0 or rb.std() == 0:
        return np.nan
    return np.corrcoef(ra, rb)[0, 1]

def ic_series(f, y, idx):
    return pd.Series([rank_ic(f[i], y[i]) for i in idx], index=idx)

def ic_summary(ic):
    ic = ic.dropna()
    m, s = ic.mean(), ic.std(ddof=1)
    return {"IC均值": m, "IC标准差": s, "ICIR": m / s if s else np.nan,
            "IC为正比例": (ic > 0).mean(),
            "t值": m / s * np.sqrt(len(ic)) if s else np.nan, "样本期数": len(ic)}

def nw_t(x, lags=4):
    x = np.asarray(x, dtype=float)
    x = x[np.isfinite(x)]
    n = len(x)
    if n < 10:
        return np.nan
    mu = x.mean()
    e = x - mu
    var = e @ e / n
    for lag in range(1, min(lags, n - 1) + 1):
        w = 1 - lag / (lags + 1)
        var += 2 * w * (e[lag:] @ e[:-lag]) / n
    return mu / np.sqrt(var / n) if var > 0 else np.nan

def fmb(y, X, dates, lags=4):
    """逐周横截面回归，再做 Newey-West。"""
    names = list(X)
    betas, r2 = [], []
    for t in range(len(y)):
        cols = [X[k][t] for k in names]
        m = np.isfinite(y[t])
        for c in cols:
            m &= np.isfinite(c)
        if m.sum() < 50:
            continue
        Y = y[t][m].astype(float)
        M = np.column_stack([np.ones(m.sum())] + [c[m] for c in cols]).astype(float)
        mu, sd = M[:, 1:].mean(0), M[:, 1:].std(0)
        sd[sd == 0] = 1
        M[:, 1:] = (M[:, 1:] - mu) / sd
        beta = np.linalg.lstsq(M, Y, rcond=None)[0]
        res = Y - M @ beta
        betas.append(beta[1:])
        r2.append(1 - (res ** 2).sum() / ((Y - Y.mean()) ** 2).sum())
    bdf = pd.DataFrame(betas, columns=names)
    rows = []
    for c in names:
        s = bdf[c]
        rows.append({"变量": c, "平均系数": s.mean(), "标准差": s.std(ddof=1),
                     "NW t值": nw_t(s, lags), "t值": s.mean() / s.std(ddof=1) * np.sqrt(len(s)),
                     "系数为正比例": (s > 0).mean(), "期数": len(s),
                     "平均R2": np.nanmean(r2)})
    return pd.DataFrame(rows)

def decile_test(wk, factor, n_group=10):
    """因子十分组，返回各组周收益和多空收益。"""
    wd, ok, nret = wk["wd"], wk["ok"], wk["nret"]
    out = np.full((len(wd), n_group), np.nan)
    for step, i in enumerate(wk["idx"]):
        f, r = factor[i], nret[step]
        m = ok[step] & np.isfinite(f) & np.isfinite(r)
        if m.sum() < 100:
            continue
        rank = pd.Series(f[m]).rank(method="first").values
        g = np.minimum((rank - 1) / len(rank) * n_group, n_group - 1).astype(int)
        rr = r[m]
        for j in range(n_group):
            if (g == j).any():
                out[step, j] = rr[g == j].mean()
    cols = [f"G{i+1}" for i in range(n_group)]
    gdf = pd.DataFrame(out, index=wd, columns=cols)
    return gdf, gdf.iloc[:, -1] - gdf.iloc[:, 0]

def group_table(gdf):
    rows = []
    for c in gdf.columns:
        r = gdf[c].dropna()
        a, v = (1 + r).prod() ** (52 / len(r)) - 1, r.std(ddof=1) * np.sqrt(52)
        rows.append({"分组": c, "年化收益": a, "年化波动": v,
                     "夏普比率": a / v if v else np.nan, "周胜率": (r > 0).mean()})
    df = pd.DataFrame(rows)
    rho = stats.spearmanr(range(len(df)), df["年化收益"]).statistic
    return df, rho

def long_short(ls):
    r = ls.dropna()
    a, v = (1 + r).prod() ** (52 / len(r)) - 1, r.std(ddof=1) * np.sqrt(52)
    return {"年化收益": a, "年化波动": v, "夏普比率": a / v if v else np.nan,
            "最大回撤": ((1 + r).cumprod() / (1 + r).cumprod().cummax() - 1).min(),
            "周胜率": (r > 0).mean(),
            "t值": r.mean() / r.std(ddof=1) * np.sqrt(len(r))}

def corr_table(f, controls, idx):
    rows = []
    for name, x in controls.items():
        cs = [rank_ic(f[i], x[i]) for i in idx]
        cs = [c for c in cs if np.isfinite(c)]
        rows.append({"对照因子": name, "平均相关系数": np.mean(cs),
                     "相关系数标准差": np.std(cs, ddof=1), "期数": len(cs)})
    return pd.DataFrame(rows)

def event_stats(mask, excess, dates, name):
    ii, jj = np.nonzero(mask & np.isfinite(excess))
    if len(ii) == 0:
        return None, None
    ex = excess[ii, jj].astype(float)
    df = pd.DataFrame({"date": dates[ii], "excess": ex})
    row = {"事件类型": name, "样本数": len(df), "超额收益均值": ex.mean(),
           "超额收益中位数": np.median(ex), "胜率": (ex > 0).mean(),
           "t值": ex.mean() / ex.std(ddof=1) * np.sqrt(len(ex)),
           "平均发生股票数/期": len(df) / df["date"].nunique()}
    yr = df.groupby(df["date"].dt.year)["excess"].agg(["mean", "count"])
    yr["win_rate"] = df.groupby(df["date"].dt.year)["excess"].apply(lambda x: (x > 0).mean())
    yr = yr.rename(columns={"mean": "超额收益均值", "count": "样本数"})
    yr["事件类型"] = name
    return row, yr.reset_index().rename(columns={"date": "年份"})

def event_table(panel, excess):
    """经典形态事件研究表。"""
    ev = classic_events(panel)
    combos = {
        "底部十字星+放量": ev["bottom"] & ev["doji"] & ev["volup"],
        "底部锤头线+放量": ev["bottom"] & ev["hammer"] & ev["volup"],
        "顶部十字星+放量": ev["top"] & ev["doji"] & ev["volup"],
        "顶部倒锤头线+放量": ev["top"] & ev["inv"] & ev["volup"],
        "底部反转形态(合并)": ev["bull"],
        "顶部反转形态(合并)": ev["bear"],
    }
    rows, years = [], []
    for name, mask in combos.items():
        row, yr = event_stats(mask, excess, panel.dates, name)
        if row:
            rows.append(row)
            years.append(yr)
    return ev, pd.DataFrame(rows), pd.concat(years, ignore_index=True)


# 6. 主流程
def nav_from_ret(ret):
    r = ret.dropna()
    nav = (1 + r).cumprod()
    return pd.concat([pd.Series([1.0], index=[r.index[0]]), nav])

def jsonable(x):
    if isinstance(x, dict):
        return {str(k): jsonable(v) for k, v in x.items()}
    if isinstance(x, (list, tuple)):
        return [jsonable(v) for v in x]
    if isinstance(x, (np.floating, float)):
        return None if not np.isfinite(x) else float(x)
    if isinstance(x, (np.integer,)):
        return int(x)
    if isinstance(x, (np.bool_,)):
        return bool(x)
    if isinstance(x, pd.Timestamp):
        return x.strftime("%Y-%m-%d")
    if isinstance(x, np.ndarray):
        return jsonable(x.tolist())
    return x

def run_pipeline():
    t0 = time.time()
    print("读数据 ...")
    panel = load_panel()
    bm = pd.read_csv(os.path.join(RAW_DIR, "benchmark.csv"), index_col=0, parse_dates=True)
    print("面板", panel.shape, panel.dates.min().date(), "~", panel.dates.max().date())

    adj = adj_price(panel)
    labels = {h: excess_ret(fwd_return(adj["close"], h)) for h in (5, 10, 20, 60)}

    # ---- 构造 10 个周期的 CSP
    refit, cut = score_cutoffs(panel.dates)
    print("形态按月重估：", len(cut), "期")
    factors, diag = {}, []
    for k in range(1, 11):
        s = time.time()
        res = build_csp(panel, k, labels[20], cut, refit)
        factors[f"CSP_{k}"] = res["factor"]
        diag.append({"k": k, "n_bars": res["n_bars"],
                     "n_patterns_observed": res["n_patterns"],
                     "n_patterns_scored": res["n_scored"],
                     "coverage_all": res["coverage"],
                     "mean": float(np.nanmean(res["factor"])),
                     "std": float(np.nanstd(res["factor"])),
                     "耗时秒": round(time.time() - s, 2)})
        print(f"  CSP_{k}: 覆盖 {res['coverage']:.1%}")
    factors["CSP_复合"] = combine_csp([factors[f"CSP_{k}"] for k in range(1, 11)])
    diag.append({"k": "复合", "n_bars": np.nan, "n_patterns_observed": np.nan,
                 "n_patterns_scored": np.nan,
                 "coverage_all": float(np.isfinite(factors["CSP_复合"]).mean()),
                 "mean": float(np.nanmean(factors["CSP_复合"])),
                 "std": float(np.nanstd(factors["CSP_复合"])), "耗时秒": np.nan})
    diag_df = pd.DataFrame(diag)
    save_table(diag_df, os.path.join(TAB_DIR, "t01_因子构建诊断.csv"))

    # ---- 因子检验
    wk = weekly_data(panel, bm["close"])
    dates = panel.dates
    wd = week_ends(dates)
    wd = wd[(wd >= "2021-01-01") & (wd <= "2025-11-28")]
    wi = dates.get_indexer(wd)
    md = month_ends(dates)
    md = md[(md >= "2021-01-01") & (md <= "2025-11-28")]
    mi = dates.get_indexer(md)
    print("因子检验：周度", len(wi), "期，月度", len(mi), "期")

    sel = {k: factors[k] for k in ("CSP_1", "CSP_5", "CSP_10", "CSP_复合")}
    lab = {5: "未来5日", 10: "未来10日", 20: "未来20日", 60: "未来60日"}
    rows = []
    for fn, f in sel.items():
        for h in (5, 10, 20, 60):
            rows.append({"因子": fn, "预测窗口": lab[h],
                         **ic_summary(ic_series(f, labels[h], wi))})
    ic_tbl = pd.DataFrame(rows)

    ic_m = {k: ic_series(f, labels[20], mi) for k, f in sel.items()}
    ic_month = pd.DataFrame([{"因子": k, **ic_summary(v)} for k, v in ic_m.items()])
    ic_w = {k: ic_series(f, labels[20], wi) for k, f in sel.items()}
    ic_df = pd.DataFrame({k: v.values for k, v in ic_w.items()}, index=dates[wi])
    ic_mdf = pd.DataFrame({k: v.values for k, v in ic_m.items()}, index=dates[mi])

    groups, ls = {}, {}
    gt_all, ls_rows, mono = [], [], {}
    for name in ("CSP_5", "CSP_复合"):
        g, l = decile_test(wk, factors[name])
        groups[name], ls[name] = g, l
        gt, rho = group_table(g)
        mono[name] = rho
        gt.insert(0, "因子", name)
        gt_all.append(gt)
        row = long_short(l)
        row["单调性"] = rho
        row["组合"] = name
        ls_rows.append({k: row[k] for k in ("组合", "年化收益", "年化波动", "夏普比率",
                                            "最大回撤", "周胜率", "t值", "单调性")})
    group_tbl = pd.concat(gt_all, ignore_index=True)
    ls_tbl = pd.DataFrame(ls_rows)

    controls = control_factors(panel)
    fmb_rows = []
    y = wk["nret"] * 100
    for name in ("CSP_5", "CSP_复合"):
        X = {"CSP": factors[name][wk["idx"]]}
        for c in ("市值", "BP", "动量", "反转", "波动率"):
            X[c] = controls[c][wk["idx"]]
        res = fmb(y, X, wk["wd"])
        res.insert(0, "因子", name)
        fmb_rows.append(res)
    fmb_tbl = pd.concat(fmb_rows, ignore_index=True)
    corr_tbl = corr_table(factors["CSP_5"], controls, wi)

    save_table(ic_tbl, os.path.join(TAB_DIR, "t02_IC汇总.csv"))
    save_table(ic_month, os.path.join(TAB_DIR, "t03_月度IC.csv"))
    save_table(group_tbl, os.path.join(TAB_DIR, "t04_分组收益.csv"))
    save_table(ls_tbl, os.path.join(TAB_DIR, "t05_多空组合.csv"))
    save_table(fmb_tbl, os.path.join(TAB_DIR, "t06_FamaMacBeth.csv"))
    save_table(corr_tbl, os.path.join(TAB_DIR, "t07_因子相关性.csv"))
    ic_df.to_csv(os.path.join(SER_DIR, "ic_weekly.csv"), encoding="utf-8-sig", float_format="%.6f")
    ic_mdf.to_csv(os.path.join(SER_DIR, "ic_monthly.csv"), encoding="utf-8-sig", float_format="%.6f")

    # ---- 事件研究
    print("经典形态事件研究 ...")
    ev, ev_tbl, ev_year = event_table(panel, labels[20])
    save_table(ev_tbl, os.path.join(TAB_DIR, "t08_经典形态事件研究.csv"))
    save_table(ev_year, os.path.join(TAB_DIR, "t09_事件研究分年度.csv"))

    # ---- 回测
    scores = prepare_scores(panel, wk, factors, controls, ev)
    specs = [("基准模型(Size+BP)", "bm"),
             ("研究模型(Size+BP+复合CSP)", "rmc"),
             ("研究模型(Size+BP+CSP5)", "rm5"),
             ("研究模型(Size+BP+经典形态)", "rme"),
             ("改进模型(复合CSP平滑)", "rms")]
    results = {}
    for name, key in specs:
        results[name] = run_strategy(wk, scores[key], name)
        results[name]["nav"] = nav_from_ret(results[name]["ret"])

    perf = strategy_table(list(results.values()), wk["bm"])
    all_row = perf_row(wk["all"].dropna(), "全A等权")
    all_row["年化超额(vs沪深300)"] = (wk["all"].dropna() -
                                wk["bm"].reindex(wk["all"].dropna().index)).dropna().mean() * 52
    perf = pd.concat([perf, pd.DataFrame([all_row])], ignore_index=True)
    save_table(perf, os.path.join(TAB_DIR, "t10_回测绩效.csv"))

    rm = results["研究模型(Size+BP+复合CSP)"]["ret"]
    bm_ret = results["基准模型(Size+BP)"]["ret"]
    ex_bm = (rm - bm_ret).dropna()
    extra = {"相对基准模型年化超额": ex_bm.mean() * 52,
             "相对基准模型超额胜率": (ex_bm > 0).mean(),
             "相对基准模型信息比率": ex_bm.mean() * 52 / (ex_bm.std(ddof=1) * np.sqrt(52)),
             "相对沪深300年化超额": (rm - wk["bm"].reindex(rm.index)).dropna().mean() * 52}

    ret_all = pd.DataFrame({"沪深300": wk["bm"], "全A等权": wk["all"],
                            **{k: v["ret"] for k, v in results.items()}})
    ret_all.to_csv(os.path.join(SER_DIR, "weekly_returns.csv"),
                   encoding="utf-8-sig", float_format="%.6f")
    nav = pd.DataFrame({k: v["nav"] for k, v in results.items()})
    nav["沪深300"] = (1 + wk["bm"].dropna()).cumprod()
    nav["全A等权"] = (1 + wk["all"].dropna()).cumprod()
    nav.to_csv(os.path.join(SER_DIR, "nav.csv"), encoding="utf-8-sig", float_format="%.6f")
    dd = pd.DataFrame({k: dd_series(v["ret"]) for k, v in results.items()})
    dd["沪深300"] = dd_series(wk["bm"])
    dd.to_csv(os.path.join(SER_DIR, "drawdown.csv"), encoding="utf-8-sig", float_format="%.6f")

    yearly = {k: yearly_ret(v["ret"]) for k, v in results.items()}
    yearly["沪深300"] = yearly_ret(wk["bm"])
    yearly["全A等权"] = yearly_ret(wk["all"])
    yearly_df = pd.DataFrame(yearly)
    yearly_df.index.name = "年份"
    save_table(yearly_df, os.path.join(TAB_DIR, "t11_分年度收益.csv"), index=True)

    for name, g in groups.items():
        tag = "csp5" if name == "CSP_5" else "comp"
        g.to_csv(os.path.join(SER_DIR, f"group_ret_{tag}.csv"),
                 encoding="utf-8-sig", float_format="%.6f")
        ls[name].to_csv(os.path.join(SER_DIR, f"ls_{tag}.csv"),
                        encoding="utf-8-sig", float_format="%.6f", header=True)

    # ---- 敏感性
    sens = []
    for top in (0.05, 0.10, 0.20):
        for fee in (0.0, 0.0015, 0.003):
            for reb in (1, 2, 4):
                s = perf_row(run_strategy(wk, scores["rmc"], "研究模型", top_frac=top,
                                          fee=fee, reb=reb)["ret"], "研究模型")
                sens.append({"组合分位": top, "单边成本": fee, "调仓周期(周)": reb,
                             "年化收益": s["年化收益率"], "年化波动": s["年化波动率"],
                             "夏普比率": s["夏普比率"], "最大回撤": s["最大回撤"]})
    sens_df = pd.DataFrame(sens)
    save_table(sens_df, os.path.join(TAB_DIR, "t12_参数敏感性.csv"))

    wrows = []
    for w in (0.0, 0.2, 1 / 3, 0.5, 0.8):
        sc = (1 - w) * (-scores["z"]["size"] + scores["z"]["bp"]) / 2 + w * scores["z"]["cspc"]
        res = run_strategy(wk, sc, f"CSP权重{w:.0%}")
        s = perf_row(res["ret"], "研究模型")
        wrows.append({"CSP权重": w, "年化收益": s["年化收益率"],
                      "年化波动": s["年化波动率"], "夏普比率": s["夏普比率"],
                      "最大回撤": s["最大回撤"],
                      "平均换手率": float(res["turnover"].mean())})
    wsen_df = pd.DataFrame(wrows)
    save_table(wsen_df, os.path.join(TAB_DIR, "t13_因子权重敏感性.csv"))

    crows = []
    for name, key in (("基准模型(Size+BP)", "bm"),
                      ("研究模型(Size+BP+复合CSP)", "rmc"),
                      ("改进模型(复合CSP平滑)", "rms")):
        gross = ann_ret(run_strategy(wk, scores[key], name, fee=0.0)["ret"])
        net = run_strategy(wk, scores[key], name)
        n = ann_ret(net["ret"])
        to = float(net["turnover"].mean())
        crows.append({"模型": name, "毛年化收益": gross, "净年化收益": n,
                      "成本拖累(年化)": gross - n, "平均周换手率": to,
                      "年化换手倍数": to * 52, "假设单边成本": 0.0015})
    cost_df = pd.DataFrame(crows)
    save_table(cost_df, os.path.join(TAB_DIR, "t14_成本归因.csv"))

    # ---- 处理后数据
    print("保存因子与统计 ...")
    fac = pd.DataFrame(scores["z"]["csp5"], index=wk["wd"], columns=panel.codes)
    fac.index.name = "date"
    fac.to_csv(os.path.join(OUT_DIR, "因子值_周度_CSP.csv"),
               encoding="utf-8-sig", float_format="%.4f")
    np.savez_compressed(os.path.join(OUT_DIR, "csp_factors_daily.npz"),
                        dates=panel.dates.values.astype("datetime64[ns]"),
                        codes=np.asarray(panel.codes, dtype=object),
                        CSP_1=factors["CSP_1"], CSP_5=factors["CSP_5"],
                        CSP_10=factors["CSP_10"], CSP_composite=factors["CSP_复合"])
    valid = np.isfinite(panel["close"]) & (panel["close"] > 0)
    stats = pd.DataFrame({"日期": panel.dates, "有效交易股票数": valid.sum(1),
                          "因子覆盖股票数": np.isfinite(factors["CSP_复合"]).sum(1)})
    save_table(stats, os.path.join(OUT_DIR, "样本覆盖统计.csv"))
    dist = pattern_dist(panel)
    save_table(dist, os.path.join(TAB_DIR, "t15_单K线形态分布.csv"))

    # ---- 汇总
    summary = {
        "数据": {"股票数": int(panel.shape[1]), "交易日数": int(panel.shape[0]),
                 "起始日": panel.dates.min(), "结束日": panel.dates.max(),
                 "平均每日有效股票数": float(valid.sum(1).mean())},
        "因子构建": diag_df.to_dict("records"),
        "IC": ic_tbl.to_dict("records"),
        "月度IC": ic_month.to_dict("records"),
        "分组": group_tbl.to_dict("records"),
        "多空": ls_tbl.to_dict("records"),
        "FamaMacBeth": fmb_tbl.to_dict("records"),
        "相关性": corr_tbl.to_dict("records"),
        "事件研究": ev_tbl.to_dict("records"),
        "回测绩效": perf.to_dict("records"),
        "相对基准": extra,
        "分年度": yearly_df.reset_index().to_dict("records"),
        "IC序列": {"CSP_5": ic_df["CSP_5"].to_dict(),
                 "CSP_复合": ic_df["CSP_复合"].to_dict()},
        "参数敏感性": sens_df.to_dict("records"),
        "因子权重敏感性": wsen_df.to_dict("records"),
        "成本归因": cost_df.to_dict("records"),
        "组内单调性": {k: float(v) for k, v in mono.items()},
        "回测区间": {"起始": str(pd.concat([v["ret"] for v in results.values()]).first_valid_index().date()),
                 "结束": str(pd.concat([v["ret"] for v in results.values()]).last_valid_index().date())},
        "运行耗时秒": round(time.time() - t0, 1),
    }
    with open(os.path.join(RES_DIR, "results.json"), "w", encoding="utf-8") as f:
        json.dump(jsonable(summary), f, ensure_ascii=False, indent=2)

    print(f"完成，耗时 {time.time() - t0:.1f}s")
    print(perf.to_string())


# 7. 图表
plt.rcParams.update({
    "font.sans-serif": ["Microsoft YaHei", "SimHei", "SimSun"],
    "axes.unicode_minus": False,
    "font.size": 12.5,
    "axes.titlesize": 13.5,
    "axes.labelsize": 12,
    "legend.fontsize": 10,
    "xtick.labelsize": 10,
    "ytick.labelsize": 10,
    "figure.dpi": 160,
    "savefig.dpi": 200,
    "savefig.bbox": "tight",
    "axes.grid": True,
    "grid.alpha": 0.25,
    "grid.linestyle": "--",
    "axes.edgecolor": "#555555",
    "axes.linewidth": 0.8,
})

C_BLUE = "#2C5F8D"
C_RED = "#C0392B"
C_GRAY = "#7F8C8D"
C_ORANGE = "#E67E22"
C_GREEN = "#27866B"

def _save(fig, name):
    path = os.path.join(FIG_DIR, name)
    fig.savefig(path, facecolor="white")
    plt.close(fig)
    print("saved", path)

def read_table(name, **kw):
    return pd.read_csv(os.path.join(TAB_DIR, name), **kw)

def read_series(name, **kw):
    return pd.read_csv(os.path.join(SER_DIR, name), index_col=0, parse_dates=True, **kw)

def fig_sample():
    cov = pd.read_csv(os.path.join(ROOT, "数据", "处理后数据", "样本覆盖统计.csv"),
                      parse_dates=["日期"])
    fig, axes = plt.subplots(1, 2, figsize=(12, 4.4))
    ax = axes[0]
    ax.plot(cov["日期"], cov["有效交易股票数"], color=C_BLUE, lw=1.2)
    ax.set_title("(a) 每日有效交易股票数")
    ax.set_ylabel("只")
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))

    ax = axes[1]
    ax.plot(cov["日期"], cov["因子覆盖股票数"], color=C_RED, lw=1.2)
    ax.set_title("(b) 复合 CSP 因子覆盖股票数")
    ax.set_ylabel("只")
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    fig.tight_layout()
    _save(fig, "fig01_样本覆盖.png")

def fig_pattern_dist():
    dist = read_table("t15_单K线形态分布.csv")
    fig, ax = plt.subplots(figsize=(12, 4.6))
    colors = [C_GREEN if n.startswith("阳") else C_RED for n in dist["name"]]
    ax.bar(range(len(dist)), dist["share"] * 100, color=colors, width=0.72)
    ax.set_xticks(range(len(dist)))
    ax.set_xticklabels(dist["name"], rotation=60, ha="right", fontsize=8.5)
    ax.set_ylabel("样本占比（%）")
    ax.set_title("单根日 K 线 24 种形态的样本分布（阳线为绿、阴线为红）")
    ax.yaxis.set_major_formatter(PercentFormatter(decimals=0))
    fig.tight_layout()
    _save(fig, "fig02_单K线形态分布.png")

def fig_ic():
    ic_m = read_series("ic_monthly.csv")
    fig, axes = plt.subplots(1, 2, figsize=(12.5, 4.8))
    ax = axes[0]
    x = np.arange(len(ic_m))
    ax.bar(x, ic_m["CSP_复合"] * 100, color=C_RED, alpha=0.65, width=0.8, label="复合CSP")
    ax.plot(x, ic_m["CSP_5"] * 100, color=C_BLUE, lw=1.1, label="CSP(周线)")
    ax.axhline(0, color="black", lw=0.6)
    ticks = np.linspace(0, len(ic_m) - 1, 8).astype(int)
    ax.set_xticks(ticks)
    ax.set_xticklabels([ic_m.index[i].strftime("%Y-%m") for i in ticks], rotation=45, ha="right")
    ax.set_ylabel("月度 Rank IC（%）")
    ax.set_title("(a) 月度 IC 序列")
    ax.legend(loc="upper left", frameon=False)

    ax = axes[1]
    ax.plot(ic_m.index, ic_m["CSP_复合"].cumsum() * 100, color=C_RED, lw=1.4, label="复合CSP")
    ax.plot(ic_m.index, ic_m["CSP_5"].cumsum() * 100, color=C_BLUE, lw=1.2, label="CSP(周线)")
    ax.set_title("(b) IC 累积曲线")
    ax.set_ylabel("累计 IC（%）")
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    ax.legend(loc="upper left", frameon=False)
    fig.tight_layout()
    _save(fig, "fig03_IC分析.png")

def fig_groups():
    grp = read_table("t04_分组收益.csv")
    ls_comp = read_series("ls_comp.csv").iloc[:, 0]
    ls_csp5 = read_series("ls_csp5.csv").iloc[:, 0]

    fig, axes = plt.subplots(1, 2, figsize=(12.5, 4.8))
    ax = axes[0]
    comp = grp[grp["因子"] == "CSP_复合"]
    csp5 = grp[grp["因子"] == "CSP_5"]
    x = np.arange(10)
    ax.bar(x - 0.2, csp5["年化收益"] * 100, width=0.38, color=C_BLUE, alpha=0.85, label="CSP(周线)")
    ax.bar(x + 0.2, comp["年化收益"] * 100, width=0.38, color=C_RED, alpha=0.85, label="复合CSP")
    ax.set_xticks(x)
    ax.set_xticklabels([f"G{i+1}" for i in range(10)])
    ax.axhline(0, color="black", lw=0.6)
    ax.set_ylabel("年化收益（%）")
    ax.set_title("(a) 因子十分组年化收益")
    ax.legend(frameon=False)

    ax = axes[1]
    ax.plot(ls_comp.index, (1 + ls_comp).cumprod(), color=C_RED, lw=1.4, label="复合CSP多空")
    ax.plot(ls_csp5.index, (1 + ls_csp5).cumprod(), color=C_BLUE, lw=1.1, label="CSP(周线)多空")
    ax.set_yscale("log")
    ax.set_title("(b) 多空组合累计净值（对数轴）")
    ax.set_ylabel("净值")
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    ax.legend(loc="upper left", frameon=False)
    fig.tight_layout()
    _save(fig, "fig04_分组与多空.png")

def fig_nav():
    nav = read_series("nav.csv")
    fig, axes = plt.subplots(1, 2, figsize=(12.5, 4.8))
    ax = axes[0]
    styles = {
        "沪深300": (C_GRAY, 1.2, "-"),
        "基准模型(Size+BP)": (C_BLUE, 1.5, "-"),
        "研究模型(Size+BP+复合CSP)": (C_RED, 1.5, "-"),
        "改进模型(复合CSP平滑)": (C_ORANGE, 1.5, "--"),
        "全A等权": ("#B0B0B0", 1.0, ":"),
    }
    for col, (c, lw, ls) in styles.items():
        if col in nav:
            ax.plot(nav.index, nav[col], color=c, lw=lw, ls=ls, label=col)
    ax.set_yscale("log")
    ax.set_ylabel("累计净值（对数轴）")
    ax.set_title("(a) 组合累计净值")
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    ax.legend(loc="upper left", frameon=False, fontsize=8)

    ax = axes[1]
    for col, (c, lw, ls) in styles.items():
        if col in nav and col != "沪深300":
            ax.plot(nav.index, nav[col] / nav["沪深300"], color=c, lw=lw, ls=ls, label=col)
    ax.axhline(1, color="black", lw=0.6)
    ax.set_ylabel("相对沪深300净值")
    ax.set_title("(b) 相对沪深300的超额净值")
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    ax.legend(loc="upper left", frameon=False, fontsize=8)
    fig.tight_layout()
    _save(fig, "fig05_回测净值.png")

def fig_drawdown():
    dd = read_series("drawdown.csv")
    fig, ax = plt.subplots(figsize=(12, 4.0))
    for col, c, lw in (("沪深300", C_GRAY, 1.2),
                       ("基准模型(Size+BP)", C_BLUE, 1.4),
                       ("研究模型(Size+BP+复合CSP)", C_RED, 1.4),
                       ("改进模型(复合CSP平滑)", C_ORANGE, 1.4)):
        if col in dd:
            ax.fill_between(dd.index, dd[col] * 100, 0, color=c, alpha=0.10)
            ax.plot(dd.index, dd[col] * 100, color=c, lw=lw, label=col)
    ax.set_ylabel("回撤（%）")
    ax.set_title("组合回撤对比")
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    ax.legend(loc="lower left", frameon=False, ncol=2)
    fig.tight_layout()
    _save(fig, "fig06_回撤.png")

def fig_yearly():
    yr = read_table("t11_分年度收益.csv")
    yr = yr.dropna(subset=["基准模型(Size+BP)"])
    cols = ["基准模型(Size+BP)", "研究模型(Size+BP+复合CSP)",
            "改进模型(复合CSP平滑)", "沪深300", "全A等权"]
    fig, ax = plt.subplots(figsize=(12, 4.4))
    x = np.arange(len(yr))
    w = 0.16
    colors = [C_BLUE, C_RED, C_ORANGE, C_GRAY, "#B0B0B0"]
    for i, (col, c) in enumerate(zip(cols, colors)):
        ax.bar(x + (i - 2.5) * w, yr[col] * 100, width=w * 0.92, color=c, label=col)
    ax.set_xticks(x)
    ax.set_xticklabels(yr["年份"].astype(int))
    ax.axhline(0, color="black", lw=0.6)
    ax.set_ylabel("年度收益（%）")
    ax.set_title("分年度收益对比")
    ax.legend(frameon=False, ncol=3, fontsize=8.5)
    fig.tight_layout()
    _save(fig, "fig07_分年度收益.png")

def fig_event():
    ev = read_table("t08_经典形态事件研究.csv")
    ev = ev[~ev["事件类型"].str.contains("合并")]
    fig, axes = plt.subplots(1, 2, figsize=(12.5, 4.8))
    ax = axes[0]
    x = np.arange(len(ev))
    vals = ev["超额收益均值"] * 100
    colors = [C_RED if v < 0 else C_GREEN for v in vals]
    ax.bar(x, vals, color=colors, width=0.55)
    ax.axhline(0, color="black", lw=0.6)
    for i, (v, n) in enumerate(zip(vals, ev["样本数"])):
        if v >= 0:
            ax.text(i, v + 0.10, f"{v:.2f}%\nN={n}", ha="center", va="bottom", fontsize=9)
        else:
            ax.text(i, v - 0.10, f"{v:.2f}%\nN={n}", ha="center", va="top", fontsize=9)
    ax.set_ylim(min(vals) - 1.3, max(vals) + 0.9)
    ax.set_xticks(x)
    ax.set_xticklabels(["底部\n十字星", "底部\n锤头线", "顶部\n十字星", "顶部\n倒锤头线"], fontsize=9)
    ax.set_ylabel("未来20日超额收益（%）")
    ax.set_title("(a) 经典形态事件后 20 日超额收益")

    yearly = read_table("t09_事件研究分年度.csv")
    ax = axes[1]
    for name, c in (("底部反转形态(合并)", C_GREEN), ("顶部反转形态(合并)", C_RED)):
        sub = yearly[yearly["事件类型"] == name]
        ax.plot(sub["年份"], sub["超额收益均值"] * 100, marker="o", ms=3.5, lw=1.2, color=c, label=name)
    ax.axhline(0, color="black", lw=0.6)
    ax.set_ylabel("未来20日超额收益（%）")
    ax.set_title("(b) 分年度超额收益")
    ax.legend(frameon=False, fontsize=8.5)
    fig.tight_layout()
    _save(fig, "fig08_经典形态事件.png")

def fig_corr():
    corr = read_table("t07_因子相关性.csv")
    fig, ax = plt.subplots(figsize=(9.5, 3.8))
    x = np.arange(len(corr))
    colors = [C_RED if abs(v) >= 0.1 else C_BLUE for v in corr["平均相关系数"]]
    ax.bar(x, corr["平均相关系数"] * 100, color=colors, width=0.55)
    ax.errorbar(x, corr["平均相关系数"] * 100, yerr=corr["相关系数标准差"] * 100,
                fmt="none", ecolor="#333333", capsize=3, lw=0.9)
    ax.axhline(0, color="black", lw=0.6)
    ax.set_xticks(x)
    ax.set_xticklabels(corr["对照因子"])
    ax.set_ylabel("平均截面相关系数（%）")
    ax.set_title("CSP 因子与常用量价因子的截面相关性（均值 ± 标准差）")
    fig.tight_layout()
    _save(fig, "fig09_因子相关性.png")

def fig_sensitivity():
    ws = read_table("t13_因子权重敏感性.csv")
    sens = read_table("t12_参数敏感性.csv")
    fig, axes = plt.subplots(1, 2, figsize=(12.5, 4.8))
    ax = axes[0]
    ax.plot(ws["CSP权重"] * 100, ws["年化收益"] * 100, marker="o", color=C_RED, label="年化收益")
    ax.set_xlabel("CSP 在综合得分中的权重（%）")
    ax.set_ylabel("年化收益（%）")
    ax2 = ax.twinx()
    ax2.plot(ws["CSP权重"] * 100, ws["夏普比率"], marker="s", color=C_BLUE, label="夏普比率")
    ax2.set_ylabel("夏普比率")
    ax2.grid(False)
    ax.set_title("(a) 因子权重敏感性")
    h1, l1 = ax.get_legend_handles_labels()
    h2, l2 = ax2.get_legend_handles_labels()
    ax.legend(h1 + h2, l1 + l2, loc="upper right", frameon=False, fontsize=8.5)

    ax = axes[1]
    sub = sens[(sens["组合分位"] == 0.10) & (sens["调仓周期(周)"] == 1)]
    ax.plot(sub["单边成本"] * 10000, sub["年化收益"] * 100, marker="o", color=C_GREEN,
            label="研究模型")
    ax.set_xlabel("单边交易成本（bp）")
    ax.set_ylabel("年化收益（%）")
    ax.set_title("(b) 交易成本敏感性（周度调仓）")
    ax.legend(frameon=False)
    fig.tight_layout()
    _save(fig, "fig10_敏感性.png")

def run_figures():
    fig_sample()
    fig_pattern_dist()
    fig_ic()
    fig_groups()
    fig_nav()
    fig_drawdown()
    fig_yearly()
    fig_event()
    fig_corr()
    fig_sensitivity()


if __name__ == "__main__":
    run_pipeline()
    run_figures()
