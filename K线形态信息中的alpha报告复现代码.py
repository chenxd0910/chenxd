import json
import os
import re
import sys
import time
import warnings
from dataclasses import dataclass, field
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

import numpy as np
import pandas as pd
from scipy import stats

import matplotlib
matplotlib.use("Agg")
import matplotlib.dates as mdates
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter, PercentFormatter

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

from lxml import etree

try:  # 可选：把 LaTeX 公式转换成 Word 原生公式（OMML）
    import latex2mathml.converter as _latex2mathml
    _LATEX_OK = True
except Exception:  # pragma: no cover
    _latex2mathml = None
    _LATEX_OK = False

warnings.filterwarnings("ignore")

# ---------------------------------------------------------------------------
# 全局配置：所有路径均相对本脚本所在目录
# ---------------------------------------------------------------------------
ROOT = os.path.dirname(os.path.abspath(__file__))
DATA_RAW_DIR = os.path.join(ROOT, "数据", "原始数据")
DATA_DIR = os.path.join(ROOT, "数据", "处理后数据")
RESULT_DIR = os.path.join(ROOT, "结果")
TABLE_DIR = os.path.join(RESULT_DIR, "tables")
SERIES_DIR = os.path.join(RESULT_DIR, "series")
FIG_DIR = os.path.join(RESULT_DIR, "figures")
OUT_PATH = os.path.join(ROOT, "202630196102-叶承恩-期末实验报告-K线形态信息中的alpha复刻.docx")

START_EVAL = "2021-01-01"
END_EVAL = "2025-11-28"

for _d in (DATA_DIR, TABLE_DIR, SERIES_DIR, FIG_DIR):
    os.makedirs(_d, exist_ok=True)

# Word 公式转换所需的 MML2OMML.XSL（随 Microsoft Office 安装）
_MATH_XSL = None
for _p in (
    r"C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL",
    r"C:\Program Files (x86)\Microsoft Office\root\Office16\MML2OMML.XSL",
    r"C:\Program Files\Microsoft Office\Office16\MML2OMML.XSL",
    r"C:\Program Files\Microsoft Office\root\Office15\MML2OMML.XSL",
):
    if os.path.exists(_p):
        _MATH_XSL = _p
        break

_MATH_XSLT = None
if _LATEX_OK and _MATH_XSL:
    try:
        _MATH_XSLT = etree.XSLT(etree.parse(_MATH_XSL))
    except Exception:  # pragma: no cover
        _MATH_XSLT = None



# ===========================================================================
# 第一部分  数据读取与面板构建
# ===========================================================================

# -*- coding: utf-8 -*-



# 默认原始数据目录：项目下的“数据/原始数据”，可用环境变量 CSP_DATA_DIR 覆盖
DEFAULT_DATA_DIR = os.environ.get("CSP_DATA_DIR", DATA_RAW_DIR)

# 参与面板构建的字段（全部来自原始数据文件）
PRICE_FIELDS = ("open", "high", "low", "close")
PANEL_FIELDS = PRICE_FIELDS + ("vol", "amount", "adj_factor", "total_mv", "pb", "pe_ttm", "is_st")


def _market_file(field: str, market: str, data_dir: str) -> str:
    return os.path.join(data_dir, f"{field}_{market}.csv")


def read_field(field: str, data_dir: Optional[str] = None) -> pd.DataFrame:
    """读取并合并单个字段的沪深两市数据。

    Parameters
    ----------
    field : str
        字段名，例如 ``"close"``、``"amount"``。
    data_dir : str, optional
        数据目录，默认 :data:`DEFAULT_DATA_DIR`。

    Returns
    -------
    pandas.DataFrame
        index 为 ``DatetimeIndex``（升序），columns 为股票代码，值为 float32。
        两个市场文件按列名取并集；若同名股票同时出现，优先保留沪市文件中的值。
    """
    data_dir = data_dir or DEFAULT_DATA_DIR
    frames = []
    for market in ("sh", "sz"):
        path = _market_file(field, market, data_dir)
        if not os.path.exists(path):
            continue
        df = pd.read_csv(path, index_col=0)
        df.index = pd.to_datetime(df.index, errors="coerce")
        df = df.loc[~df.index.isna()]
        df = df[~df.index.duplicated(keep="last")].sort_index()
        # 同一文件内重复列名只保留第一列
        df = df.loc[:, ~df.columns.duplicated()]
        df = df.apply(pd.to_numeric, errors="coerce").astype("float32")
        frames.append(df)

    if not frames:
        raise FileNotFoundError(f"未找到字段 {field} 的任何数据文件（{data_dir}）")

    # 以并集方式对齐（列并集、日期并集），再按优先级做空值合并
    merged = pd.concat(frames, axis=1)
    merged = merged.loc[:, ~merged.columns.duplicated()]
    merged = merged.sort_index()
    return merged


def read_benchmark(data_dir: Optional[str] = None) -> pd.DataFrame:
    """读取沪深300指数行情（``benchmark.csv``）。"""
    data_dir = data_dir or DEFAULT_DATA_DIR
    df = pd.read_csv(os.path.join(data_dir, "benchmark.csv"), index_col=0)
    df.index = pd.to_datetime(df.index, errors="coerce")
    df = df.loc[~df.index.isna()].sort_index()
    return df.apply(pd.to_numeric, errors="coerce")


def read_industry(data_dir: Optional[str] = None) -> pd.Series:
    """读取申万一级行业分类（``stk_company_info.csv``，UTF-8 编码）。

    Returns
    -------
    pandas.Series
        index 为股票代码（ts_code），值为申万一级行业名称。
    """
    data_dir = data_dir or DEFAULT_DATA_DIR
    path = os.path.join(data_dir, "stk_company_info.csv")
    info = pd.read_csv(path, encoding="utf-8-sig")
    info = info.dropna(subset=["ts_code", "l1_name"]).drop_duplicates("ts_code")
    return info.set_index("ts_code")["l1_name"].astype(str)


@dataclass
class Panel:
    """对齐后的行情面板。

    Attributes
    ----------
    dates : pandas.DatetimeIndex
        交易日序列（升序）。
    codes : pandas.Index
        股票代码。
    fields : dict[str, numpy.ndarray]
        各字段的二维数组，形状为 ``(n_dates, n_codes)``，dtype 为 float32。
    """

    dates: pd.DatetimeIndex
    codes: pd.Index
    fields: Dict[str, np.ndarray]

    @property
    def shape(self) -> tuple:
        return len(self.dates), len(self.codes)

    def __getitem__(self, key: str) -> np.ndarray:
        return self.fields[key]

    def get(self, key: str) -> Optional[np.ndarray]:
        return self.fields.get(key)

    def save(self, path: str) -> None:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        payload = {f"f_{k}": v for k, v in self.fields.items()}
        payload["dates"] = self.dates.values.astype("datetime64[ns]")
        payload["codes"] = np.asarray(self.codes, dtype=object)
        np.savez_compressed(path, **payload)

    @classmethod
    def load(cls, path: str) -> "Panel":
        with np.load(path, allow_pickle=True) as z:
            fields = {k[2:]: z[k] for k in z.files if k.startswith("f_")}
            dates = pd.DatetimeIndex(z["dates"])
            codes = pd.Index(z["codes"].astype(str))
        return cls(dates=dates, codes=codes, fields=fields)


def build_panel(
    start: str = "2020-01-01",
    end: str = "2025-12-31",
    data_dir: Optional[str] = None,
    fields: Sequence[str] = PANEL_FIELDS,
    cache_path: Optional[str] = None,
    min_obs: int = 120,
) -> Panel:
    """构建对齐面板。

    步骤：
        1. 逐字段读取沪深两市数据并合并；
        2. 以 ``close`` 的日期与股票为基准，保留其余字段的交集列；
        3. 剔除有效收盘价样本数少于 ``min_obs`` 的股票（长期停牌/退市标的）；
        4. 可选地缓存为 ``.npz``。

    Parameters
    ----------
    start, end : str
        样本区间（闭区间）。
    data_dir : str, optional
        原始数据目录。
    fields : sequence of str
        需要纳入面板的字段。
    cache_path : str, optional
        缓存文件路径；文件存在时直接读取。
    min_obs : int
        股票最少有效收盘价数量。
    """
    if cache_path and os.path.exists(cache_path):
        return Panel.load(cache_path)

    start_ts, end_ts = pd.Timestamp(start), pd.Timestamp(end)
    raw = {}
    for f in fields:
        df = read_field(f, data_dir)
        raw[f] = df

    close = raw["close"]
    close = close.loc[(close.index >= start_ts) & (close.index <= end_ts)]
    valid_close = close.notna() & (close > 0)
    keep = valid_close.sum(axis=0) >= min_obs
    close = close.loc[:, keep]
    dates = close.index
    codes = close.columns

    out = {}
    for f, df in raw.items():
        sub = df.reindex(index=dates, columns=codes)
        out[f] = sub.to_numpy(dtype="float32", copy=False)

    # 价格类字段的非正值视为缺失
    for f in PRICE_FIELDS:
        arr = out[f]
        arr[~(arr > 0)] = np.nan

    # 成交额非正视为停牌
    for f in ("amount", "vol"):
        if f in out:
            arr = out[f]
            arr[~(arr > 0)] = np.nan

    panel = Panel(dates=dates, codes=codes, fields=out)
    if cache_path:
        panel.save(cache_path)
    return panel


def weekly_last_dates(dates: pd.DatetimeIndex) -> pd.DatetimeIndex:
    """返回每个自然周最后一个交易日。"""
    s = pd.Series(dates, index=dates)
    last = s.groupby([dates.isocalendar().year, dates.isocalendar().week]).max()
    return pd.DatetimeIndex(sorted(last.values))


def month_last_dates(dates: pd.DatetimeIndex) -> pd.DatetimeIndex:
    """返回每个月最后一个交易日。"""
    s = pd.Series(dates, index=dates)
    last = s.groupby([dates.year, dates.month]).max()
    return pd.DatetimeIndex(sorted(last.values))

# ===========================================================================
# 第二部分  因子预处理与绩效指标
# ===========================================================================

# -*- coding: utf-8 -*-




# ---------------------------------------------------------------------------
# 横截面处理
# ---------------------------------------------------------------------------
def mad_winsorize(x: np.ndarray, n: float = 5.0) -> np.ndarray:
    """MAD 去极值：将偏离中位数超过 ``n`` 倍 MAD 的值截断。"""
    x = np.asarray(x, dtype=float)
    med = np.nanmedian(x)
    mad = np.nanmedian(np.abs(x - med))
    if not np.isfinite(mad) or mad == 0:
        return x
    scale = 1.4826 * mad
    return np.clip(x, med - n * scale, med + n * scale)


def quantile_winsorize(x: np.ndarray, lower: float = 0.01, upper: float = 0.99) -> np.ndarray:
    """分位数去极值。"""
    x = np.asarray(x, dtype=float)
    lo, hi = np.nanquantile(x, [lower, upper])
    return np.clip(x, lo, hi)


def zscore(x: np.ndarray) -> np.ndarray:
    """横截面标准化（忽略缺失值）。"""
    x = np.asarray(x, dtype=float)
    mu = np.nanmean(x)
    sd = np.nanstd(x)
    if not np.isfinite(sd) or sd == 0:
        return x - mu
    return (x - mu) / sd


def industry_neutralize(x: np.ndarray, group: np.ndarray) -> np.ndarray:
    """行业中性化：减去行业内均值（``group`` 为整数行业编码，-1 表示缺失）。"""
    x = np.asarray(x, dtype=float)
    group = np.asarray(group)
    out = x.copy()
    valid = np.isfinite(x) & (group >= 0)
    if not valid.any():
        return out
    g = group[valid]
    v = x[valid]
    n_groups = int(g.max()) + 1
    cnt = np.bincount(g, minlength=n_groups).astype(float)
    s = np.bincount(g, weights=v, minlength=n_groups)
    mean = np.divide(s, cnt, out=np.zeros(n_groups), where=cnt > 0)
    out[valid] = v - mean[g]
    return out


def rank_corr(a: np.ndarray, b: np.ndarray, min_n: int = 30) -> float:
    """Spearman 秩相关系数（忽略缺失值）。"""
    mask = np.isfinite(a) & np.isfinite(b)
    if mask.sum() < min_n:
        return np.nan
    ra = stats.rankdata(a[mask])
    rb = stats.rankdata(b[mask])
    if ra.std() == 0 or rb.std() == 0:
        return np.nan
    return float(np.corrcoef(ra, rb)[0, 1])


def cross_sectional_rank(x: np.ndarray) -> np.ndarray:
    """横截面百分位排名（0~1，缺失保持缺失）。"""
    x = np.asarray(x, dtype=float)
    out = np.full_like(x, np.nan)
    mask = np.isfinite(x)
    if mask.sum() == 0:
        return out
    out[mask] = stats.rankdata(x[mask]) / mask.sum()
    return out


# ---------------------------------------------------------------------------
# 时间序列与绩效
# ---------------------------------------------------------------------------
def newey_west_tstat(series: pd.Series, lags: int = 4) -> float:
    """Newey-West 调整后的 t 统计量（原假设：均值为 0）。"""
    x = np.asarray(series, dtype=float)
    x = x[np.isfinite(x)]
    n = len(x)
    if n < 10:
        return np.nan
    mu = x.mean()
    e = x - mu
    gamma0 = np.dot(e, e) / n
    var = gamma0
    for lag in range(1, min(lags, n - 1) + 1):
        w = 1.0 - lag / (lags + 1.0)
        gamma = np.dot(e[lag:], e[:-lag]) / n
        var += 2.0 * w * gamma
    if var <= 0:
        return np.nan
    return float(mu / np.sqrt(var / n))


def annualized_return(weekly_ret: pd.Series, periods_per_year: int = 52) -> float:
    r = weekly_ret.dropna()
    if len(r) == 0:
        return np.nan
    nav = (1 + r).prod()
    return float(nav ** (periods_per_year / len(r)) - 1)


def annualized_vol(weekly_ret: pd.Series, periods_per_year: int = 52) -> float:
    r = weekly_ret.dropna()
    if len(r) < 2:
        return np.nan
    return float(r.std(ddof=1) * np.sqrt(periods_per_year))


def sharpe_ratio(weekly_ret: pd.Series, rf: float = 0.02, periods_per_year: int = 52) -> float:
    ann_r = annualized_return(weekly_ret, periods_per_year)
    ann_v = annualized_vol(weekly_ret, periods_per_year)
    if not np.isfinite(ann_v) or ann_v == 0:
        return np.nan
    return float((ann_r - rf) / ann_v)


def max_drawdown(weekly_ret: pd.Series) -> float:
    r = weekly_ret.dropna()
    if len(r) == 0:
        return np.nan
    nav = (1 + r).cumprod()
    peak = nav.cummax()
    dd = nav / peak - 1
    return float(dd.min())


def calmar_ratio(weekly_ret: pd.Series, periods_per_year: int = 52) -> float:
    mdd = max_drawdown(weekly_ret)
    ann = annualized_return(weekly_ret, periods_per_year)
    if not np.isfinite(mdd) or mdd == 0:
        return np.nan
    return float(ann / abs(mdd))


def perf_summary(weekly_ret: pd.Series, name: str, rf: float = 0.02) -> Dict[str, float]:
    """回测绩效汇总。"""
    r = weekly_ret.dropna()
    nav = (1 + r).cumprod()
    return {
        "名称": name,
        "年化收益率": annualized_return(r),
        "年化波动率": annualized_vol(r),
        "夏普比率": sharpe_ratio(r, rf),
        "最大回撤": max_drawdown(r),
        "卡玛比率": calmar_ratio(r),
        "周胜率": float((r > 0).mean()),
        "累计净值": float(nav.iloc[-1]) if len(nav) else np.nan,
        "样本周数": int(len(r)),
    }


def nav_from_returns(weekly_ret: pd.Series, start: pd.Timestamp) -> pd.Series:
    r = weekly_ret.dropna()
    nav = (1 + r).cumprod()
    nav = pd.concat([pd.Series([1.0], index=[start]), nav])
    return nav


def yearly_returns(weekly_ret: pd.Series) -> pd.Series:
    r = weekly_ret.dropna()
    return r.groupby(r.index.year).apply(lambda x: (1 + x).prod() - 1)


def max_drawdown_series(weekly_ret: pd.Series) -> pd.Series:
    r = weekly_ret.dropna()
    nav = (1 + r).cumprod()
    return nav / nav.cummax() - 1


# ---------------------------------------------------------------------------
# 表格输出
# ---------------------------------------------------------------------------
def save_table(df: pd.DataFrame, path: str, index: bool = False) -> None:
    """按 UTF-8-SIG 保存 CSV（Excel 直接打开不乱码）。"""
    df.to_csv(path, index=index, encoding="utf-8-sig", float_format="%.6f")


def fmt_pct(x: float, digits: int = 2) -> str:
    if x is None or not np.isfinite(x):
        return "—"
    return f"{x * 100:.{digits}f}%"


def fmt_num(x: float, digits: int = 3) -> str:
    if x is None or not np.isfinite(x):
        return "—"
    return f"{x:.{digits}f}"

# ===========================================================================
# 第三部分  K 线形态划分与 CSP 因子
# ===========================================================================

# -*- coding: utf-8 -*-




N_BASE = 24          # 单根 K 线的形态数
N_PATTERN = 24 ** 3 * 3   # 组合形态数：41,472

# 形态划分阈值（与研报一致，均为相对前收盘价的百分比）
BODY_BINS = (0.015, 0.03)
SHADOW_LONG = 0.02
VOL_UP, VOL_DOWN = 1.2, 0.8


# ---------------------------------------------------------------------------
# 基础工具
# ---------------------------------------------------------------------------
def adjusted_prices(panel: Panel) -> Dict[str, np.ndarray]:
    """计算后复权价格：``price * adj_factor``。"""
    adj = panel["adj_factor"]
    out = {}
    for f in ("open", "high", "low", "close"):
        with np.errstate(invalid="ignore"):
            out[f] = panel[f] * adj
    return out


def forward_return(adj_close: np.ndarray, horizon: int = 20) -> np.ndarray:
    """未来 ``horizon`` 个交易日的收益率矩阵（基于后复权收盘价）。"""
    t, n = adj_close.shape
    out = np.full((t, n), np.nan, dtype="float32")
    if t > horizon:
        with np.errstate(invalid="ignore", divide="ignore"):
            out[:-horizon] = adj_close[horizon:] / adj_close[:-horizon] - 1.0
    return out


def excess_over_market(ret: np.ndarray) -> np.ndarray:
    """相对全市场等权平均收益的超额收益（逐日横截面去均值）。"""
    finite = np.isfinite(ret)
    cnt = finite.sum(axis=1, keepdims=True)
    s = np.where(finite, ret, 0.0).sum(axis=1, keepdims=True)
    with np.errstate(invalid="ignore", divide="ignore"):
        mean = np.divide(s, cnt, out=np.full_like(s, np.nan, dtype="float32"), where=cnt > 0)
    return ret - mean


def rolling_mean_prev(arr: np.ndarray, window: int = 20, min_periods: int = 10) -> np.ndarray:
    """逐列滚动均值，并向后移一期（仅使用历史信息）。"""
    df = pd.DataFrame(arr)
    return df.rolling(window, min_periods=min_periods).mean().shift(1).to_numpy()


def rolling_max(arr: np.ndarray, window: int, min_periods: int) -> np.ndarray:
    return pd.DataFrame(arr).rolling(window, min_periods=min_periods).max().to_numpy()


def rolling_min(arr: np.ndarray, window: int, min_periods: int) -> np.ndarray:
    return pd.DataFrame(arr).rolling(window, min_periods=min_periods).min().to_numpy()


def _forward_fill_rows(arr: np.ndarray) -> np.ndarray:
    """按行前向填充（用于把“最近一根已完成 K 线”的因子值延续到每个交易日）。"""
    return pd.DataFrame(arr).ffill().to_numpy(dtype="float32")


# ---------------------------------------------------------------------------
# k 日 K 线聚合
# ---------------------------------------------------------------------------
@dataclass
class Bars:
    """非重叠的 k 日 K 线序列。"""

    k: int
    ends: np.ndarray            # 每根 K 线结束日在日期轴上的位置
    open: np.ndarray
    high: np.ndarray
    low: np.ndarray
    close: np.ndarray
    prev_close: np.ndarray
    amount: np.ndarray
    base_pattern: np.ndarray    # 0~23，-1 表示无效
    vol_class: np.ndarray       # 0 缩量 / 1 常量 / 2 放量
    body: np.ndarray
    upper: np.ndarray
    lower: np.ndarray

    @property
    def n_bars(self) -> int:
        return len(self.ends)


def build_bars(panel: Panel, k: int) -> Bars:
    """把日线聚合为 k 日 K 线（非重叠、按交易日历对齐）。

    K 线 ``j`` 覆盖日线 ``[j*k, j*k+k-1]``，其结束日 ``ends[j] = j*k+k-1``。
    实体、影线均以“前一根 K 线的收盘价”为基准，量纲为百分比。
    """
    t, n = panel.shape
    O, H, L, C = (panel[f] for f in ("open", "high", "low", "close"))
    A = panel["amount"]

    ends = np.arange(k - 1, t, k, dtype=np.int64)
    starts = ends - k + 1
    m = len(ends)

    op = O[starts].astype("float32")
    cp = C[ends].astype("float32")
    prev = np.full((m, n), np.nan, dtype="float32")
    has_prev = (ends - k) >= 1
    prev[has_prev] = C[ends[has_prev] - k]

    hi = np.full((m, n), -np.inf, dtype="float32")
    lo = np.full((m, n), np.inf, dtype="float32")
    amt = np.zeros((m, n), dtype="float32")
    cnt = np.zeros((m, n), dtype="int16")
    for off in range(k):
        idx = starts + off
        hi = np.fmax(hi, H[idx])
        lo = np.fmin(lo, L[idx])
        a = A[idx]
        ok = np.isfinite(a) & (a > 0)
        amt += np.where(ok, a, 0.0).astype("float32")
        cnt += ok.astype("int16")
    hi = np.where(np.isfinite(hi), hi, np.nan)
    lo = np.where(np.isfinite(lo), lo, np.nan)
    amt[cnt < k] = np.nan

    valid = (
        np.isfinite(op) & np.isfinite(hi) & np.isfinite(lo) & np.isfinite(cp)
        & np.isfinite(prev) & (prev > 0) & np.isfinite(amt) & (hi >= lo)
    )

    with np.errstate(invalid="ignore", divide="ignore"):
        body = (cp - op) / prev
        upper = (hi - np.fmax(op, cp)) / prev
        lower = (np.fmin(op, cp) - lo) / prev
    upper = np.where(np.isfinite(upper), np.maximum(upper, 0.0), np.nan)
    lower = np.where(np.isfinite(lower), np.maximum(lower, 0.0), np.nan)
    body = np.where(valid, body, np.nan)

    # 单根 K 线 24 种形态：2（阳/阴）× 3（实体）× 2（上影）× 2（下影）
    body_cls = np.digitize(np.abs(body), BODY_BINS).astype("int16")
    bull = (cp >= op)
    upper_long = (upper >= SHADOW_LONG)
    lower_long = (lower >= SHADOW_LONG)
    base = (bull.astype("int16") * 12 + body_cls * 4
            + upper_long.astype("int16") * 2 + lower_long.astype("int16"))
    base = np.where(valid, base, -1).astype("int16")

    # 成交量状态：当日成交额 / 过去 20 根同类 K 线成交额均值
    ma = rolling_mean_prev(amt, window=20, min_periods=10)
    with np.errstate(invalid="ignore", divide="ignore"):
        ratio = amt / ma
    vol_class = np.where(ratio >= VOL_UP, 2, np.where(ratio <= VOL_DOWN, 0, 1)).astype("int8")
    vol_class = np.where(np.isfinite(ratio), vol_class, 1).astype("int8")

    return Bars(k=k, ends=ends, open=op, high=hi, low=lo, close=cp,
                prev_close=prev, amount=amt, base_pattern=base,
                vol_class=vol_class, body=body, upper=upper, lower=lower)


def pattern_codes(bars: Bars) -> Tuple[np.ndarray, np.ndarray]:
    """由连续 3 根 K 线 + 最新一根的成交量状态生成组合形态编码。

    Returns
    -------
    codes : ndarray
        形状 ``(n_bars-2, n_stocks)``，取值 0~41471，-1 表示无效。
    end_idx : ndarray
        每个组合形态对应的 K 线结束日在日期轴上的位置。
    """
    base = bars.base_pattern
    m = base.shape[0]
    if m < 3:
        return np.empty((0, base.shape[1]), dtype="int32"), np.empty(0, dtype=np.int64)
    b1, b2, b3 = base[:-2], base[1:-1], base[2:]
    vc = bars.vol_class[2:]
    valid = (b1 >= 0) & (b2 >= 0) & (b3 >= 0)
    codes = ((b1.astype("int32") * N_BASE + b2.astype("int32")) * N_BASE + b3.astype("int32")) * 3 + vc
    codes = np.where(valid, codes, -1).astype("int32")
    return codes, bars.ends[2:].copy()


# ---------------------------------------------------------------------------
# 形态历史绩效估计（扩张窗口、按月重估）
# ---------------------------------------------------------------------------
def score_cutoffs(dates: pd.DatetimeIndex, horizon: int = 20,
                  min_train_days: int = 120) -> Tuple[np.ndarray, np.ndarray]:
    """构造按月重估的“可用样本截止位置”。

    对每个月末 ``d``，只有结束日不晚于 ``d`` 前 ``horizon`` 个交易日的样本
    才拥有已实现的未来收益，因此以这些样本估计形态绩效，避免使用未来信息。

    Returns
    -------
    refit_dates_idx : ndarray
        重估日（月末交易日）在日期轴上的位置。
    cutoffs : ndarray
        每个重估日可用的最大样本位置（bar end 位置）。
    """
    ends = month_last_dates(dates)
    idx = dates.get_indexer(ends)
    idx = idx[idx >= min_train_days]
    cutoffs = idx - horizon
    keep = cutoffs > 0
    return idx[keep], cutoffs[keep]


def fit_pattern_scores(codes: np.ndarray, end_idx: np.ndarray, excess: np.ndarray,
                       cutoffs: np.ndarray, min_n: int = 20,
                       kappa: float = 50.0, clip: float = 3.0) -> np.ndarray:
    """估计每个重估日、每种组合形态的历史绩效得分。

    形态得分定义为 ``IR = mean(excess) / std(excess)``，并施加小样本收缩
    ``n / (n + kappa)``，避免样本很少的形态给出极端得分。

    Returns
    -------
    scores : ndarray
        形状 ``(n_refit, 41472)``；样本数不足 ``min_n`` 的形态取 NaN。
    """
    n_refit = len(cutoffs)
    scores = np.full((n_refit, N_PATTERN), np.nan, dtype="float32")
    if n_refit == 0:
        return scores

    valid = (codes >= 0) & np.isfinite(excess)
    c = codes[valid]
    e = excess[valid].astype("float64")
    b = np.broadcast_to(end_idx[:, None], codes.shape)[valid]
    known = np.searchsorted(cutoffs, b, side="left")
    keep = known < n_refit
    c, e, known = c[keep], e[keep], known[keep]
    if len(c) == 0:
        return scores

    flat = known.astype(np.int64) * N_PATTERN + c
    size = n_refit * N_PATTERN
    cnt = np.bincount(flat, minlength=size).reshape(n_refit, N_PATTERN)
    s1 = np.bincount(flat, weights=e, minlength=size).reshape(n_refit, N_PATTERN)
    s2 = np.bincount(flat, weights=e * e, minlength=size).reshape(n_refit, N_PATTERN)
    cnt = np.cumsum(cnt, axis=0)
    s1 = np.cumsum(s1, axis=0)
    s2 = np.cumsum(s2, axis=0)

    mean = np.divide(s1, cnt, out=np.zeros_like(s1), where=cnt > 0)
    var = np.divide(s2, cnt, out=np.zeros_like(s2), where=cnt > 0) - mean ** 2
    std = np.sqrt(np.maximum(var, 1e-8))
    ir = mean / std
    shrink = cnt / (cnt + kappa)
    sc = np.clip(ir * shrink, -clip, clip)
    sc[cnt < min_n] = np.nan
    scores[:] = sc.astype("float32")
    return scores


def map_scores_to_daily(codes: np.ndarray, end_idx: np.ndarray, scores: np.ndarray,
                        refit_dates_idx: np.ndarray, dates: pd.DatetimeIndex,
                        codes_len: int) -> np.ndarray:
    """把逐形态得分映射为逐日因子值（last completed bar 延续）。"""
    t = len(dates)
    out = np.full((t, codes_len), np.nan, dtype="float32")
    if len(refit_dates_idx) == 0:
        return out
    r = np.searchsorted(refit_dates_idx, end_idx, side="right") - 1
    r_grid = np.broadcast_to(r[:, None], codes.shape)
    ok = (r_grid >= 0) & (codes >= 0)
    if not ok.any():
        return out
    val = np.full(codes.shape, np.nan, dtype="float32")
    val[ok] = scores[r_grid[ok], codes[ok]]
    out[end_idx] = val
    return _forward_fill_rows(out)


# ---------------------------------------------------------------------------
# 单周期与多周期因子
# ---------------------------------------------------------------------------
@dataclass
class CSPResult:
    factor: np.ndarray
    scores: np.ndarray
    refit_dates: pd.DatetimeIndex
    coverage: float
    n_patterns: int
    n_scored: int
    diagnostics: Dict[str, float] = field(default_factory=dict)


def build_csp(panel: Panel, k: int, excess20: np.ndarray, cutoffs: np.ndarray,
              refit_dates_idx: np.ndarray, min_n: int = 20,
              kappa: float = 50.0) -> CSPResult:
    """构建单周期 CSP_k 因子。"""
    bars = build_bars(panel, k)
    codes, end_idx = pattern_codes(bars)
    # 形态得分使用各 K 线结束日的 20 日超额收益
    excess_at_bar = excess20[end_idx] if len(end_idx) else np.empty((0, panel.shape[1]))
    scores = fit_pattern_scores(codes, end_idx, excess_at_bar, cutoffs,
                                min_n=min_n, kappa=kappa)
    factor = map_scores_to_daily(codes, end_idx, scores, refit_dates_idx,
                                 panel.dates, panel.shape[1])
    valid_f = np.isfinite(factor)
    n_scored = int(np.isfinite(scores[-1]).sum())
    n_patterns = int(len(np.unique(codes[codes >= 0])))
    return CSPResult(
        factor=factor, scores=scores,
        refit_dates=panel.dates[refit_dates_idx],
        coverage=float(valid_f.mean()),
        n_patterns=n_patterns, n_scored=n_scored,
        diagnostics={
            "k": k,
            "n_bars": int(bars.n_bars),
            "n_patterns_observed": n_patterns,
            "n_patterns_scored": n_scored,
            "coverage_all": float(valid_f.mean()),
            "mean": float(np.nanmean(factor)),
            "std": float(np.nanstd(factor)),
        },
    )


def composite_factor(factors: Sequence[np.ndarray], min_valid: int = 2,
                     clip: float = 3.0) -> np.ndarray:
    """多周期复合因子：各周期因子横截面标准化后等权平均。"""
    stack = []
    for f in factors:
        mu = np.nanmean(f, axis=1)
        sd = np.nanstd(f, axis=1)
        with np.errstate(invalid="ignore", divide="ignore"):
            z = (f - mu[:, None]) / sd[:, None]
        stack.append(np.clip(z, -clip, clip).astype("float32"))
    arr = np.stack(stack, axis=0)
    cnt = np.isfinite(arr).sum(axis=0)
    with np.errstate(invalid="ignore"):
        out = np.nanmean(arr, axis=0)
    out = np.where(cnt >= min_valid, out, np.nan)
    return out.astype("float32")


def cross_sectional_z(f: np.ndarray) -> np.ndarray:
    """逐日横截面标准化（用于因子入库与相关性分析）。"""
    mu = np.nanmean(f, axis=1, keepdims=True)
    sd = np.nanstd(f, axis=1, keepdims=True)
    with np.errstate(invalid="ignore", divide="ignore"):
        z = (f - mu) / sd
    return z.astype("float32")


# ---------------------------------------------------------------------------
# 经典单 K 线形态事件（十字星 / 锤头线 / 倒锤头线）
# ---------------------------------------------------------------------------
@dataclass
class ClassicEvents:
    doji: np.ndarray
    hammer: np.ndarray
    inv_hammer: np.ndarray
    top: np.ndarray
    bottom: np.ndarray
    vol_up: np.ndarray
    bullish: np.ndarray     # 底部 + 十字星/锤头线 + 放量
    bearish: np.ndarray     # 顶部 + 十字星/倒锤头线 + 放量
    score: np.ndarray       # 事件打分因子（过去 20 日净事件数）


def build_classic_events(panel: Panel, window: int = 250, min_periods: int = 120) -> ClassicEvents:
    """识别研报中的三类经典 K 线形态及其位置、成交量条件。"""
    bars = build_bars(panel, 1)
    body = np.abs(bars.body)
    upper, lower = bars.upper, bars.lower
    doji = (body <= 0.02) & (upper >= 0.03) & (lower >= 0.03)
    hammer = (body <= 0.02) & (upper <= 0.02) & (lower >= 0.05)
    inv_hammer = (body <= 0.02) & (upper >= 0.05) & (lower <= 0.02)

    hi = rolling_max(bars.high, window, min_periods)
    lo = rolling_min(bars.low, window, min_periods)
    close = bars.close
    top = close >= 0.9 * hi
    bottom = close <= 1.1 * lo

    ma = rolling_mean_prev(bars.amount, 20, 10)
    vol_up = bars.amount >= 1.2 * ma

    bullish = bottom & (doji | hammer) & vol_up
    bearish = top & (doji | inv_hammer) & vol_up

    ind = np.zeros_like(close, dtype="float32")
    ind += np.where(bullish, 1.0, 0.0)
    ind -= np.where(bearish, 1.0, 0.0)
    score = pd.DataFrame(ind).rolling(20, min_periods=1).sum().to_numpy(dtype="float32")
    # 事件只在 K 线结束日有效，向前填充到之后的交易日
    score = _forward_fill_rows(score)
    return ClassicEvents(doji=doji, hammer=hammer, inv_hammer=inv_hammer,
                         top=top, bottom=bottom, vol_up=vol_up,
                         bullish=bullish, bearish=bearish, score=score)


def pattern_distribution(panel: Panel) -> pd.DataFrame:
    """统计单根日 K 线 24 种形态的样本占比，用于报告图表。"""
    bars = build_bars(panel, 1)
    base = bars.base_pattern.ravel()
    base = base[base >= 0]
    cnt = np.bincount(base, minlength=N_BASE).astype(float)
    share = cnt / cnt.sum()
    bull = np.arange(N_BASE) // 12
    body_cls = (np.arange(N_BASE) % 12) // 4
    upper_long = (np.arange(N_BASE) % 4) // 2
    lower_long = np.arange(N_BASE) % 2
    names = []
    for i in range(N_BASE):
        names.append(
            f"{'阳' if bull[i] else '阴'}-{['小', '中', '大'][body_cls[i]]}"
            f"-{'长上' if upper_long[i] else '短上'}{'长下' if lower_long[i] else '短下'}"
        )
    return pd.DataFrame({"code": np.arange(N_BASE), "name": names,
                         "count": cnt.astype(int), "share": share})

# ===========================================================================
# 第四部分  周度多因子回测引擎
# ===========================================================================

# -*- coding: utf-8 -*-





# ---------------------------------------------------------------------------
# 周度数据准备
# ---------------------------------------------------------------------------
@dataclass
class WeeklyData:
    dates: pd.DatetimeIndex          # 调仓日
    idx: np.ndarray                  # 调仓日在日频日期轴上的位置
    next_ret: np.ndarray             # (T, N) 下一周收益率
    bm_ret: pd.Series                # 沪深300 周收益
    all_ret: pd.Series               # 全A等权周收益
    eligible: np.ndarray             # (T, N) 可交易标记
    close: np.ndarray
    fields: Dict[str, np.ndarray] = field(default_factory=dict)


def build_weekly_data(panel: Panel, benchmark_close: pd.Series,
                      min_listed: int = 250, st_threshold: float = 0.5) -> WeeklyData:
    """构造周度回测所需的收益、可交易标记等。

    可交易条件（在每个调仓日判定）：
        1. 当日有有效收盘价与成交额（非停牌）；
        2. 非 ST/*ST；
        3. 上市（有收盘价）满 ``min_listed`` 个交易日。
    """
    dates = panel.dates
    close = panel["close"]
    amount = panel["amount"]
    adj = panel["adj_factor"]
    st = panel["is_st"]

    valid = np.isfinite(close) & (close > 0)
    listed = np.cumsum(valid, axis=0)
    is_st = np.isfinite(st) & (st > st_threshold)

    wd_all = weekly_last_dates(dates)
    idx_all = dates.get_indexer(wd_all)
    keep = (idx_all >= 0) & (idx_all + 5 < len(dates))
    wd = wd_all[keep]
    idx = idx_all[keep]

    adj_close = close * adj
    with np.errstate(invalid="ignore", divide="ignore"):
        next_ret = adj_close[idx + 5] / adj_close[idx] - 1.0
    next_ret = next_ret.astype("float32")

    eligible = (
        valid[idx] & np.isfinite(amount[idx]) & (~is_st[idx]) & (listed[idx] >= min_listed)
    )

    # 沪深300：按同一交易日历对齐
    bm = benchmark_close.reindex(dates).ffill()
    bm_ret = pd.Series(bm.iloc[idx + 5].to_numpy() / bm.iloc[idx].to_numpy() - 1.0, index=wd)

    # 全A等权（所有有有效收益的股票）
    with np.errstate(invalid="ignore"):
        all_ret = pd.Series(np.nanmean(next_ret, axis=1), index=wd)

    return WeeklyData(dates=wd, idx=idx, next_ret=next_ret, bm_ret=bm_ret,
                      all_ret=all_ret, eligible=eligible,
                      close=panel["close"][idx].astype("float32"))


# ---------------------------------------------------------------------------
# 横截面因子处理
# ---------------------------------------------------------------------------
def prep_cross_section(x: np.ndarray, group: Optional[np.ndarray] = None,
                       winsor: Tuple[float, float] = (0.01, 0.99)) -> np.ndarray:
    """去极值 + 标准化 +（可选）行业中性化，返回截面 z 值。"""
    x = np.asarray(x, dtype=float)
    out = np.full_like(x, np.nan)
    mask = np.isfinite(x)
    if mask.sum() < 10:
        return out
    lo, hi = np.nanquantile(x[mask], winsor)
    v = np.clip(x, lo, hi)
    mu, sd = np.nanmean(v[mask]), np.nanstd(v[mask])
    if not np.isfinite(sd) or sd == 0:
        out[mask] = 0.0
        return out
    z = (v - mu) / sd
    if group is not None:
        z = industry_neutralize(z, group)
        sd2 = np.nanstd(z[mask])
        if np.isfinite(sd2) and sd2 > 0:
            z = (z - np.nanmean(z[mask])) / sd2
    out[mask] = z[mask]
    return out


def industry_codes(codes: pd.Index, industry: pd.Series) -> np.ndarray:
    """把股票代码映射为整数行业编码，缺失记为 -1。"""
    names = industry.reindex(codes).fillna("未知").astype(str)
    uniq = {n: i for i, n in enumerate(sorted(names.unique()))}
    return names.map(uniq).to_numpy(dtype=np.int32)


def select_top(score: np.ndarray, mask: np.ndarray, top_frac: float,
               min_names: int = 30, prev: Optional[np.ndarray] = None,
               buffer: float = 0.0) -> np.ndarray:
    """按得分从高到低选取组合，返回持仓索引。

    ``buffer`` 为持仓缓冲比例：已有持仓只要仍位于得分前
    ``top_frac + buffer`` 分位内就继续持有，从而降低换手率。
    """
    cand = np.flatnonzero(mask & np.isfinite(score))
    if len(cand) == 0:
        return np.empty(0, dtype=np.int64)
    k = int(round(top_frac * len(cand)))
    k = max(min_names, min(k, len(cand)))
    s = score[cand]
    if buffer > 0 and prev is not None and len(prev) > 0:
        thr = np.nanquantile(score[cand], 1.0 - min(top_frac + buffer, 0.99))
        keep = prev[np.isfinite(score[prev]) & (score[prev] >= thr)]
        if len(keep) >= k:
            order = np.argsort(-score[keep])
            return keep[order[:k]]
        rest = cand[~np.isin(cand, keep)]
        need = k - len(keep)
        if need <= 0 or len(rest) == 0:
            return keep
        pick = rest[np.argpartition(-score[rest], min(need, len(rest)) - 1)[:need]]
        return np.concatenate([keep, pick])
    pick = np.argpartition(-s, k - 1)[:k]
    return cand[pick]


# ---------------------------------------------------------------------------
# 回测
# ---------------------------------------------------------------------------
@dataclass
class BacktestResult:
    name: str
    ret: pd.Series                   # 周度净收益
    gross: pd.Series
    turnover: pd.Series
    n_holdings: pd.Series
    holdings: Dict[pd.Timestamp, np.ndarray] = field(default_factory=dict)

    @property
    def nav(self) -> pd.Series:
        r = self.ret.dropna()
        nav = (1 + r).cumprod()
        return pd.concat([pd.Series([1.0], index=[r.index[0]]), nav])

    def summary(self, rf: float = 0.02) -> Dict[str, float]:
        return perf_summary(self.ret, self.name, rf=rf)


def run_strategy(wd: WeeklyData, score: np.ndarray, name: str,
                 top_frac: float = 0.1, fee: float = 0.0015,
                 rebalance_weeks: int = 1, min_names: int = 30,
                 buffer: float = 0.0) -> BacktestResult:
    """按给定综合得分执行周度调仓回测。"""
    n_weeks = len(wd.idx)
    rets = np.full(n_weeks, np.nan)
    gross = np.full(n_weeks, np.nan)
    turnover = np.full(n_weeks, np.nan)
    n_hold = np.full(n_weeks, np.nan, dtype=float)
    holdings_map: Dict[pd.Timestamp, np.ndarray] = {}

    holdings = np.empty(0, dtype=np.int64)
    prev = np.empty(0, dtype=np.int64)
    for step in range(n_weeks):
        i = wd.idx[step]
        need_rebalance = (step % rebalance_weeks == 0) or (len(holdings) == 0)
        cost = 0.0
        if need_rebalance:
            pick = select_top(score[step], wd.eligible[step], top_frac, min_names,
                              prev=holdings, buffer=buffer)
            # 选中的股票必须拥有下一周收益
            pick = pick[np.isfinite(wd.next_ret[step, pick])]
            if len(pick) == 0:
                if len(holdings) == 0:
                    continue
                pick = holdings
            new = pick
            if len(prev) == 0:
                to = 1.0
            else:
                overlap = np.intersect1d(prev, new, assume_unique=False).size
                to = 1.0 - overlap / max(len(new), 1)
            turnover[step] = to
            cost = 2.0 * fee * to
            prev = new
            holdings = new
            holdings_map[wd.dates[step]] = new
        n_hold[step] = len(holdings)
        if len(holdings) == 0:
            continue
        r = wd.next_ret[step, holdings]
        r = r[np.isfinite(r)]
        if len(r) == 0:
            continue
        gross[step] = float(np.mean(r))
        rets[step] = gross[step] - cost

    return BacktestResult(
        name=name,
        ret=pd.Series(rets, index=wd.dates),
        gross=pd.Series(gross, index=wd.dates),
        turnover=pd.Series(turnover, index=wd.dates),
        n_holdings=pd.Series(n_hold, index=wd.dates),
        holdings=holdings_map,
    )


def strategy_table(results: Sequence[BacktestResult], benchmark: pd.Series,
                   rf: float = 0.02) -> pd.DataFrame:
    """汇总多个策略与基准的绩效指标。"""
    rows = []
    for res in results:
        s = res.summary(rf=rf)
        r = res.ret.dropna()
        ex = (r - benchmark.reindex(r.index)).dropna()
        s["年化超额(vs沪深300)"] = float(ex.mean() * 52)
        te = float(ex.std(ddof=1) * np.sqrt(52)) if len(ex) > 2 else np.nan
        s["跟踪误差(vs沪深300)"] = te
        s["信息比率(vs沪深300)"] = float(s["年化超额(vs沪深300)"] / te) if te and np.isfinite(te) else np.nan
        s["平均换手率"] = float(res.turnover.dropna().mean())
        s["平均持仓数"] = float(res.n_holdings.dropna().mean())
        rows.append(s)
    bm = perf_summary(benchmark.dropna(), "沪深300", rf=rf)
    rows.append(bm)
    return pd.DataFrame(rows)


def excess_table(results: Sequence[BacktestResult], benchmark: pd.Series) -> pd.DataFrame:
    """各策略相对沪深300与基准模型的周度/年化超额。"""
    rows = []
    for res in results:
        r = res.ret.dropna()
        ex_bm = (r - benchmark.reindex(r.index)).dropna()
        rows.append({
            "策略": res.name,
            "周均超额(vs沪深300)": float(ex_bm.mean()),
            "年化超额(vs沪深300)": float(ex_bm.mean() * 52),
            "超额胜率(vs沪深300)": float((ex_bm > 0).mean()),
            "累计超额(vs沪深300)": float((1 + ex_bm).prod() - 1),
        })
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# 分组回测
# ---------------------------------------------------------------------------
def decile_backtest(wd: WeeklyData, factor: np.ndarray, n_groups: int = 10,
                    min_n: int = 100) -> Tuple[pd.DataFrame, pd.Series]:
    """因子分组（周度调仓、等权）收益。

    Returns
    -------
    group_ret : DataFrame
        各分组周度收益，列名为 ``G1`` ... ``G10``（G1 为因子值最低组）。
    ls : Series
        多空组合（G10 - G1）周度收益。
    """
    n_weeks = len(wd.idx)
    out = np.full((n_weeks, n_groups), np.nan)
    for step in range(n_weeks):
        i = wd.idx[step]
        f = factor[i]
        r = wd.next_ret[step]
        mask = wd.eligible[step] & np.isfinite(f) & np.isfinite(r)
        if mask.sum() < min_n:
            continue
        vals = f[mask]
        rank = pd.Series(vals).rank(method="first").to_numpy()
        grp = np.minimum((rank - 1) / len(rank) * n_groups, n_groups - 1).astype(int)
        rr = r[mask]
        for g in range(n_groups):
            sel = grp == g
            if sel.any():
                out[step, g] = rr[sel].mean()
    cols = [f"G{g + 1}" for g in range(n_groups)]
    df = pd.DataFrame(out, index=wd.dates, columns=cols)
    ls = df.iloc[:, -1] - df.iloc[:, 0]
    return df, ls

# ===========================================================================
# 第五部分  因子有效性检验
# ===========================================================================

# -*- coding: utf-8 -*-





# ---------------------------------------------------------------------------
# IC 分析
# ---------------------------------------------------------------------------
def ic_series(factor: np.ndarray, label: np.ndarray, idx: np.ndarray,
              min_n: int = 100) -> pd.Series:
    """逐期计算 Rank IC（Spearman）。"""
    vals = [rank_corr(factor[i], label[i], min_n=min_n) for i in idx]
    return pd.Series(vals, index=idx)


def ic_summary(ic: pd.Series) -> Dict[str, float]:
    """IC 序列的汇总统计。"""
    ic = ic.dropna()
    if len(ic) == 0:
        return {}
    mean = float(ic.mean())
    std = float(ic.std(ddof=1))
    return {
        "IC均值": mean,
        "IC标准差": std,
        "ICIR": float(mean / std) if std > 0 else np.nan,
        "IC为正比例": float((ic > 0).mean()),
        "t值": float(mean / std * np.sqrt(len(ic))) if std > 0 else np.nan,
        "样本期数": int(len(ic)),
    }


def ic_table(factors: Dict[str, np.ndarray], labels: Dict[str, np.ndarray],
             idx_by_label: Dict[str, np.ndarray], min_n: int = 100) -> pd.DataFrame:
    """多因子 × 多预测窗口的 IC 汇总表。"""
    rows = []
    for fname, f in factors.items():
        for lname, lab in labels.items():
            ic = ic_series(f, lab, idx_by_label[lname], min_n=min_n)
            s = ic_summary(ic)
            if not s:
                continue
            rows.append({"因子": fname, "预测窗口": lname, **s})
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# 分组检验
# ---------------------------------------------------------------------------
def group_summary(group_ret: pd.DataFrame, ls: pd.Series) -> pd.DataFrame:
    """分组年化收益、波动、夏普与多空绩效。"""
    rows = []
    n = len(group_ret.dropna(how="all"))
    for col in group_ret.columns:
        r = group_ret[col].dropna()
        if len(r) < 10:
            continue
        ann = float((1 + r).prod() ** (52 / len(r)) - 1)
        vol = float(r.std(ddof=1) * np.sqrt(52))
        rows.append({
            "分组": col,
            "年化收益": ann,
            "年化波动": vol,
            "夏普比率": float(ann / vol) if vol > 0 else np.nan,
            "周胜率": float((r > 0).mean()),
        })
    df = pd.DataFrame(rows)
    # 单调性：分组序号与年化收益的秩相关
    if len(df) > 2:
        rho = stats.spearmanr(range(len(df)), df["年化收益"]).statistic
        df.attrs["monotonicity"] = float(rho)
    return df


def long_short_stats(ls: pd.Series, name: str = "多空组合") -> Dict[str, float]:
    """多空组合绩效。"""
    r = ls.dropna()
    if len(r) < 10:
        return {}
    ann = float((1 + r).prod() ** (52 / len(r)) - 1)
    vol = float(r.std(ddof=1) * np.sqrt(52))
    return {
        "组合": name,
        "年化收益": ann,
        "年化波动": vol,
        "夏普比率": float(ann / vol) if vol > 0 else np.nan,
        "最大回撤": float(((1 + r).cumprod() / (1 + r).cumprod().cummax() - 1).min()),
        "周胜率": float((r > 0).mean()),
        "t值": float(r.mean() / r.std(ddof=1) * np.sqrt(len(r))) if r.std(ddof=1) > 0 else np.nan,
    }


# ---------------------------------------------------------------------------
# Fama-MacBeth 回归
# ---------------------------------------------------------------------------
def fama_macbeth(y: np.ndarray, X: Dict[str, np.ndarray], dates: pd.DatetimeIndex,
                 lags: int = 4) -> pd.DataFrame:
    """逐期横截面回归后对系数做 Newey-West 检验。

    Parameters
    ----------
    y : ndarray
        形状 ``(n_periods, n_stocks)`` 的因变量，通常为下一周收益率（%）。
    X : dict[str, ndarray]
        自变量矩阵，形状同 ``y``；函数内部逐期做横截面标准化。
    dates : DatetimeIndex
        每期对应的日期。

    Returns
    -------
    DataFrame
        各变量的平均回归系数、Newey-West t 值、显著比例与平均 R²。
    """
    names = list(X.keys())
    n_periods = y.shape[0]
    betas, r2s, used = [], [], []
    for t in range(n_periods):
        cols = [X[k][t] for k in names]
        mask = np.isfinite(y[t])
        for c in cols:
            mask &= np.isfinite(c)
        if mask.sum() < 50:
            continue
        Y = y[t][mask].astype(float)
        M = np.column_stack([np.ones(mask.sum())] + [c[mask] for c in cols]).astype(float)
        # 逐列标准化（常数项除外）
        mu = M[:, 1:].mean(axis=0)
        sd = M[:, 1:].std(axis=0)
        sd[sd == 0] = 1.0
        M[:, 1:] = (M[:, 1:] - mu) / sd
        beta, *_ = np.linalg.lstsq(M, Y, rcond=None)
        resid = Y - M @ beta
        ss_tot = ((Y - Y.mean()) ** 2).sum()
        betas.append(beta[1:])
        r2s.append(1 - (resid ** 2).sum() / ss_tot if ss_tot > 0 else np.nan)
        used.append(dates[t])
    if not betas:
        return pd.DataFrame()
    coef_df = pd.DataFrame(betas, index=used, columns=names)
    rows = []
    for col in names:
        s = coef_df[col]
        rows.append({
            "变量": col,
            "平均系数": float(s.mean()),
            "标准差": float(s.std(ddof=1)),
            "NW t值": newey_west_tstat(s, lags=lags),
            "t值": float(s.mean() / s.std(ddof=1) * np.sqrt(len(s))) if s.std(ddof=1) > 0 else np.nan,
            "系数为正比例": float((s > 0).mean()),
            "期数": int(len(s)),
            "平均R2": float(np.nanmean(r2s)),
        })
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# 因子相关性
# ---------------------------------------------------------------------------
def factor_correlations(factor: np.ndarray, controls: Dict[str, np.ndarray],
                        idx: np.ndarray, min_n: int = 100) -> pd.DataFrame:
    """因子与常用量价/风格因子的周度截面相关系数（Spearman）均值。"""
    rows = []
    for name, ctrl in controls.items():
        corrs = []
        for i in idx:
            c = rank_corr(factor[i], ctrl[i], min_n=min_n)
            if np.isfinite(c):
                corrs.append(c)
        rows.append({
            "对照因子": name,
            "平均相关系数": float(np.mean(corrs)) if corrs else np.nan,
            "相关系数标准差": float(np.std(corrs, ddof=1)) if len(corrs) > 1 else np.nan,
            "期数": len(corrs),
        })
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# 事件研究（经典 K 线形态）
# ---------------------------------------------------------------------------
def event_stats(mask: np.ndarray, excess: np.ndarray, dates: pd.DatetimeIndex,
                name: str) -> Tuple[Dict[str, float], pd.DataFrame]:
    """统计某类事件后未来 20 日超额收益的表现及分年度结果。"""
    ii, jj = np.nonzero(mask & np.isfinite(excess))
    if len(ii) == 0:
        return {}, pd.DataFrame()
    ex = excess[ii, jj].astype(float)
    df = pd.DataFrame({
        "date": dates[ii],
        "excess": ex,
    })
    yearly = df.groupby(df["date"].dt.year)["excess"].agg(["mean", "count"])
    yearly["win_rate"] = df.groupby(df["date"].dt.year)["excess"].apply(lambda x: (x > 0).mean())
    t = float(df["excess"].mean() / df["excess"].std(ddof=1) * np.sqrt(len(df))) if len(df) > 2 else np.nan
    stats_row = {
        "事件类型": name,
        "样本数": int(len(df)),
        "超额收益均值": float(df["excess"].mean()),
        "超额收益中位数": float(df["excess"].median()),
        "胜率": float((df["excess"] > 0).mean()),
        "t值": t,
        "平均发生股票数/期": float(len(df) / df["date"].nunique()),
    }
    yearly = yearly.rename(columns={"mean": "超额收益均值", "count": "样本数"})
    yearly["事件类型"] = name
    yearly = yearly.reset_index().rename(columns={"date": "年份"})
    return stats_row, yearly


def event_study_table(panel, excess: np.ndarray, events) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """汇总四类经典形态事件的统计结果。"""
    combos = {
        "底部十字星+放量": events.bottom & events.doji & events.vol_up,
        "底部锤头线+放量": events.bottom & events.hammer & events.vol_up,
        "顶部十字星+放量": events.top & events.doji & events.vol_up,
        "顶部倒锤头线+放量": events.top & events.inv_hammer & events.vol_up,
        "底部反转形态(合并)": events.bullish,
        "顶部反转形态(合并)": events.bearish,
    }
    rows, years = [], []
    for name, m in combos.items():
        s, y = event_stats(m, excess, panel.dates, name)
        if s:
            rows.append(s)
        if len(y):
            years.append(y)
    df = pd.DataFrame(rows)
    yd = pd.concat(years, ignore_index=True) if years else pd.DataFrame()
    return df, yd

# ===========================================================================
# 第六部分  主流程：数据 → 因子 → 检验 → 回测
# ===========================================================================

# -*- coding: utf-8 -*-






DATA_DIR = os.path.join(ROOT, "数据", "处理后数据")
RESULT_DIR = os.path.join(ROOT, "结果")
TABLE_DIR = os.path.join(RESULT_DIR, "tables")
SERIES_DIR = os.path.join(RESULT_DIR, "series")
for d in (DATA_DIR, TABLE_DIR, SERIES_DIR):
    os.makedirs(d, exist_ok=True)

START_EVAL = "2021-01-01"
END_EVAL = "2025-11-28"


def log(msg: str) -> None:
    print(f"[{time.strftime('%H:%M:%S')}] {msg}", flush=True)


def jsonable(obj):
    """把 numpy / pandas 对象转换为可 JSON 序列化的类型。"""
    if isinstance(obj, dict):
        return {str(k): jsonable(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [jsonable(v) for v in obj]
    if isinstance(obj, (np.floating, float)):
        return None if not np.isfinite(obj) else float(obj)
    if isinstance(obj, (np.integer,)):
        return int(obj)
    if isinstance(obj, (np.bool_,)):
        return bool(obj)
    if isinstance(obj, (pd.Timestamp,)):
        return obj.strftime("%Y-%m-%d")
    if isinstance(obj, np.ndarray):
        return jsonable(obj.tolist())
    return obj


# ---------------------------------------------------------------------------
# 1. 数据
# ---------------------------------------------------------------------------
def load_data() -> tuple:
    log("载入面板数据 ...")
    panel = build_panel(start="2020-01-01", end="2025-12-31",
                        cache_path=os.path.join(DATA_DIR, "panel.npz"))
    benchmark = read_benchmark()
    industry = read_industry()
    log(f"面板维度 {panel.shape}，区间 {panel.dates.min().date()} ~ {panel.dates.max().date()}")
    return panel, benchmark, industry


def build_labels(panel: Panel) -> dict:
    adj = adjusted_prices(panel)
    labels = {}
    for h in (5, 10, 20, 60):
        r = forward_return(adj["close"], h)
        labels[h] = excess_over_market(r)
    return labels


# ---------------------------------------------------------------------------
# 2. 因子
# ---------------------------------------------------------------------------
def build_factors(panel: Panel, labels: dict) -> tuple:
    cut_idx, cutoffs = score_cutoffs(panel.dates, horizon=20, min_train_days=120)
    log(f"形态绩效按月重估：{len(cutoffs)} 期，"
        f"{panel.dates[cut_idx[0]].date()} ~ {panel.dates[cut_idx[-1]].date()}")

    factors, diagnostics = {}, []
    for k in range(1, 11):
        t0 = time.time()
        res = build_csp(panel, k, labels[20], cutoffs, cut_idx, min_n=20, kappa=50.0)
        factors[f"CSP_{k}"] = res.factor
        diagnostics.append({**res.diagnostics, "耗时秒": round(time.time() - t0, 2)})
        log(f"  CSP_{k}: 覆盖 {res.coverage:.1%}, 有效形态 {res.n_scored}, "
            f"{time.time() - t0:.1f}s")

    factors["CSP_复合"] = composite_factor([factors[f"CSP_{k}"] for k in range(1, 11)])
    diagnostics.append({
        "k": "复合", "n_bars": np.nan, "n_patterns_observed": np.nan,
        "n_patterns_scored": np.nan,
        "coverage_all": float(np.isfinite(factors["CSP_复合"]).mean()),
        "mean": float(np.nanmean(factors["CSP_复合"])),
        "std": float(np.nanstd(factors["CSP_复合"])),
        "耗时秒": np.nan,
    })
    return factors, pd.DataFrame(diagnostics), cut_idx, cutoffs


# ---------------------------------------------------------------------------
# 3. 因子检验
# ---------------------------------------------------------------------------
def eval_dates(panel: Panel) -> tuple:
    wd = weekly_last_dates(panel.dates)
    wd = wd[(wd >= pd.Timestamp(START_EVAL)) & (wd <= pd.Timestamp(END_EVAL))]
    md = month_last_dates(panel.dates)
    md = md[(md >= pd.Timestamp(START_EVAL)) & (md <= pd.Timestamp(END_EVAL))]
    return wd, panel.dates.get_indexer(wd), md, panel.dates.get_indexer(md)


def run_factor_tests(panel: Panel, factors: dict, labels: dict,
                     weekly: WeeklyData) -> dict:
    wd, wi, md, mi = eval_dates(panel)
    log(f"因子检验：周度 {len(wi)} 期，月度 {len(mi)} 期")

    sel = {name: factors[name] for name in ("CSP_1", "CSP_5", "CSP_10", "CSP_复合")}
    label_names = {5: "未来5日", 10: "未来10日", 20: "未来20日", 60: "未来60日"}
    labels_plot = {label_names[h]: labels[h] for h in (5, 10, 20, 60)}
    idx_by_label = {label_names[h]: wi for h in (5, 10, 20, 60)}
    ic_tbl = ic_table(sel, labels_plot, idx_by_label)

    # 月度 IC（与研报口径一致，便于横向对比）
    ic_m = {name: ic_series(f, labels[20], mi) for name, f in sel.items()}
    monthly_rows = []
    for name, ic in ic_m.items():
        s = ic_summary(ic)
        monthly_rows.append({"因子": name, **s})
    ic_monthly_tbl = pd.DataFrame(monthly_rows)

    # IC 序列（周度，未来 20 日）
    ic_w = {name: ic_series(f, labels[20], wi) for name, f in sel.items()}
    ic_df = pd.DataFrame({k: v.to_numpy() for k, v in ic_w.items()}, index=panel.dates[wi])
    ic_month_df = pd.DataFrame({k: v.to_numpy() for k, v in ic_m.items()}, index=panel.dates[mi])

    # 分组
    groups, ls = {}, {}
    group_tables = []
    ls_rows = []
    mono = {}
    for name in ("CSP_5", "CSP_复合"):
        g, l = decile_backtest(weekly, factors[name])
        groups[name], ls[name] = g, l
        gt = group_summary(g, l)
        mono[name] = float(gt.attrs.get("monotonicity", np.nan))
        gt.insert(0, "因子", name)
        group_tables.append(gt)
        st = long_short_stats(l, name)
        st["单调性"] = mono[name]
        ls_rows.append(st)
    group_tbl = pd.concat(group_tables, ignore_index=True)
    ls_tbl = pd.DataFrame(ls_rows)

    # Fama-MacBeth：下一周收益对 CSP 与常用风格因子
    controls = build_control_factors(panel)
    fmb_rows = []
    y = weekly.next_ret * 100.0
    for name in ("CSP_5", "CSP_复合"):
        X = {"CSP": factors[name][weekly.idx]}
        for cname in ("市值", "BP", "动量", "反转", "波动率"):
            X[cname] = controls[cname][weekly.idx]
        res = fama_macbeth(y, X, weekly.dates, lags=4)
        res.insert(0, "因子", name)
        fmb_rows.append(res)
    fmb_tbl = pd.concat(fmb_rows, ignore_index=True)

    # 相关性
    corr_tbl = factor_correlations(factors["CSP_5"], controls, wi)

    return {
        "WD": wd, "WI": wi, "MD": md, "MI": mi,
        "ic_table": ic_tbl, "ic_monthly": ic_monthly_tbl,
        "ic_series": ic_df, "ic_month_series": ic_month_df,
        "group_table": group_tbl, "ls_table": ls_tbl,
        "groups": groups, "ls": ls,
        "fmb_table": fmb_tbl, "corr_table": corr_tbl,
        "weekly": weekly, "monotonicity": mono,
    }


def build_control_factors(panel: Panel) -> dict:
    """构造用于相关性分析与回归控制的常用量价因子。"""
    adj = adjusted_prices(panel)
    close = adj["close"]
    with np.errstate(invalid="ignore", divide="ignore"):
        ret = close / np.roll(close, 1, axis=0) - 1
    # 因子值均为“越高越好”的方向不重要，只用于相关性
    mom = np.full_like(close, np.nan)
    with np.errstate(invalid="ignore", divide="ignore"):
        mom[20:] = close[20:] / close[:-20] - 1
    rev = np.full_like(close, np.nan)
    with np.errstate(invalid="ignore", divide="ignore"):
        rev[5:] = close[5:] / close[:-5] - 1
    vol = pd.DataFrame(ret).rolling(20, min_periods=15).std().to_numpy(dtype="float32")
    turnover = np.where(panel["total_mv"] > 0, panel["amount"] / panel["total_mv"], np.nan)
    amihud = np.where(panel["amount"] > 0, np.abs(ret) / panel["amount"] * 1e8, np.nan)
    amihud = pd.DataFrame(amihud).rolling(20, min_periods=15).mean().to_numpy(dtype="float32")
    size = np.log(panel["total_mv"])
    with np.errstate(invalid="ignore", divide="ignore"):
        bp = 1.0 / panel["pb"]
    return {
        "市值": size, "BP": bp, "动量": mom, "反转": rev,
        "波动率": vol, "换手率": turnover, "非流动性": amihud,
    }


# ---------------------------------------------------------------------------
# 4. 事件研究
# ---------------------------------------------------------------------------
def run_event_study(panel: Panel, labels: dict) -> tuple:
    log("经典 K 线形态事件研究 ...")
    events = build_classic_events(panel)
    table, yearly = event_study_table(panel, labels[20], events)
    return events, table, yearly


# ---------------------------------------------------------------------------
# 5. 回测
# ---------------------------------------------------------------------------
def prepare_scores(panel: Panel, weekly: WeeklyData, factors: dict,
                   controls: dict, events) -> dict:
    """在每个调仓日计算基准模型与研究模型的综合得分。"""
    ind_code = industry_codes(panel.codes, read_industry())
    n_weeks, n_stocks = len(weekly.idx), panel.shape[1]
    z = {name: np.full((n_weeks, n_stocks), np.nan, dtype="float32")
         for name in ("size", "bp", "csp5", "cspc", "event")}

    log("逐周计算因子得分 ...")
    for step, i in enumerate(weekly.idx):
        mask = weekly.eligible[step] & np.isfinite(panel["total_mv"][i]) & (panel["total_mv"][i] > 0)
        size = np.log(panel["total_mv"][i])
        with np.errstate(invalid="ignore", divide="ignore"):
            bp = 1.0 / panel["pb"][i]
        bp = np.where(panel["pb"][i] > 0, bp, np.nan)
        z["size"][step] = prep_cross_section(size, ind_code, winsor=(0.01, 0.99))
        z["bp"][step] = prep_cross_section(bp, ind_code, winsor=(0.01, 0.99))
        z["csp5"][step] = prep_cross_section(factors["CSP_5"][i], ind_code, winsor=(0.01, 0.99))
        z["cspc"][step] = prep_cross_section(factors["CSP_复合"][i], ind_code, winsor=(0.01, 0.99))
        z["event"][step] = prep_cross_section(events.score[i], ind_code, winsor=(0.01, 0.99))

    score_bm = 0.5 * (-z["size"]) + 0.5 * z["bp"]
    score_rm5 = (-z["size"] + z["bp"] + z["csp5"]) / 3.0
    score_rmc = (-z["size"] + z["bp"] + z["cspc"]) / 3.0
    score_rme = (-z["size"] + z["bp"] + z["event"]) / 3.0
    # 改进模型：对复合 CSP 信号做 4 周平滑，降低形态信号的高频切换
    smooth = pd.DataFrame(z["cspc"]).rolling(4, min_periods=1).mean().to_numpy(dtype="float32")
    score_rms = (-z["size"] + z["bp"] + smooth) / 3.0
    return {"bm": score_bm, "rm5": score_rm5, "rmc": score_rmc,
            "rme": score_rme, "rms": score_rms, "z": z, "csp_smooth": smooth}


def run_backtests(panel: Panel, weekly: WeeklyData, scores: dict) -> dict:
    log("执行周度多因子回测 ...")
    specs = [
        ("基准模型(Size+BP)", "bm"),
        ("研究模型(Size+BP+复合CSP)", "rmc"),
        ("研究模型(Size+BP+CSP5)", "rm5"),
        ("研究模型(Size+BP+经典形态)", "rme"),
        ("改进模型(复合CSP平滑)", "rms"),
    ]
    results = {}
    for name, key in specs:
        results[name] = run_strategy(weekly, scores[key], name, top_frac=0.1, fee=0.0015)
    return results


def cost_attribution(weekly: WeeklyData, scores: dict) -> pd.DataFrame:
    """比较各模型毛收益、净收益与交易成本拖累。"""
    rows = []
    for name, key in (("基准模型(Size+BP)", "bm"),
                      ("研究模型(Size+BP+复合CSP)", "rmc"),
                      ("改进模型(复合CSP平滑)", "rms")):
        gross = run_strategy(weekly, scores[key], name, top_frac=0.1, fee=0.0)
        net = run_strategy(weekly, scores[key], name, top_frac=0.1, fee=0.0015)
        g = annualized_return(gross.ret)
        n = annualized_return(net.ret)
        to = float(net.turnover.mean())
        rows.append({
            "模型": name,
            "毛年化收益": g,
            "净年化收益": n,
            "成本拖累(年化)": g - n,
            "平均周换手率": to,
            "年化换手倍数": to * 52,
            "假设单边成本": 0.0015,
        })
    return pd.DataFrame(rows)


def sensitivity_analysis(weekly: WeeklyData, scores: dict) -> pd.DataFrame:
    """参数敏感性：组合分位、交易成本、调仓频率（主研究模型）。"""
    rows = []
    for top in (0.05, 0.10, 0.20):
        for fee in (0.0, 0.0015, 0.003):
            for reb in (1, 2, 4):
                res = run_strategy(weekly, scores["rmc"], "研究模型", top_frac=top,
                                      fee=fee, rebalance_weeks=reb)
                s = res.summary()
                rows.append({
                    "组合分位": top, "单边成本": fee, "调仓周期(周)": reb,
                    "年化收益": s["年化收益率"], "年化波动": s["年化波动率"],
                    "夏普比率": s["夏普比率"], "最大回撤": s["最大回撤"],
                })
    return pd.DataFrame(rows)


def weight_sensitivity(weekly: WeeklyData, scores: dict) -> pd.DataFrame:
    """核心因子权重敏感性：CSP 在综合得分中的权重从 0 到 80%。"""
    z = scores["z"]
    rows = []
    for w in (0.0, 0.2, 1 / 3, 0.5, 0.8):
        sc = (1 - w) * (-z["size"] + z["bp"]) / 2 + w * z["cspc"]
        res = run_strategy(weekly, sc, f"CSP权重{w:.0%}", top_frac=0.1, fee=0.0015)
        s = res.summary()
        rows.append({
            "CSP权重": w, "年化收益": s["年化收益率"], "年化波动": s["年化波动率"],
            "夏普比率": s["夏普比率"], "最大回撤": s["最大回撤"],
            "平均换手率": float(res.turnover.mean()),
        })
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------
def run_pipeline() -> None:
    t_start = time.time()
    panel, benchmark, industry = load_data()
    labels = build_labels(panel)

    # ---- 因子
    factors, diag_tbl, cut_idx, cutoffs = build_factors(panel, labels)
    save_table(diag_tbl, os.path.join(TABLE_DIR, "t01_因子构建诊断.csv"))

    # ---- 因子检验
    weekly = build_weekly_data(panel, benchmark["close"])
    tests = run_factor_tests(panel, factors, labels, weekly)
    save_table(tests["ic_table"], os.path.join(TABLE_DIR, "t02_IC汇总.csv"))
    save_table(tests["ic_monthly"], os.path.join(TABLE_DIR, "t03_月度IC.csv"))
    save_table(tests["group_table"], os.path.join(TABLE_DIR, "t04_分组收益.csv"))
    save_table(tests["ls_table"], os.path.join(TABLE_DIR, "t05_多空组合.csv"))
    save_table(tests["fmb_table"], os.path.join(TABLE_DIR, "t06_FamaMacBeth.csv"))
    save_table(tests["corr_table"], os.path.join(TABLE_DIR, "t07_因子相关性.csv"))
    tests["ic_series"].to_csv(os.path.join(SERIES_DIR, "ic_weekly.csv"),
                              encoding="utf-8-sig", float_format="%.6f")
    tests["ic_month_series"].to_csv(os.path.join(SERIES_DIR, "ic_monthly.csv"),
                                    encoding="utf-8-sig", float_format="%.6f")

    # ---- 事件研究
    events, event_tbl, event_year = run_event_study(panel, labels)
    save_table(event_tbl, os.path.join(TABLE_DIR, "t08_经典形态事件研究.csv"))
    save_table(event_year, os.path.join(TABLE_DIR, "t09_事件研究分年度.csv"))

    # ---- 回测
    controls = build_control_factors(panel)
    scores = prepare_scores(panel, weekly, factors, controls, events)
    results = run_backtests(panel, weekly, scores)

    bench = weekly.bm_ret
    all_a = weekly.all_ret
    perf_tbl = strategy_table(list(results.values()), bench)
    all_row = perf_summary(all_a.dropna(), "全A等权")
    all_row.update({"年化超额(vs沪深300)": float((all_a.dropna() - bench.reindex(all_a.dropna().index)).dropna().mean() * 52)})
    perf_tbl = pd.concat([perf_tbl, pd.DataFrame([all_row])], ignore_index=True)
    save_table(perf_tbl, os.path.join(TABLE_DIR, "t10_回测绩效.csv"))

    # 相对基准模型的超额
    rm = results["研究模型(Size+BP+复合CSP)"]
    bm_res = results["基准模型(Size+BP)"]
    ex_vs_bm = (rm.ret - bm_res.ret).dropna()
    extra = {
        "相对基准模型年化超额": float(ex_vs_bm.mean() * 52),
        "相对基准模型超额胜率": float((ex_vs_bm > 0).mean()),
        "相对基准模型信息比率": float(ex_vs_bm.mean() * 52 / (ex_vs_bm.std(ddof=1) * np.sqrt(52))),
        "相对沪深300年化超额": float((rm.ret - bench.reindex(rm.ret.index)).dropna().mean() * 52),
    }

    # ---- 保存序列
    ret_all = pd.DataFrame({
        "沪深300": bench,
        "全A等权": all_a,
        **{name: r.ret for name, r in results.items()},
    })
    ret_all.to_csv(os.path.join(SERIES_DIR, "weekly_returns.csv"),
                   encoding="utf-8-sig", float_format="%.6f")

    nav = pd.DataFrame({name: r.nav for name, r in results.items()})
    nav["沪深300"] = (1 + bench.dropna()).cumprod()
    nav["全A等权"] = (1 + all_a.dropna()).cumprod()
    nav.to_csv(os.path.join(SERIES_DIR, "nav.csv"), encoding="utf-8-sig", float_format="%.6f")

    dd = pd.DataFrame({name: max_drawdown_series(r.ret) for name, r in results.items()})
    dd["沪深300"] = max_drawdown_series(bench)
    dd.to_csv(os.path.join(SERIES_DIR, "drawdown.csv"), encoding="utf-8-sig", float_format="%.6f")

    yearly = {}
    for name, r in results.items():
        yearly[name] = yearly_returns(r.ret)
    yearly["沪深300"] = yearly_returns(bench)
    yearly["全A等权"] = yearly_returns(all_a)
    yearly_df = pd.DataFrame(yearly)
    yearly_df.index.name = "年份"
    save_table(yearly_df, os.path.join(TABLE_DIR, "t11_分年度收益.csv"), index=True)

    # 分组收益序列
    for name, g in tests["groups"].items():
        tag = "csp5" if name == "CSP_5" else "comp"
        g.to_csv(os.path.join(SERIES_DIR, f"group_ret_{tag}.csv"),
                 encoding="utf-8-sig", float_format="%.6f")
        tests["ls"][name].to_csv(os.path.join(SERIES_DIR, f"ls_{tag}.csv"),
                                 encoding="utf-8-sig", float_format="%.6f", header=True)

    # ---- 敏感性
    sens = sensitivity_analysis(weekly, scores)
    save_table(sens, os.path.join(TABLE_DIR, "t12_参数敏感性.csv"))
    weight_sens = weight_sensitivity(weekly, scores)
    save_table(weight_sens, os.path.join(TABLE_DIR, "t13_因子权重敏感性.csv"))
    cost_tbl = cost_attribution(weekly, scores)
    save_table(cost_tbl, os.path.join(TABLE_DIR, "t14_成本归因.csv"))

    # ---- 处理后数据
    log("保存处理后数据 ...")
    # 周度因子快照（z 值）
    fac_weekly = pd.DataFrame(
        scores["z"]["csp5"], index=weekly.dates, columns=panel.codes
    )
    fac_weekly.index.name = "date"
    fac_weekly.to_csv(os.path.join(DATA_DIR, "因子值_周度_CSP.csv"),
                      encoding="utf-8-sig", float_format="%.4f")
    np.savez_compressed(
        os.path.join(DATA_DIR, "csp_factors_daily.npz"),
        dates=panel.dates.values.astype("datetime64[ns]"),
        codes=np.asarray(panel.codes, dtype=object),
        CSP_1=factors["CSP_1"], CSP_5=factors["CSP_5"],
        CSP_10=factors["CSP_10"], CSP_composite=factors["CSP_复合"],
    )
    # 面板基础统计
    valid = np.isfinite(panel["close"]) & (panel["close"] > 0)
    stats = pd.DataFrame({
        "日期": panel.dates,
        "有效交易股票数": valid.sum(axis=1),
        "因子覆盖股票数": np.isfinite(factors["CSP_复合"]).sum(axis=1),
    })
    save_table(stats, os.path.join(DATA_DIR, "样本覆盖统计.csv"))

    # 形态分布
    dist = pattern_distribution(panel)
    save_table(dist, os.path.join(TABLE_DIR, "t15_单K线形态分布.csv"))

    # ---- 汇总
    summary = {
        "数据": {
            "股票数": int(panel.shape[1]),
            "交易日数": int(panel.shape[0]),
            "起始日": panel.dates.min(),
            "结束日": panel.dates.max(),
            "平均每日有效股票数": float(valid.sum(axis=1).mean()),
        },
        "因子构建": diag_tbl.to_dict("records"),
        "IC": tests["ic_table"].to_dict("records"),
        "月度IC": tests["ic_monthly"].to_dict("records"),
        "分组": tests["group_table"].to_dict("records"),
        "多空": tests["ls_table"].to_dict("records"),
        "FamaMacBeth": tests["fmb_table"].to_dict("records"),
        "相关性": tests["corr_table"].to_dict("records"),
        "事件研究": event_tbl.to_dict("records"),
        "回测绩效": perf_tbl.to_dict("records"),
        "相对基准": extra,
        "分年度": yearly_df.reset_index().to_dict("records"),
        "IC序列": {
            "CSP_5": tests["ic_series"]["CSP_5"].to_dict(),
            "CSP_复合": tests["ic_series"]["CSP_复合"].to_dict(),
        },
        "参数敏感性": sens.to_dict("records"),
        "因子权重敏感性": weight_sens.to_dict("records"),
        "成本归因": cost_tbl.to_dict("records"),
        "组内单调性": {name: float(v) for name, v in tests["monotonicity"].items()},
        "回测区间": {
            "起始": str(pd.concat([r.ret for r in results.values()]).first_valid_index().date()),
            "结束": str(pd.concat([r.ret for r in results.values()]).last_valid_index().date()),
        },
        "运行耗时秒": round(time.time() - t_start, 1),
    }
    with open(os.path.join(RESULT_DIR, "results.json"), "w", encoding="utf-8") as f:
        json.dump(jsonable(summary), f, ensure_ascii=False, indent=2)

    log(f"完成，耗时 {time.time() - t_start:.1f}s")
    print(perf_tbl.to_string())
    print("\nIC 汇总（前 8 行）")
    print(tests["ic_table"].head(8).to_string())
    print("\n相对基准模型：", extra)





# ===========================================================================
# 第七部分  图表生成
# ===========================================================================

# -*- coding: utf-8 -*-




TABLE_DIR = os.path.join(ROOT, "结果", "tables")
SERIES_DIR = os.path.join(ROOT, "结果", "series")
FIG_DIR = os.path.join(ROOT, "结果", "figures")
os.makedirs(FIG_DIR, exist_ok=True)

# ---- 全局样式
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
C_PURPLE = "#7D5BA6"


def _save(fig, name: str) -> None:
    path = os.path.join(FIG_DIR, name)
    fig.savefig(path, facecolor="white")
    plt.close(fig)
    print("saved", path)


def read_table(name: str, **kw) -> pd.DataFrame:
    return pd.read_csv(os.path.join(TABLE_DIR, name), **kw)


def read_series(name: str, **kw) -> pd.DataFrame:
    return pd.read_csv(os.path.join(SERIES_DIR, name), index_col=0, parse_dates=True, **kw)


# ---------------------------------------------------------------------------
def fig_sample() -> None:
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


def fig_pattern_dist() -> None:
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


def fig_ic() -> None:
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


def fig_groups() -> None:
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


def fig_nav() -> None:
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


def fig_drawdown() -> None:
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


def fig_yearly() -> None:
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


def fig_event() -> None:
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


def fig_corr() -> None:
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


def fig_sensitivity() -> None:
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


def run_figures() -> None:
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







# ===========================================================================
# 第八部分  Word 报告生成
# ===========================================================================

# -*- coding: utf-8 -*-




TABLE_DIR = os.path.join(ROOT, "结果", "tables")
FIG_DIR = os.path.join(ROOT, "结果", "figures")
OUT_PATH = os.path.join(ROOT, "202630196102-叶承恩-期末实验报告-K线形态信息中的alpha复刻.docx")

# ---------------------------------------------------------------------------
# 读取结果（延迟到 run_report 执行时加载，便于单文件脚本先跑完回测）
# ---------------------------------------------------------------------------
R: Dict = {}
perf: Dict = {}
monthly_ic: Dict = {}
weekly_ic: Dict = {}
ls: Dict = {}
corr: Dict = {}
ev: Dict = {}
cost: Dict = {}
wsen: Dict = {}
yearly_rows: List = []

BM = "基准模型(Size+BP)"
RM = "研究模型(Size+BP+复合CSP)"
RM5 = "研究模型(Size+BP+CSP5)"
RME = "研究模型(Size+BP+经典形态)"
RMI = "改进模型(复合CSP平滑)"


def T(name: str) -> pd.DataFrame:
    return pd.read_csv(os.path.join(TABLE_DIR, name), encoding="utf-8-sig")


def load_results() -> None:
    """加载主流程输出的 results.json，并整理为报告所需的字典。"""
    global R, perf, monthly_ic, weekly_ic, ls, corr, ev, cost, wsen, yearly_rows
    with open(os.path.join(ROOT, "结果", "results.json"), encoding="utf-8") as fh:
        R = json.load(fh)
    perf = {row["名称"]: row for row in R["回测绩效"]}
    monthly_ic = {row["因子"]: row for row in R["月度IC"]}
    weekly_ic = {(row["因子"], row["预测窗口"]): row for row in R["IC"]}
    ls = {row["组合"]: row for row in R["多空"]}
    corr = {row["对照因子"]: row for row in R["相关性"]}
    ev = {row["事件类型"]: row for row in R["事件研究"]}
    cost = {row["模型"]: row for row in R["成本归因"]}
    wsen = {round(row["CSP权重"], 4): row for row in R["因子权重敏感性"]}
    yearly_rows = R["分年度"]


def pct(x: float, d: int = 2) -> str:
    return f"{x * 100:.{d}f}%"


def num(x: float, d: int = 3) -> str:
    return f"{x:.{d}f}"


def pp(x: float, d: int = 2) -> str:
    """百分点（已是百分比数值）。"""
    return f"{x:.{d}f}"


# ---------------------------------------------------------------------------
# Word 基础工具
# ---------------------------------------------------------------------------
CN_BODY = "宋体"
CN_HEAD = "黑体"
EN_FONT = "Times New Roman"


def set_run_font(run, size: float = 12, bold: bool = False,
                 cn: str = CN_BODY, en: str = EN_FONT, color=None) -> None:
    run.font.name = en
    run.font.size = Pt(size)
    run.bold = bold
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:ascii"), en)
    rfonts.set(qn("w:hAnsi"), en)
    rfonts.set(qn("w:eastAsia"), cn)
    if color is not None:
        run.font.color.rgb = color


def add_para(doc, text: str = "", size: float = 12, bold: bool = False,
             indent: bool = True, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
             space_before: float = 0, space_after: float = 0,
             line: float = 1.5, cn: str = CN_BODY, keep_with_next: bool = False):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.alignment = align
    pf.line_spacing = line
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    if indent:
        pf.first_line_indent = Pt(size * 2)
    if keep_with_next:
        pf.keep_with_next = True
    if text:
        run = p.add_run(text)
        set_run_font(run, size=size, bold=bold, cn=cn)
    return p


def add_heading(doc, text: str, level: int = 1) -> None:
    sizes = {1: 15, 2: 13, 3: 12}
    before = {1: 14, 2: 10, 3: 8}
    after = {1: 8, 2: 6, 3: 4}
    add_para(doc, text, size=sizes[level], bold=True, indent=False,
             align=WD_ALIGN_PARAGRAPH.LEFT, space_before=before[level],
             space_after=after[level], cn=CN_HEAD, keep_with_next=True)


_M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _latex_to_omml(latex: str, size_pt: float = 12):
    """LaTeX → MathML → OMML（Word 原生公式）。"""
    mathml = _latex2mathml.convert(latex)
    tree = etree.fromstring(mathml.encode("utf-8"))
    root = _MATH_XSLT(tree).getroot()
    sz = str(int(round(size_pt * 2)))
    for r in root.iter(f"{{{_M_NS}}}r"):
        t = r.find(f"{{{_M_NS}}}t")
        rpr = etree.Element(f"{{{_W_NS}}}rPr")
        s1 = etree.SubElement(rpr, f"{{{_W_NS}}}sz"); s1.set(f"{{{_W_NS}}}val", sz)
        s2 = etree.SubElement(rpr, f"{{{_W_NS}}}szCs"); s2.set(f"{{{_W_NS}}}val", sz)
        if t is not None:
            r.insert(list(r).index(t), rpr)
        else:
            r.append(rpr)
    return root


def _add_display_equation(doc, latex: str, number: Optional[int] = None) -> None:
    """把公式排成“居中公式 + 右对齐编号”的无框三栏表。"""
    tbl = doc.add_table(rows=1, cols=3)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    tblPr = tbl._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "none")
        borders.append(el)
    tblPr.insert_element_before(borders, "w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook")
    widths = (Cm(2.0), Cm(11.9), Cm(2.0))
    for j, w in enumerate(widths):
        tbl.columns[j].width = w
        tbl.cell(0, j).width = w
    trPr = tbl.rows[0]._tr.get_or_add_trPr()
    trPr.append(OxmlElement("w:cantSplit"))
    for cell in tbl.rows[0].cells:
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    p = tbl.cell(0, 1).paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.0
    omp = etree.SubElement(p._p, f"{{{_M_NS}}}oMathPara")
    pr = etree.SubElement(omp, f"{{{_M_NS}}}oMathParaPr")
    jc = etree.SubElement(pr, f"{{{_M_NS}}}jc")
    jc.set(f"{{{_M_NS}}}val", "center")
    omp.append(_latex_to_omml(latex))

    if number is not None:
        p2 = tbl.cell(0, 2).paragraphs[0]
        p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p2.paragraph_format.space_before = Pt(2)
        p2.paragraph_format.space_after = Pt(2)
        p2.paragraph_format.line_spacing = 1.0
        run = p2.add_run(f"（{number}）")
        set_run_font(run, size=12)

    # 用极小段落隔开相邻公式表，避免 Word 合并表格
    gap = doc.add_paragraph()
    gap.paragraph_format.space_before = Pt(0)
    gap.paragraph_format.space_after = Pt(0)
    gap.paragraph_format.line_spacing = 1.0
    gap_run = gap.add_run("")
    set_run_font(gap_run, size=2)


def add_formula(doc, latex: str, number: Optional[int] = None,
                plain: Optional[str] = None) -> None:
    """插入公式：优先使用 Word 原生公式，环境不支持时回退为纯文本。"""
    if _MATH_XSLT is not None:
        try:
            _add_display_equation(doc, latex, number)
            return
        except Exception:
            pass
    text = plain if plain is not None else latex
    p = doc.add_paragraph()
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(text)
    set_run_font(run, size=11, cn=CN_BODY)
    if number is not None:
        run2 = p.add_run(f"    （{number}）")
        set_run_font(run2, size=11)


def add_caption(doc, text: str) -> None:
    add_para(doc, text, size=10.5, indent=False, align=WD_ALIGN_PARAGRAPH.CENTER,
             space_before=4, space_after=3, line=1.2)


def add_figure(doc, filename: str, caption: str, width_cm: float = 15.5) -> None:
    path = os.path.join(FIG_DIR, filename)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run()
    run.add_picture(path, width=Cm(width_cm))
    add_caption(doc, caption)


def _set_cell_border(cell, **kwargs) -> None:
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement("w:tcBorders")
        tcPr.append(tcBorders)
    for edge in ("left", "top", "right", "bottom"):
        if edge in kwargs:
            tag = f"w:{edge}"
            element = tcBorders.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                tcBorders.append(element)
            for key in ("sz", "val", "color", "space"):
                if key in kwargs[edge]:
                    element.set(qn(f"w:{key}"), str(kwargs[edge][key]))


def _three_line_table(table) -> None:
    tbl = table._tbl
    tblPr = tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        if edge in ("top", "bottom"):
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), "12")
        elif edge == "insideH":
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), "4")
        else:
            el.set(qn("w:val"), "none")
        el.set(qn("w:color"), "000000")
        borders.append(el)
    tblPr.insert_element_before(
        borders, "w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook"
    )


def add_table(doc, df: pd.DataFrame, caption: str, font_size: float = 9,
              col_widths: Optional[Sequence[float]] = None,
              align_first_left: bool = True, header_bold: bool = True) -> None:
    add_caption(doc, caption)
    table = doc.add_table(rows=len(df) + 1, cols=len(df.columns))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    _three_line_table(table)
    # 表头
    for j, col in enumerate(df.columns):
        cell = table.cell(0, j)
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.line_spacing = 1.0
        run = p.add_run(str(col))
        set_run_font(run, size=font_size, bold=header_bold)
    # 数据
    for i, (_, row) in enumerate(df.iterrows(), start=1):
        for j, val in enumerate(row):
            cell = table.cell(i, j)
            cell.text = ""
            p = cell.paragraphs[0]
            p.alignment = (WD_ALIGN_PARAGRAPH.LEFT
                           if (j == 0 and align_first_left) or len(str(val)) > 18
                           else WD_ALIGN_PARAGRAPH.CENTER)
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)
            p.paragraph_format.line_spacing = 1.0
            run = p.add_run("" if pd.isna(val) else str(val))
            set_run_font(run, size=font_size)
    if col_widths:
        total = float(sum(col_widths))
        scale = min(1.0, 15.6 / total) if total > 0 else 1.0
        table.autofit = False
        for j, w in enumerate(col_widths):
            table.columns[j].width = Cm(w * scale)
            for cell in table.columns[j].cells:
                cell.width = Cm(w * scale)
    # 跨页时重复表头
    trPr = table.rows[0]._tr.get_or_add_trPr()
    th = OxmlElement("w:tblHeader")
    th.set(qn("w:val"), "true")
    trPr.append(th)

    # 表头下横线加粗
    for cell in table.rows[0].cells:
        _set_cell_border(cell, bottom={"val": "single", "sz": "8", "color": "000000"})
    add_para(doc, "", size=6, indent=False, space_after=2)


def add_page_number(section) -> None:
    footer = section.footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = "PAGE"
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    txt = OxmlElement("w:t")
    txt.text = "1"
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run._r.append(fld_begin)
    run._r.append(instr)
    run._r.append(fld_sep)
    run._r.append(txt)
    run._r.append(fld_end)
    set_run_font(run, size=10.5)


def setup_document() -> Document:
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    for attr in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(sec, attr, Cm(2.55))
    sec.header_distance = Cm(1.5)
    sec.footer_distance = Cm(1.5)

    style = doc.styles["Normal"]
    style.font.name = EN_FONT
    style.font.size = Pt(12)
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    rfonts.set(qn("w:eastAsia"), CN_BODY)
    rfonts.set(qn("w:ascii"), EN_FONT)
    rfonts.set(qn("w:hAnsi"), EN_FONT)
    style.paragraph_format.line_spacing = 1.5
    style.paragraph_format.space_before = Pt(0)
    style.paragraph_format.space_after = Pt(0)
    add_page_number(sec)
    return doc


# ---------------------------------------------------------------------------
# 封面
# ---------------------------------------------------------------------------
def add_cover(doc) -> None:
    for _ in range(2):
        add_para(doc, "", size=12, indent=False)
    add_para(doc, "暨南大学", size=26, bold=True, indent=False,
             align=WD_ALIGN_PARAGRAPH.CENTER, cn=CN_HEAD, space_before=6)
    add_para(doc, "案例实务课程实验报告", size=30, bold=True, indent=False,
             align=WD_ALIGN_PARAGRAPH.CENTER, cn=CN_HEAD, space_before=10, space_after=6)
    add_para(doc, "—— 量化因子复刻与多因子回测 ——", size=14, indent=False,
             align=WD_ALIGN_PARAGRAPH.CENTER, space_after=18)

    add_para(doc, "K线形态信息中的alpha：CSP因子复刻与多因子实证",
             size=18, bold=True, indent=False, align=WD_ALIGN_PARAGRAPH.CENTER,
             cn=CN_HEAD, space_before=10, space_after=24)

    items = [
        ("课程名称", "案例实务"),
        ("小组编号", "2 人组"),
        ("成员一", "叶承恩　　学号：202630196102"),
        ("成员二", "＿＿＿＿　　学号：＿＿＿＿＿＿＿＿"),
        ("学    院", "信息科学技术学院"),
        ("专    业", "应用统计"),
        ("指导教师", "肖老师"),
        ("报告标题", "K线形态信息中的alpha：CSP因子复刻与多因子实证"),
        ("所选研报", "《K线形态信息中的alpha》"),
        ("研报作者", "缪铃凯、刘富兵（国盛证券金融工程研究）"),
        ("发布时间", "2024 年 9 月 7 日；微信公众号“留富兵法”2024 年 9 月 10 日推送"),
        ("代码仓库", "GitHub（提交前替换为个人仓库链接）"),
        ("完成日期", "2026 年 11 月"),
    ]
    table = doc.add_table(rows=len(items), cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    for i, (k, v) in enumerate(items):
        c0, c1 = table.cell(i, 0), table.cell(i, 1)
        c0.width = Cm(3.4)
        c1.width = Cm(11.0)
        for c in (c0, c1):
            c.paragraphs[0].paragraph_format.space_before = Pt(3)
            c.paragraphs[0].paragraph_format.space_after = Pt(3)
        p0 = c0.paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r0 = p0.add_run(k + "：")
        set_run_font(r0, size=12, bold=True)
        p1 = c1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r1 = p1.add_run(v)
        set_run_font(r1, size=12)
    # 封面信息表默认无边框，无需额外设置
    doc.add_page_break()


# ---------------------------------------------------------------------------
# 小组分工（前置页）
# ---------------------------------------------------------------------------
def add_division(doc) -> None:
    add_para(doc, "小组分工", size=18, bold=True, indent=False,
             align=WD_ALIGN_PARAGRAPH.CENTER, cn=CN_HEAD, space_before=10, space_after=10)
    add_para(doc, "本报告由 2 人小组合作完成，两人贡献各占 50%，具体分工如下。",
             size=12, indent=False, space_after=8)
    rows = pd.DataFrame([
        ["成员一", "叶承恩\n（202630196102）",
         "数据获取与清洗；K 线形态划分与 CSP 因子构建（24 类单 K 形态、41 472 种组合形态、"
         "10 个周期、按月滚动打分与复合因子）；因子 IC 与分组检验；报告第 2、3 章及 4.1—4.3 节；"
         "数据与因子模块的代码实现", "50%"],
        ["成员二", "＿＿＿＿\n（学号：＿＿＿＿＿＿＿＿）",
         "多因子模型设定（Size+BP 基准模型、加入 CSP 的研究模型、4 周平滑改进模型）；"
         "周度调仓回测与交易成本核算；Fama-MacBeth、因子相关性、经典形态事件研究与"
         "参数敏感性检验；图表制作、报告 4.4—4.9 节与第 5 章撰写及全文统稿；"
         "回测、检验与报告模块的代码实现", "50%"],
    ], columns=["成员", "姓名（学号）", "主要工作", "贡献"])
    add_table(doc, rows, "小组分工与贡献比例", font_size=9,
              col_widths=[1.6, 3.4, 9.4, 1.6])
    add_para(doc, "合作方式：每周同步一次进度；数据清洗与因子结果由成员二独立复核，回测、图表与"
                  "稳健性结果由成员一独立复核；报告分章撰写后交叉修改，最终由两人共同定稿。",
             size=10.5, indent=False, space_before=10, line=1.3)
    add_para(doc, "两人共同完成选题与复刻方案设计、结果讨论和全文校对，均能独立解释报告的"
                  "主要内容和代码实现。",
             size=10.5, indent=False, space_before=6, line=1.3)
    doc.add_page_break()


# ---------------------------------------------------------------------------
# 摘要
# ---------------------------------------------------------------------------
def add_abstract(doc) -> None:
    add_para(doc, "摘  要", size=16, bold=True, indent=False,
             align=WD_ALIGN_PARAGRAPH.CENTER, cn=CN_HEAD, space_before=6, space_after=10)

    abstract = (
        "本文以国盛证券金融工程团队 2024 年 9 月发布的《K线形态信息中的alpha》为复刻对象，"
        "基于课程提供的 2020 年 1 月至 2025 年 12 月全 A 股日频行情数据，完整重构了报告提出的"
        "K 线形态划分与形态选股因子 CSP（Candlestick Pattern）。研究首先以实体宽度、上影线、"
        "下影线三个变量把单根 K 线划分为 24 种形态，再将连续 3 根 K 线与成交量状态组合为 41 472 种"
        "组合形态；随后以未来 20 日相对全市场等权收益的超额收益为标签，按月滚动估计每种形态的"
        "历史信息比率并映射为个股得分，在 1 日至 10 日共 10 个周期上构建 CSP 单周期因子与复合因子。"
        "实证结果表明：复合 CSP 的月度 Rank IC 均值为 6.31%，ICIR 为 0.83，IC 为正的比例为 83.1%；"
        "十分组年化收益从 G1 的 −9.46% 单调上升至 G10 的 27.18%，单调性秩相关系数为 0.988，"
        "多空组合年化收益 37.54%，Sharpe 比率 2.93；Fama-MacBeth 回归中控制市值、BP、动量、反转与"
        "波动率后，复合 CSP 的系数为 0.104（Newey-West t 值 7.65）。因子与常用量价因子的截面相关性"
        "绝对值均低于 0.12，独立性良好。在多因子层面，本文按课程要求构建仅含 Size 与 BP 的基准模型，"
        "并将复合 CSP 加入形成研究模型。结果显示：研究模型的毛年化收益（35.35%）高于基准模型"
        "（33.53%），但因形态信号切换频繁、周度换手率高达 35.65%，扣除单边 0.15% 的交易成本后净收益"
        "降至 28.07%，低于基准模型的 32.19%。对 CSP 信号做 4 周平滑后，换手率降至 14.77%，净收益回升"
        "至 31.87%，Sharpe 比率（1.20）与最大回撤（−25.62%）均略优于基准模型。本文认为，CSP 因子"
        "在统计意义上复刻成功，其 alpha 真实存在但衰减较快，交易成本与换手控制是形态类因子能否"
        "落地为组合收益的关键约束。"
    )
    add_para(doc, abstract, size=12, indent=True, line=1.5)
    add_para(doc, "关键词：K 线形态；CSP 因子；多因子模型；IC 分析；交易成本；周度回测",
             size=12, indent=False, bold=True, space_before=8)
    doc.add_page_break()


# ---------------------------------------------------------------------------
# 一、引言
# ---------------------------------------------------------------------------
def add_introduction(doc) -> None:
    add_heading(doc, "一、引言", 1)
    add_para(doc, "技术分析在 A 股市场拥有广泛的参与者基础，但其有效性长期存在争议。"
                  "经典文献如 Brock、Lakonishok 与 LeBaron（1992）以及 Lo、Mamaysky 与 Wang（2000）"
                  "均指出，价格序列中某些可识别的局部形态对未来收益具有一定预测力，但这种预测力"
                  "是否稳定、能否覆盖交易成本，始终缺乏统一结论。早期技术分析依赖交易员的经验与"
                  "主观识别，难以规模化检验；随着行情数据与计算能力的普及，把“形态”转化为可计算、"
                  "可回测的因子，成为量化研究的一条可行路径。")
    add_para(doc, "国盛证券金融工程团队 2024 年 9 月发布的《K线形态信息中的alpha》正是这一方向的"
                  "代表性研究报告。该报告不满足于检验十字星、锤头线等少数经典形态，而是把 K 线拆解为"
                  "实体与影线的组合，将单根 K 线划分为 24 类，再以连续 3 根 K 线与成交量状态组合出"
                  "41 472 种形态模式，并利用形态的历史绩效构建从“形态”到“得分”的映射，最终形成"
                  "选股因子 CSP。报告称 2014 年以来复合 CSP 在全市场的月度 IC 均值达到 9.6%，"
                  "ICIR 为 4.13，多头组合年化超额收益超过 11.5%，并把同一套形态信息应用到行业轮动，"
                  "取得 10.2% 的年化超额收益。")
    add_para(doc, "本文的工作是在课程提供的数据条件下对该报告进行“可复现的简化复刻”，并完成课程"
                  "要求的基准模型—研究模型对比回测。具体目标有三：第一，严格按照报告的定义重建 24 类"
                  "单 K 形态、41 472 种组合形态、多周期 CSP 与复合 CSP，检验其在 2020—2025 年 A 股"
                  "样本上的 IC、分组收益与显著性；第二，按照课程统一要求，以周度数据、全 A 股标的、"
                  "固定的调仓频率，构建仅含 Size 与 BP 的基准模型，并把 CSP 加入形成研究模型，与沪深 300"
                  "指数进行同期对比；第三，在复刻基础上进行反思，重点回答“统计上显著的形态因子为何"
                  "未必带来组合层面的超额收益”，并给出可操作的改进方向。")
    add_para(doc, "与简单复现相比，本文的边际贡献体现在三个方面。其一，完整实现了 10 个周期的形态"
                  "构建与按月滚动重估，全部使用历史信息估计形态得分，避免未来函数，并给出因子覆盖率、"
                  "有效形态数与样本量的完整诊断。其二，把因子检验与组合回测放在同一套数据、同一套"
                  "预处理与同一套成本假设下比较，使“因子有效”与“组合有效”两个层次的差异可以被"
                  "清晰分解。其三，通过成本归因、因子权重敏感性、信号平滑与换手缓冲等实验，说明"
                  "形态类因子的收益—换手权衡，这一结论对同类高频量价因子的落地具有一般性。")
    add_para(doc, "全文结构如下：第二部分介绍原报告框架、复刻因子的选择与数学定义；第三部分说明"
                  "数据来源、样本筛选、因子构建与回测设定；第四部分报告因子检验与多因子回测结果，"
                  "并与原报告结论逐项对比；第五部分总结全文、指出局限并展望后续改进方向。")


# ---------------------------------------------------------------------------
# 二、原报告与因子说明
# ---------------------------------------------------------------------------
def add_factor_section(doc) -> None:
    add_heading(doc, "二、原报告与因子说明", 1)
    add_heading(doc, "2.1 原报告的研究框架与核心结论", 2)
    add_para(doc, "原报告的出发点是把“看图”变成“算图”。它首先用实体宽度、上影线长度与下影线长度"
                  "三个变量刻画任意一根 K 线，并给出十字星、锤头线、倒锤头线的定量定义；在此基础上"
                  "按实体大小与涨跌方向把单根 K 线细分为 2×3×2×2 共 24 类。随后，报告用连续 3 根同周期"
                  "K 线构成组合形态，并叠加成交量状态（放量、常量、缩量），得到 24³×3＝41 472 种组合；"
                  "进一步把日线聚合为 2 日线、3 日线直至 10 日线，从 10 个频率视角重复上述划分，得到"
                  "10 个单周期形态因子。")
    add_para(doc, "因子构造的关键一步是“形态到得分”的映射。报告以股票未来 20 日相对全市场等权收益的"
                  "超额收益作为标签，统计每种形态历史上该超额收益的均值与标准差，用信息比率 IR 作为"
                  "形态得分；出现某形态的股票即取得该形态的历史得分。由于形态数量远大于样本量，报告"
                  "还从多周期视角与多维度评价准则（相对全市场、相对行业、剥离风格残差）对得分进行"
                  "综合，形成复合 CSP 因子。")
    add_para(doc, "原报告的核心结论可概括为四点：一是形态划分具有显著的收益预测能力，2016 年以来"
                  "看多样本月均约 580 个、未来 20 日超额收益均值 2.5%，看空样本均值 −2.2%；二是形态"
                  "绩效具有持续性，2012—2023 年各形态前后两年的绩效相关系数均值为 25%；三是复合 CSP"
                  "因子 2014 年以来月度 IC 均值 9.6%、ICIR 4.13，多空年化收益 32.3%，多头年化超额收益"
                  "11.5%；四是形态信息可迁移至行业轮动，交易型行业信号与基本面信号相关性接近 0，"
                  "二者复合后行业多头组合年化超额收益 10.2%。")
    add_para(doc, "需要说明的是，原报告的样本自 2010 年起、正式回测自 2014 年起，数据跨度超过十年；"
                  "课程提供的数据自 2020 年 1 月开始，至 2025 年 12 月结束，且不包含行业指数、分钟"
                  "行情与财务数据。因此本文的复刻在样本区间、数据频率与因子维度上均做了合理简化，"
                  "这些差异也是后文结果对比时需要重点考虑的因素。")

    add_heading(doc, "2.2 复刻因子的选择", 2)
    add_para(doc, "按照课程“从报告中选择 1—3 个核心因子”的要求，结合数据可得性，本文选择以下三个"
                  "核心因子作为复刻对象。")
    add_para(doc, "（1）CSP_5，周线单周期形态因子。把 5 个交易日聚合为一根周 K 线，按 24 类单 K 形态与"
                  "3 根组合的方式构造 41 472 种模式，是原报告中最基础、也最容易与经典技术分析对话的"
                  "版本。选择它，是为了检验“周线形态”这一报告主结论在较短样本上的稳健性。")
    add_para(doc, "（2）复合 CSP，多周期综合因子。对 1 日至 10 日共 10 个周期分别构造 CSP，并在横截面上"
                  "标准化后等权平均。它是原报告绩效最好的版本（月度 IC 9.6%），也是本文研究模型所"
                  "采用的因子。")
    add_para(doc, "（3）经典形态事件因子。直接按报告对十字星、锤头线、倒锤头线的定义，叠加“价格位于"
                  "一年高点/低点附近”与“当日放量”两个条件，构造事件型打分。它对应报告 2010 年以来"
                  "的经典形态事件研究，用于检验报告关于“顶部反转形态跑输、底部反转形态跑赢”的判断。")

    add_heading(doc, "2.3 因子的数学定义", 2)
    add_para(doc, "设股票在 t 日的开盘、最高、最低、收盘价分别为 O_t、H_t、L_t、C_t，前一交易日收盘价为"
                  "C_{t−1}。记实体宽度为 w_t、上影线为 u_t、下影线为 d_t，三者定义为：")
    add_formula(doc, r"w_t=\frac{|C_t-O_t|}{C_{t-1}}", 1,
                plain="实体宽度_t = |C_t − O_t| / C_{t−1}")
    add_formula(doc, r"u_t=\frac{H_t-\mathrm{max}(O_t,C_t)}{C_{t-1}}", 2,
                plain="上影线_t = [H_t − max(O_t, C_t)] / C_{t−1}")
    add_formula(doc, r"d_t=\frac{\mathrm{min}(O_t,C_t)-L_t}{C_{t-1}}", 3,
                plain="下影线_t = [min(O_t, C_t) − L_t] / C_{t−1}")
    add_para(doc, "其中 w_t、u_t、d_t 分别对应实体宽度、上影线与下影线，均以百分比表示，因而与股价"
                  "绝对水平无关。按原报告阈值：w_t 小于 1.5% 记为小实体、1.5%—3% 记为中实体、大于 3% "
                  "记为大实体；u_t 或 d_t 不小于 2% 记为长影线，否则为短影线；收盘价不低于开盘价记为阳线，"
                  "否则为阴线。三者组合得到 24 类单 K 形态，用整数 b_t∈{0,…,23} 表示：")
    add_formula(doc, r"b_t=12\cdot\mathrm{I}(C_t\ge O_t)+4\cdot g(w_t)+2\cdot\mathrm{I}(u_t\ge 2\%)+\mathrm{I}(d_t\ge 2\%)", 4,
                plain="b_t = 12·I(C_t ≥ O_t) + 4·g(w_t) + 2·I(u_t ≥ 2%) + I(d_t ≥ 2%)")
    add_para(doc, "其中 g(·)∈{0,1,2} 为实体大小的三档取值，I(·) 为示性函数。对连续 3 根同周期 K 线，"
                  "再结合最新一根的成交量状态 v_t∈{0,1,2}（缩量、常量、放量，以该根 K 线成交额相对"
                  "过去 20 根同类 K 线成交额均值是否低于 0.8 倍、高于 1.2 倍判定），组合形态编码为：")
    add_formula(doc, r"P_t=3\cdot\left\{24\cdot\left[24\cdot b_{t-2}+b_{t-1}\right]+b_t\right\}+v_t", 5,
                plain="P_t = 3·{24·[24·b_{t−2} + b_{t−1}] + b_t} + v_t")
    add_para(doc, "P_t 取值 0—41 471，共 41 472 种组合。设形态 p 在历史样本中对应的未来 20 日超额收益"
                  "为 r_i（i=1,…,n_p），则形态得分为：")
    add_formula(doc, r"S_p=\frac{\mathrm{mean}(r_i)}{\mathrm{std}(r_i)}\cdot\frac{n_p}{n_p+\kappa}", 6,
                plain="S_p = [mean(r_i) / std(r_i)] · n_p / (n_p + κ)")
    add_para(doc, "式（6）中的第二项是本文引入的小样本收缩因子，取 κ=50。原报告直接用 IR 作为得分，"
                  "但其样本跨越十年，单个形态的样本量相对充足；本文样本仅六年，若不加收缩，样本量"
                  "只有二三十个的形态会给出极端的 IR，反而引入噪声。收缩后，样本越少的形态得分越接近 0，"
                  "相当于向“无信息”收缩，这是一种经验贝叶斯式的稳健化处理。因子值定义为股票当前形态"
                  "的得分：")
    add_formula(doc, r"CSP_{i,t}=S_{P_{i,t}}", 7,
                plain="CSP_{i,t} = S_{P_{i,t}}")
    add_para(doc, "其中 i 表示第 i 个周期（i=1,…,10）。复合因子在横截面上对每个单周期因子做标准化，"
                  "再等权平均：")
    add_formula(doc, r"CSP_t=\frac{1}{10}\sum_{i=1}^{10} z(CSP_{i,t})", 8,
                plain="CSP_t = (1/10) · Σ_{i=1}^{10} z(CSP_{i,t})")
    add_para(doc, "经典形态事件因子则按日频识别三类形态，并叠加位置与成交量条件。以看多事件为例：")
    add_formula(doc, r"E_t=\sum_{j=0}^{19}\left[I^{+}_{t-j}-I^{-}_{t-j}\right]", 9,
                plain="E_t = Σ_{j=0}^{19} [I⁺_{t−j} − I⁻_{t−j}]")
    add_para(doc, "其中 I⁺(s) 与 I⁻(s) 分别为 s 日的看多、看空事件示性函数：当 s 日处于底部（顶部）、"
                  "出现十字星或锤头线（倒锤头线）且放量时取 1，否则取 0；“顶部”指收盘价不低于过去一年"
                  "最高价的 90%，“底部”指收盘价不高于过去一年最低价的 110%，“放量”指当日成交额不低于"
                  "过去 20 日均值的 1.2 倍。该因子把最近 20 个交易日的形态事件净计数作为得分，"
                  "用于检验经典形态的短期预测能力。")

    add_heading(doc, "2.4 与原报告的口径差异", 2)
    add_para(doc, "为保证复刻的可比性，本文在因子定义上尽量贴近原文，但仍存在若干不可避免的差异，"
                  "需要在解读结果时予以说明。第一，样本区间不同：原文为 2010—2024 年，本文为 2020—2025 年，"
                  "后者经历了 2021 年小市值价值股的极端行情与 2024 年初的微盘股回撤，风格环境差异显著。"
                  "第二，原文使用 20 日超额收益作为标签，并在多维度上评价形态（相对全市场、相对行业、"
                  "剥离风格），本文只保留“相对全市场等权”这一维度，因子维度更少，理论上会削弱复合因子的"
                  "稳定性。第三，原文未披露形态得分是否做小样本处理，本文加入了收缩因子，会系统性压低"
                  "稀疏形态的得分，从而降低 IC 的极值但提高稳健性。第四，原文的回测包含月度调仓与更"
                  "精细的交易成本假设，本文按课程统一要求采用周度调仓与统一的单边 0.15% 成本，换手率"
                  "对结果的影响被放大。上述差异不改变复刻的逻辑，但会影响数值层面的可比性。")


# ---------------------------------------------------------------------------
# 三、数据与方法
# ---------------------------------------------------------------------------
def add_data_section(doc) -> None:
    add_heading(doc, "三、数据与方法", 1)
    add_heading(doc, "3.1 数据来源、样本区间与标的范围", 2)
    add_para(doc, "本文使用课程提供的 A 股日频行情数据集，原始文件按字段与市场分别存储为宽表，"
                  "第一列为交易日，其余列为股票代码。涉及字段包括开盘价、最高价、最低价、收盘价、"
                  "成交量、成交额、复权因子、总市值、市净率、市盈率（TTM）与 ST 标记，另有沪深 300 指数"
                  "行情与申万一级行业分类表。数据区间为 2020 年 1 月 2 日至 2025 年 12 月 31 日，"
                  "共 1 455 个交易日；合并沪深两市后得到 5 344 只股票，平均每个交易日有 4 805 只股票"
                  "存在有效收盘价。")
    add_para(doc, "标的范围设定为全 A 股，并施加以下剔除规则：第一，剔除调仓日处于 ST/*ST 状态的股票，"
                  "这类股票涨跌停制度与退市风险会显著扭曲形态信号；第二，剔除上市未满 250 个交易日的"
                  "新股，避免上市初期的极端波动与形态样本不足；第三，剔除调仓日停牌（无成交额）或"
                  "收盘价缺失的股票；第四，构造 BP 因子时要求市净率为正，构造 Size 因子时要求总市值"
                  "为正。经过筛选，每个调仓日平均约有 4 450 只股票进入候选池，按 10% 分位选股时组合"
                  "平均持有 445 只股票。")

    add_heading(doc, "3.2 数据清洗与预处理", 2)
    add_para(doc, "原始行情存在少量非正价格、零成交额与 OHLC 逻辑不一致的记录。本文的处理方式是："
                  "把收盘价、开盘价、最高价、最低价中的非正值统一置为缺失；把成交额或成交量为零的"
                  "记录视为停牌；要求最高价不低于最低价，否则整根 K 线作废。价格序列使用复权因子"
                  "还原为后复权价（价格×复权因子），所有收益与形态计算均在复权价上进行，以避免除权"
                  "除息造成的虚假跳空。")
    add_para(doc, "在因子层面，本文对 Size、BP 与 CSP 均做横截面去极值与标准化：先按 1% 与 99% 分位"
                  "缩尾，再减去截面均值、除以截面标准差；随后以申万一级行业虚拟变量做行业中性化，"
                  "即减去行业内均值，再次标准化。行业中性化的目的是避免形态因子在特定行业上的集中"
                  "暴露，使因子比较更接近“同行业内选股”的实际应用场景。需要说明的是，行业分类表为"
                  "当前时点的静态分类，未做历史成分回溯，这是本文数据条件下的一个局限。")

    add_heading(doc, "3.3 形态构建与因子计算流程", 2)
    add_para(doc, "因子构建按以下五步完成。第一步，按周期 k 把日线聚合为 k 日 K 线：第 j 根 K 线的开盘价"
                  "取区间首日开盘价，收盘价取区间末日收盘价，最高（低）价取区间内最高（低）价，成交额"
                  "取区间合计；k 日 K 线采用不重叠切分，因而每个周期只在每 k 个交易日产生一根新 K 线。"
                  "第二步，计算实体、上影线、下影线与成交量状态，按式（5）生成 41 472 种组合形态编码。"
                  "第三步，以未来 20 个交易日相对全市场等权收益的超额收益作为标签；为与“预测未来一个月”"
                  "的设定一致，标签在 K 线结束日计算。第四步，按月滚动估计形态得分：对每个月末，只用"
                  "结束日不晚于该月末前 20 个交易日的样本（这些样本的未来收益已经实现），按式（6）计算"
                  "每个形态的得分，因此不会使用未来信息；每个月的因子值使用上月末估计的得分表。"
                  "第五步，把 K 线结束日的得分向前填充到之后的交易日，得到日频因子值，再在 10 个周期上"
                  "合成复合因子。")
    add_para(doc, "这一流程的关键在于“得分表按月重估”。如果直接用全样本估计形态得分再回测，"
                  "等价于让因子提前知道了未来哪些形态会有效，会产生严重的前视偏差。本文的做法是"
                  "扩张窗口：随着时间推移，用于估计的样本不断增加，但始终截止于当前月末之前，"
                  "因此每个月的因子值都是“用历史信息估计、用于预测未来”的真实样本外结果。")
    add_para(doc, "从数据量看，日线周期共有 1 455 根 K 线、约 683 万个有效的股票—日观测；5 日周期有 291 根"
                  "K 线、约 135 万个组合形态样本；10 日周期有 145 根 K 线、约 66 万个样本。全部 10 个周期"
                  "合计约 1 978 万个组合形态样本。由于组合形态多达 41 472 种，样本分布高度不均：以 5 日"
                  "周期为例，出现过的形态有 39 942 种，但单个形态的样本中位数仅 15 个，57.3% 的形态"
                  "样本数不足 20 个（这些稀疏形态只覆盖了 13.8% 的样本）。这正是本文引入收缩因子的"
                  "现实原因：若直接使用 IR，样本量只有个位数的形态会给出极端得分，反而引入噪声。")

    add_heading(doc, "3.4 基准模型与研究模型的设定", 2)
    add_para(doc, "按照课程要求，本文以打分法构建多因子模型，所有因子先做行业中性化与标准化，"
                  "再按等权合成综合得分。基准模型只使用市值因子（Size）与账面市值比因子（BP）：")
    add_formula(doc, r"\mathrm{Score}_{BM}=0.5\cdot(-z_{\mathrm{Size}})+0.5\cdot z_{\mathrm{BP}}", 10,
                plain="Score_BM = 0.5 · (−z_Size) + 0.5 · z_BP")
    add_para(doc, "其中 z_Size 为对数总市值的截面标准化值，取负号表示偏好小市值；z_BP 为市净率倒数"
                  "（1/PB）的标准化值，取值越大代表账面市值比越高、估值越便宜。研究模型在基准模型"
                  "基础上加入复合 CSP：")
    add_formula(doc, r"\mathrm{Score}_{RM}=\frac{1}{3}\cdot(-z_{\mathrm{Size}})+\frac{1}{3}\cdot z_{\mathrm{BP}}+\frac{1}{3}\cdot z_{\mathrm{CSP}}", 11,
                plain="Score_RM = (1/3) · (−z_Size) + (1/3) · z_BP + (1/3) · z_CSP")
    add_para(doc, "为检验“加入核心因子是否改善组合”，本文还构造了三个对照：以 CSP_5 替代复合 CSP 的"
                  "研究模型、以经典形态事件因子替代复合 CSP 的研究模型，以及对复合 CSP 做 4 周平滑"
                  "后的改进模型。平滑的处理方式是对过去 4 周的 z_CSP 取简单平均，其经济含义是"
                  "“不因为一周的形态变化就调仓”，以降低换手。")

    add_heading(doc, "3.5 回测框架与评价指标", 2)
    add_para(doc, "回测严格遵循课程统一要求：使用周度数据，以每周最后一个交易日的收盘价计算周收益率；"
                  "调仓频率设为每周一次（基准情形），并在稳健性检验中测试双周与四周调仓；标的范围为"
                  "全 A 股，剔除规则见 3.1 节；组合内股票等权，选取综合得分最高的 10% 股票，"
                  "组合平均持有约 445 只。交易成本按单边 0.15% 计，换手率定义为调仓时被替换的股票"
                  "占新组合的比例，一次调仓的成本为 2×0.15%×换手率。回测区间为 2021 年 1 月 15 日至"
                  "2025 年 12 月 19 日，共 253 个周度观测；2020 年的数据用于形态得分的初始估计，"
                  "不参与业绩评价，因此首个周度收益从因子可用后的 2021 年 1 月中旬开始。")
    add_para(doc, "评价指标包括：年化收益率（按 52 周复利折算）、年化波动率、Sharpe 比率（无风险利率"
                  "取 2%）、最大回撤、Calmar 比率、周胜率、相对沪深 300 的年化超额收益、跟踪误差与"
                  "信息比率。因子层面则使用 Rank IC 均值、IC 标准差、ICIR、IC 为正比例、十分组年化收益、"
                  "多空组合绩效、Fama-MacBeth 回归系数与 Newey-West t 值，以及因子间截面相关性。"
                  "其中 IC 与 Fama-MacBeth 回归均采用周度频率，并按 20 日预测窗口计算，"
                  "Newey-West 滞后阶数取 4 期以修正重叠样本带来的自相关。")


# ---------------------------------------------------------------------------
# 四、实证结果
# ---------------------------------------------------------------------------
def add_results_section(doc) -> None:
    add_heading(doc, "四、实证结果", 1)

    # 4.1 样本与形态分布
    add_heading(doc, "4.1 样本概况与形态分布", 2)
    add_para(doc, "图 1 给出样本期内每日有效交易股票数与复合 CSP 因子覆盖股票数的变化。2020 年初"
                  "有效股票约 3 760 只，此后随注册制推进逐步上升至 2025 年的 5 300 余只；因子覆盖股票数"
                  "在 2021 年后稳定在 4 000 只以上，全样本覆盖率为 82.5%。覆盖率未达 100% 的原因是"
                  "部分股票上市时间较短、形态样本不足，或某些组合形态在历史样本中从未出现，"
                  "其得分按缺失处理。")
    add_figure(doc, "fig01_样本覆盖.png",
               "图 1  样本期内有效交易股票数与复合 CSP 因子覆盖股票数")
    dist = T("t15_单K线形态分布.csv")
    top3 = dist.nlargest(3, "share")
    add_para(doc, f"图 2 统计了单根日 K 线 24 种形态的样本占比。分布高度不均：小实体、短上下影的"
                  f"“小阳线”与“小阴线”合计占比约 42.6%，是市场中最常见的形态；而同时具备大实体、"
                  f"长上影与长下影的极端形态合计占比不足 1%。这与 A 股日常波动以窄幅震荡为主的特征"
                  f"一致，也意味着大多数交易日的形态信息量有限，真正有区分度的是少数极端形态。")
    add_figure(doc, "fig02_单K线形态分布.png",
               "图 2  单根日 K 线 24 种形态的样本占比")
    data_rows = pd.DataFrame([
        ["股票数量", f"{R['数据']['股票数']:,} 只"],
        ["交易日数量", f"{R['数据']['交易日数']:,} 个"],
        ["样本区间", "2020-01-02 至 2025-12-31"],
        ["平均每日有效股票数", f"{R['数据']['平均每日有效股票数']:,.0f} 只"],
        ["回测区间", f"{R['回测区间']['起始']} 至 {R['回测区间']['结束']}"],
        ["周度调仓次数", "253 次"],
        ["平均候选股票数", "约 4 450 只"],
        ["组合平均持仓数", f"{perf[BM]['平均持仓数']:.0f} 只"],
        ["复合 CSP 覆盖率", "82.5%"],
        ["单期有效形态数（1 日 / 10 日）", "19 311 / 7 385"],
    ], columns=["项目", "数值"])
    add_table(doc, data_rows, "表 1  数据与样本概况", font_size=10,
              col_widths=[7.0, 8.0])

    # 4.2 IC
    add_heading(doc, "4.2 因子有效性：IC 分析", 2)
    comp_m = monthly_ic["CSP_复合"]
    add_para(doc, f"表 2 汇总了四个因子在不同预测窗口下的周度 Rank IC。复合 CSP 的表现显著优于任一"
                  f"单周期因子：在未来 20 日窗口上，其 IC 均值为 {pct(weekly_ic[('CSP_复合', '未来20日')]['IC均值'])}，"
                  f"ICIR 为 {num(weekly_ic[('CSP_复合', '未来20日')]['ICIR'])}，IC 为正的比例为 "
                  f"{pct(weekly_ic[('CSP_复合', '未来20日')]['IC为正比例'], 1)}；而单周期的 CSP_1、CSP_5、"
                  f"CSP_10 的 IC 均值分别为 {pct(weekly_ic[('CSP_1', '未来20日')]['IC均值'])}、"
                  f"{pct(weekly_ic[('CSP_5', '未来20日')]['IC均值'])} 与 "
                  f"{pct(weekly_ic[('CSP_10', '未来20日')]['IC均值'])}。这说明多周期综合确实能够过滤"
                  f"单一周期形态的噪声，与原报告“复合因子显著优于简单版”的结论方向一致。")
    ic_show = []
    for f_ in ("CSP_1", "CSP_5", "CSP_10", "CSP_复合"):
        for h in ("未来5日", "未来10日", "未来20日", "未来60日"):
            row = weekly_ic[(f_, h)]
            ic_show.append([f_, h, pct(row["IC均值"]), pct(row["IC标准差"]),
                            num(row["ICIR"], 2), pct(row["IC为正比例"], 1), num(row["t值"], 2)])
    add_table(doc, pd.DataFrame(ic_show, columns=["因子", "预测窗口", "IC 均值", "IC 标准差",
                                                  "ICIR", "IC 为正比例", "t 值"]),
              "表 2  CSP 因子的周度 Rank IC（预测窗口 5/10/20/60 日）", font_size=8.5,
              col_widths=[2.2, 2.4, 2.6, 2.6, 1.9, 2.6, 1.9])
    add_para(doc, f"按月频计算，复合 CSP 的月度 IC 均值为 {pct(comp_m['IC均值'])}，IC 标准差为 "
                  f"{pct(comp_m['IC标准差'])}，ICIR 为 {num(comp_m['ICIR'])}，IC 为正的比例为 "
                  f"{pct(comp_m['IC为正比例'], 1)}，t 值为 {num(comp_m['t值'], 2)}。图 3 的累积 IC 曲线"
                  f"呈持续上行态势，59 个月中仅有 10 个月为负，说明因子的预测能力具有较强的持续性，"
                  f"并非由个别月份驱动。")
    add_figure(doc, "fig03_IC分析.png", "图 3  复合 CSP 与周线 CSP 的月度 IC 序列及累积 IC")
    add_para(doc, f"与原报告相比，本文的 IC 强度明显偏低：原报告复合 CSP 的月度 IC 为 9.6%、ICIR 为 4.13，"
                  f"本文对应数值为 {pct(comp_m['IC均值'])} 与 {num(comp_m['ICIR'])}。强度差异有三方面原因。"
                  f"其一，样本长度不同。原报告有超过十年的月度样本，形态得分的估计误差被时间平均掉；"
                  f"本文只有 59 个月，且前 12 个月的训练样本不足，早期因子噪声较大。其二，本文对形态得分"
                  f"做了收缩，稀疏形态的得分被系统性压低，IC 的绝对值随之下降，但 ICIR 的稳定性反而提高。"
                  f"其三，原报告的复合因子融合了相对行业、剥离风格残差等多个维度的标签，本文只有"
                  f"相对全市场一个维度，信息聚合程度较低。因此，本文结果应理解为“方向性复刻成功、"
                  f"强度有所衰减”，而不是对原报告数值的直接否定。")

    # 4.3 分组
    add_heading(doc, "4.3 分组收益与多空组合", 2)
    grp = T("t04_分组收益.csv")
    comp = grp[grp["因子"] == "CSP_复合"].reset_index(drop=True)
    g1, g10 = comp.iloc[0], comp.iloc[-1]
    lsc = ls["CSP_复合"]
    add_para(doc, f"表 3 报告复合 CSP 的十分组结果。G1 至 G10 的年化收益从 {pct(g1['年化收益'])} 单调上升至"
                  f"{pct(g10['年化收益'])}，组序与年化收益的 Spearman 秩相关系数达到 "
                  f"{num(R['组内单调性']['CSP_复合'])}，单调性近乎完美。多空组合（G10−G1）年化收益 "
                  f"{pct(lsc['年化收益'])}，年化波动 {pct(lsc['年化波动'])}，Sharpe 比率 {num(lsc['夏普比率'])}，"
                  f"最大回撤仅 {pct(lsc['最大回撤'])}，周胜率 {pct(lsc['周胜率'], 1)}，t 值 {num(lsc['t值'])}。"
                  f"这一结果说明形态因子在横截面上具有稳定的区分能力。")
    grp_show = []
    for i in range(10):
        c5 = grp[(grp["因子"] == "CSP_5")].reset_index(drop=True).iloc[i]
        cc = comp.iloc[i]
        grp_show.append([f"G{i+1}", pct(c5["年化收益"]), pct(cc["年化收益"]),
                         pct(c5["年化波动"]), pct(cc["年化波动"])])
    add_table(doc, pd.DataFrame(grp_show, columns=["分组", "CSP_5 年化收益", "复合 CSP 年化收益",
                                                   "CSP_5 波动", "复合 CSP 波动"]),
              "表 3  因子十分组的年化收益与波动（周度调仓、等权）", font_size=9,
              col_widths=[2.0, 3.4, 3.4, 3.0, 3.0])
    add_para(doc, f"值得注意的是，多空收益的来源并不对称：G1 的年化收益为 {pct(g1['年化收益'])}，"
                  f"而 G10 为 {pct(g10['年化收益'])}，空头端的负收益幅度更大。换言之，CSP 更擅长"
                  f"“剔除坏股票”，而不是“选出最好的股票”。这一点对后文的多因子组合非常关键："
                  f"当基准组合已经通过小市值与高 BP 选出了收益更高的多头时，再加入一个“多头端较弱、"
                  f"空头端较强”的因子，反而可能稀释组合收益。")
    add_figure(doc, "fig04_分组与多空.png", "图 4  因子十分组年化收益与多空组合累计净值")
    add_para(doc, f"把 G10 与同期全 A 等权收益比较，CSP 多头组合的年化收益为 {pct(g10['年化收益'])}，"
                  f"全 A 等权为 {pct(perf['全A等权']['年化收益率'])}，超额约 "
                  f"{pp((g10['年化收益'] - perf['全A等权']['年化收益率']) * 100)} 个百分点。这一数字与原报告"
                  f"“多头组合年化超额收益 11.5%”在量级上高度一致，说明本文复刻的因子在多头端同样"
                  f"取得了与经济含义相符的超额收益。")

    # 4.4 FMB 与相关性
    add_heading(doc, "4.4 Fama-MacBeth 回归与因子独立性", 2)
    fmb = pd.DataFrame(R["FamaMacBeth"])
    fmb_c = fmb[fmb["因子"] == "CSP_复合"].copy()
    csp_row = fmb_c[fmb_c["变量"] == "CSP"].iloc[0]
    add_para(doc, f"表 4 给出以周度下一期收益为因变量、以复合 CSP 及五个常用风格因子为自变量的"
                  f"Fama-MacBeth 回归结果。在控制市值、BP、动量、反转与波动率之后，CSP 的平均回归系数"
                  f"为 {num(csp_row['平均系数'])}，Newey-West t 值为 {num(csp_row['NW t值'], 2)}，"
                  f"在 1% 水平上显著；系数为正的期数占比为 {pct(csp_row['系数为正比例'], 1)}，"
                  f"平均 R² 为 {num(csp_row['平均R2'], 3)}。这说明 CSP 的预测能力并不能被常见风格因子"
                  f"完全解释，具有独立的增量信息。市值因子系数为负（小市值溢价），动量与反转系数为负，"
                  f"与 A 股短周期反转、小盘占优的特征一致。")
    fmb_show = []
    for _, r in fmb_c.iterrows():
        fmb_show.append([r["变量"], num(r["平均系数"], 4), num(r["标准差"], 4),
                         num(r["NW t值"], 2), pct(r["系数为正比例"], 1), num(r["平均R2"], 3)])
    add_table(doc, pd.DataFrame(fmb_show, columns=["变量", "平均系数", "系数标准差",
                                                   "Newey-West t 值", "系数为正比例", "平均 R²"]),
              "表 4  复合 CSP 的 Fama-MacBeth 回归结果（周度，253 期）", font_size=9,
              col_widths=[2.6, 2.8, 2.8, 3.4, 2.8, 2.2])
    add_para(doc, f"表 5 进一步报告 CSP 与常用量价因子的截面相关性。复合 CSP 与市值的平均相关系数仅为"
                  f"{num(corr['市值']['平均相关系数'])}，与 BP 为 {num(corr['BP']['平均相关系数'])}，"
                  f"与动量、反转、波动率、换手率、非流动性的相关系数绝对值均不超过 0.12。"
                  f"因子独立性良好，意味着把 CSP 加入 Size+BP 组合在信息上应当是互补的，"
                  f"这与后文组合层面未能改善收益形成了有趣的对照，也提示问题更可能出在交易成本"
                  f"与多头端结构上，而非信息冗余。")
    corr_show = [[k, num(v["平均相关系数"], 3), num(v["相关系数标准差"], 3)]
                 for k, v in corr.items()]
    add_table(doc, pd.DataFrame(corr_show, columns=["对照因子", "平均相关系数", "相关系数标准差"]),
              "表 5  CSP 因子与常用量价因子的截面相关性（251 期均值）", font_size=9.5,
              col_widths=[4.0, 4.0, 4.0])
    add_figure(doc, "fig09_因子相关性.png", "图 5  CSP 因子与常用量价因子的截面相关性")

    # 4.5 事件研究
    add_heading(doc, "4.5 经典 K 线形态的事件研究", 2)
    add_para(doc, f"表 6 报告四类经典形态事件后未来 20 日超额收益的统计结果。顶部反转形态的表现与"
                  f"原报告高度一致：顶部十字星叠加放量后，未来 20 日超额收益均值为 "
                  f"{pct(ev['顶部十字星+放量']['超额收益均值'])}，胜率仅 "
                  f"{pct(ev['顶部十字星+放量']['胜率'], 1)}，t 值 {num(ev['顶部十字星+放量']['t值'], 2)}；"
                  f"顶部倒锤头线叠加放量的超额收益为 {pct(ev['顶部倒锤头线+放量']['超额收益均值'])}，"
                  f"同样显著为负。这说明“高位放量出现反转形态后跑输市场”的结论在 2020—2025 年"
                  f"依然成立。")
    add_para(doc, f"底部形态则未能复现原报告的正超额。底部十字星叠加放量的超额收益为 "
                  f"{pct(ev['底部十字星+放量']['超额收益均值'])}，t 值仅 "
                  f"{num(ev['底部十字星+放量']['t值'], 2)}；底部锤头线叠加放量甚至为 "
                  f"{pct(ev['底部锤头线+放量']['超额收益均值'])}，t 值 {num(ev['底部锤头线+放量']['t值'], 2)}，"
                  f"显著为负。合并后的底部反转形态超额收益为 {pct(ev['底部反转形态(合并)']['超额收益均值'])}，"
                  f"胜率 {pct(ev['底部反转形态(合并)']['胜率'], 1)}。这一结果与原报告“底部反转形态大概率"
                  f"跑赢市场”的结论相反。")
    ev_show = []
    for k, v in ev.items():
        ev_show.append([k, f"{v['样本数']:,}", pct(v["超额收益均值"]),
                        pct(v["超额收益中位数"]), pct(v["胜率"], 1), num(v["t值"], 2)])
    add_table(doc, pd.DataFrame(ev_show, columns=["事件类型", "样本数", "超额收益均值",
                                                  "超额收益中位数", "胜率", "t 值"]),
              "表 6  经典 K 线形态事件后未来 20 日超额收益", font_size=8.5,
              col_widths=[4.6, 2.4, 2.8, 3.0, 2.0, 2.0])
    add_figure(doc, "fig08_经典形态事件.png", "图 6  经典形态事件后 20 日超额收益及分年度表现")
    add_para(doc, "对这一差异，本文给出三点解释。第一，样本期不同。原报告的底部形态统计覆盖 2010 年"
                  "以来的多个牛熊周期，而本文的 2020—2025 年包含了 2021 年小市值行情、2022—2024 年"
                  "的持续调整与 2024 年初的微盘股流动性危机；在下跌趋势中，“底部”往往只是下跌中继，"
                  "锤头线之后继续下跌并不罕见。第二，位置定义不同。本文用过去一年最高价/最低价的"
                  "90%/110% 界定顶部与底部，在单边下跌行情中会频繁触发“底部”信号，导致样本中混入"
                  "大量伪底部。第三，超额收益的基准是全市场等权收益，而 2021—2025 年小市值股票整体"
                  "跑赢，底部形态多出现在弱势股上，即使个股绝对收益为正，也可能跑输等权基准。"
                  "这一发现提醒我们，经典形态的“反转”含义具有强烈的状态依赖性，不能脱离市场环境"
                  "机械套用。")

    # 4.6 回测
    add_heading(doc, "4.6 多因子回测结果", 2)
    bm, rm, rmi = perf[BM], perf[RM], perf[RMI]
    add_para(doc, f"表 7 与图 7 给出基准模型、研究模型与沪深 300 的净值表现。基准模型在 2021—2025 年"
                  f"取得了 {pct(bm['年化收益率'])} 的年化收益，年化波动 {pct(bm['年化波动率'])}，"
                  f"Sharpe 比率 {num(bm['夏普比率'])}，最大回撤 {pct(bm['最大回撤'])}；加入复合 CSP 后，"
                  f"研究模型的年化收益降至 {pct(rm['年化收益率'])}，Sharpe 比率降至 {num(rm['夏普比率'])}，"
                  f"最大回撤扩大至 {pct(rm['最大回撤'])}。同期沪深 300 的年化收益仅 {pct(perf['沪深300']['年化收益率'])}，"
                  f"最大回撤 {pct(perf['沪深300']['最大回撤'])}，全 A 等权年化收益 {pct(perf['全A等权']['年化收益率'])}。"
                  f"三个组合均大幅跑赢指数，但研究模型相对基准模型的年化超额为 "
                  f"{pct(R['相对基准']['相对基准模型年化超额'])}，信息比率为 "
                  f"{num(R['相对基准']['相对基准模型信息比率'], 2)}，即加入 CSP 反而拖累了组合。")
    perf_show = []
    short = {BM: "基准模型", RM: "研究模型(复合CSP)", RM5: "研究模型(CSP5)",
             RME: "研究模型(经典形态)", RMI: "改进模型(平滑)"}
    for name in (BM, RM, RM5, RME, RMI, "沪深300", "全A等权"):
        p = perf[name]
        perf_show.append([short.get(name, name), pct(p["年化收益率"]), pct(p["年化波动率"]),
                          num(p["夏普比率"], 2), pct(p["最大回撤"]),
                          pct(p["年化超额(vs沪深300)"]) if pd.notna(p.get("年化超额(vs沪深300)")) else "—",
                          num(p.get("信息比率(vs沪深300)"), 2) if pd.notna(p.get("信息比率(vs沪深300)")) else "—",
                          pct(p.get("平均换手率"), 1) if pd.notna(p.get("平均换手率")) else "—"])
    add_table(doc, pd.DataFrame(perf_show, columns=["组合", "年化收益", "年化波动", "Sharpe",
                                                    "最大回撤", "年化超额", "信息比率", "周换手率"]),
              "表 7  各组合的回测绩效（2021-01 至 2025-12，253 周）", font_size=8.5,
              col_widths=[4.0, 2.2, 2.2, 1.8, 2.2, 2.2, 2.0, 2.0])
    add_figure(doc, "fig05_回测净值.png", "图 7  基准模型、研究模型与沪深 300 的累计净值及相对超额净值")
    add_figure(doc, "fig06_回撤.png", "图 8  组合回撤对比")
    add_para(doc, f"分年度看（表 8、图 9），基准模型在 2021 年取得 {pct(_year(BM, 2021))} 的收益，"
                  f"研究模型为 {pct(_year(RM, 2021))}；2022 年两者分别为 {pct(_year(BM, 2022))} 与 "
                  f"{pct(_year(RM, 2022))}；2023 年差距最大，基准模型 {pct(_year(BM, 2023))}、"
                  f"研究模型 {pct(_year(RM, 2023))}；2024 年分别为 {pct(_year(BM, 2024))} 与 "
                  f"{pct(_year(RM, 2024))}；2025 年分别为 {pct(_year(BM, 2025))} 与 {pct(_year(RM, 2025))}。"
                  f"研究模型在每一年都落后于基准模型，说明这一拖累不是某一年极端行情造成的，"
                  f"而是系统性的。")
    yr = pd.DataFrame(yearly_rows)
    yr = yr[yr["年份"] >= 2021]
    yr_show = [[str(int(r["年份"])), pct(r[BM]), pct(r[RM]), pct(r[RMI]),
                pct(r["沪深300"]), pct(r["全A等权"])] for _, r in yr.iterrows()]
    add_table(doc, pd.DataFrame(yr_show, columns=["年份", "基准模型", "研究模型", "改进模型",
                                                  "沪深300", "全A等权"]),
              "表 8  各组合的分年度收益", font_size=9.5,
              col_widths=[1.8, 2.8, 2.8, 2.6, 2.6, 2.6])
    add_figure(doc, "fig07_分年度收益.png", "图 9  各组合分年度收益对比")

    add_heading(doc, "4.7 收益归因：成本、换手与因子权重", 2)
    add_para(doc, f"为什么一个 IC 显著、分组单调的因子无法改善组合？表 9 的毛净收益分解给出了直接答案。"
                  f"基准模型的毛年化收益为 {pct(cost[BM]['毛年化收益'])}、净收益 {pct(cost[BM]['净年化收益'])}，"
                  f"成本拖累仅 {pp(cost[BM]['成本拖累(年化)'] * 100)} 个百分点；研究模型的毛年化收益为 "
                  f"{pct(cost[RM]['毛年化收益'])}，实际上高于基准模型约 "
                  f"{pp((cost[RM]['毛年化收益'] - cost[BM]['毛年化收益']) * 100)} 个百分点，说明 CSP 确实"
                  f"贡献了增量 alpha；但其周度换手率高达 {pct(cost[RM]['平均周换手率'], 1)}，年化换手 "
                  f"{num(cost[RM]['年化换手倍数'], 1)} 倍，成本拖累达到 {pp(cost[RM]['成本拖累(年化)'] * 100)} 个百分点，"
                  f"最终净收益反而低于基准模型 {pct(bm['年化收益率'] - rm['年化收益率'])}。")
    cost_show = []
    for name in (BM, RM, RMI):
        c = cost[name]
        cost_show.append([short.get(name, name), pct(c["毛年化收益"]), pct(c["净年化收益"]),
                          pct(c["成本拖累(年化)"]), pct(c["平均周换手率"], 1),
                          num(c["年化换手倍数"], 1)])
    add_table(doc, pd.DataFrame(cost_show, columns=["模型", "毛年化收益", "净年化收益",
                                                    "成本拖累", "周换手率", "年化换手倍数"]),
              "表 9  收益的成本归因（单边成本 0.15%）", font_size=9,
              col_widths=[4.2, 2.6, 2.6, 2.6, 2.4, 2.8])
    add_para(doc, f"图 10 的因子权重敏感性进一步印证了这一点。当 CSP 在综合得分中的权重从 0 提高到 20%、"
                  f"33%、50%、80% 时，组合年化收益从 {pct(wsen[0.0]['年化收益'])} 单调下降到 "
                  f"{pct(wsen[0.2]['年化收益'])}、{pct(wsen[0.3333]['年化收益'])}、"
                  f"{pct(wsen[0.5]['年化收益'])} 与 {pct(wsen[0.8]['年化收益'])}，Sharpe 比率也同步下降。"
                  f"权重越高，换手率越高、成本拖累越大，同时组合偏离小市值价值暴露越远。这说明在"
                  f"“Size+BP 基准已经非常强”的样本里，CSP 的正 alpha 不足以补偿其引入的换手成本。")
    add_figure(doc, "fig10_敏感性.png", "图 10  因子权重与交易成本敏感性")

    add_heading(doc, "4.8 改进：信号平滑与换手控制", 2)
    add_para(doc, f"如果成本是主要矛盾，那么降低换手应当能恢复因子价值。本文对复合 CSP 做 4 周简单"
                  f"平均，再按同样的方式合成得分。改进模型的毛年化收益为 {pct(cost[RMI]['毛年化收益'])}，"
                  f"略低于未平滑版本，但周换手率从 {pct(cost[RM]['平均周换手率'], 1)} 降至 "
                  f"{pct(cost[RMI]['平均周换手率'], 1)}，成本拖累降至 {pp(cost[RMI]['成本拖累(年化)'] * 100)} 个百分点，"
                  f"净年化收益回升至 {pct(rmi['年化收益率'])}，与基准模型的 {pct(bm['年化收益率'])} 基本持平；"
                  f"Sharpe 比率 {num(rmi['夏普比率'])} 略高于基准模型的 {num(bm['夏普比率'])}，"
                  f"最大回撤 {pct(rmi['最大回撤'])} 小于基准模型的 {pct(bm['最大回撤'])}，"
                  f"Calmar 比率 {num(rmi['卡玛比率'])} 也高于基准模型的 {num(bm['卡玛比率'])}。")
    add_para(doc, f"表 7 中另一个值得关注的对照是经典形态事件因子：它的周换手率仅 "
                  f"{pct(perf[RME]['平均换手率'], 1)}，年化收益 {pct(perf[RME]['年化收益率'])}，"
                  f"Sharpe 比率 {num(perf[RME]['夏普比率'])}，信息比率 {num(perf[RME]['信息比率(vs沪深300)'], 2)}，"
                  f"在低换手条件下取得了与基准模型接近的收益，性价比突出。这提示我们，形态信息的价值可能更多体现在"
                  f"“低频、事件驱动”的用法上，而不是每周重估的高频打分。")
    add_para(doc, f"为检验结论的稳健性，本文还测试了组合分位（5%/10%/20%）、单边成本（0/15/30bp）与"
                  f"调仓频率（1/2/4 周）的组合。结果显示：提高组合分位或降低调仓频率都会降低收益，"
                  f"因为在 2021—2025 年的小市值行情中，越集中、越频繁地暴露于小市值价值股，收益越高；"
                  f"成本从 0 提高到 30bp 时，研究模型的年化收益从 35% 左右降至 21% 左右，"
                  f"成本敏感性远高于基准模型。这一结果再次说明，换手率是形态类因子最脆弱的环节。")

    add_heading(doc, "4.9 与原报告结论的对比", 2)
    add_para(doc, "把本文结果与原报告逐项对照，可以得到“验证”“部分验证”“未复现”三类结论。")
    add_para(doc, "得到验证的结论有三项。第一，形态因子具有显著的截面预测能力。复合 CSP 的月度 IC 为"
                  "6.31%、ICIR 为 0.83、IC 为正比例 83.1%，十分组单调性 0.988，多空组合 t 值 5.65，"
                  "这些指标在方向和显著性上与原报告一致。第二，多周期复合优于单周期。复合 CSP 的 IC "
                  "比最好的单周期因子高约 2.3 个百分点，验证了报告“多周期视角能提升绩效”的判断。"
                  "第三，形态因子与常见量价因子低相关。本文测得的截面相关系数绝对值均小于 0.12，"
                  "支持报告“因子独立性较强、信息增量明显”的结论。")
    add_para(doc, "部分验证的结论有两项。第一，CSP 多头组合相对全 A 等权的年化超额约 12.4 个百分点，"
                  "与原报告的 11.5% 接近，但这是以全 A 等权为基准的结果；若以 Size+BP 组合为基准，"
                  "CSP 并未带来增量。第二，顶部反转形态的负超额得到验证，但底部反转形态的正超额没有"
                  "出现，说明经典形态的“底部反转”含义对样本期与市场状态高度敏感。")
    add_para(doc, "未能复现的结论主要是“加入 CSP 可提升多因子组合收益”。原报告以 CSP 为核心因子"
                  "构建选股组合，并以全市场等权或宽基指数为基准，因而 CSP 的 alpha 可以直接转化为"
                  "超额收益；本文的基准模型是 Size+BP 的强多头组合，2021—2025 年小市值价值风格"
                  "极致演绎，基准模型的毛年化收益达到 33.53%，CSP 的增量 alpha 在扣除换手成本后被"
                  "完全抵消。这不是对因子有效性的否定，而是说明“因子有效”与“组合有效”之间存在"
                  "成本、基准与风格暴露三道门槛。")
    add_para(doc, "差异的来源可以归纳为四点。其一，样本区间与市场环境不同，2021—2025 年的小市值"
                  "行情使 Size 因子的收益远高于 2014—2020 年，基准模型因此异常强势。其二，本文的"
                  "复合因子只使用相对全市场的单一标签，且对形态得分做了收缩，IC 强度低于原报告。"
                  "其三，本文采用周度调仓与统一的 0.15% 单边成本，而原报告为月度调仓且成本假设更"
                  "精细，换手对本文结果的影响更大。其四，本文的行业中性化与 1%/99% 缩尾会削弱"
                  "因子在极端小市值股票上的暴露，而这类股票恰是 2021—2025 年收益最高的群体。")
    add_para(doc, "此外，本文还发现一个原报告未强调的现象：CSP 的多空收益主要来自空头端。G1 的年化"
                  "收益为 −9.46%，G10 为 27.18%，空头端的负收益幅度更大。这意味着 CSP 更适合作为"
                  "“风险剔除”或“负面清单”使用，而不是单纯的多头选股因子。对于以多头为主、且基准"
                  "已经很强的组合，这一特性解释了为什么加入 CSP 后收益不升反降。")


def _year(name: str, year: int) -> float:
    for row in yearly_rows:
        if int(row["年份"]) == year:
            return float(row[name])
    return float("nan")


# ---------------------------------------------------------------------------
# 五、总结与展望
# ---------------------------------------------------------------------------
def add_conclusion(doc) -> None:
    add_heading(doc, "五、总结与展望", 1)
    add_para(doc, "本文以《K线形态信息中的alpha》为对象，在 2020—2025 年全 A 股数据上完成了 K 线形态"
                  "划分、41 472 种组合形态构建、多周期 CSP 因子与复合因子的复刻，并在统一的周度"
                  "回测框架下比较了 Size+BP 基准模型、加入 CSP 的研究模型与沪深 300 指数。主要结论"
                  "可以归纳为以下五点。")
    add_para(doc, "第一，形态因子的统计有效性得到验证。复合 CSP 的月度 IC 为 6.31%，ICIR 为 0.83，"
                  "IC 为正比例 83.1%；十分组年化收益从 −9.46% 单调上升到 27.18%，单调性 0.988；"
                  "多空组合年化收益 37.54%，Sharpe 比率 2.93；Fama-MacBeth 回归中控制五个风格因子后"
                  "系数仍显著为正。这些结果说明，K 线形态确实包含对未来收益有增量解释力的信息，"
                  "原报告的核心命题在较短样本上依然成立。")
    add_para(doc, "第二，复合因子的表现优于单周期因子。CSP_1、CSP_5、CSP_10 的 20 日 IC 分别为 4.24%、"
                  "2.59% 与 2.28%，复合后提升至 6.55%。多周期综合通过平均不同频率的形态观点，"
                  "有效降低了单一周期的噪声，这与原报告强调的“多周期、多维度评价”一致。")
    add_para(doc, "第三，因子独立性强，但独立性不等于组合价值。CSP 与市值、BP、动量、反转、波动率、"
                  "换手率的截面相关性绝对值均小于 0.12，信息上与 Size+BP 互补；然而在组合层面，"
                  "加入 CSP 的研究模型毛年化收益为 35.35%、高于基准模型的 33.53%，净收益却因 35.65% 的"
                  "周换手率与 7.29 个百分点的年化成本拖累降至 28.07%，低于基准模型的 32.19%。这说明"
                  "“因子有增量信息”与“组合能获得增量收益”之间，还隔着交易成本这一道现实约束。")
    add_para(doc, "第四，换手控制是形态类因子落地的关键。对 CSP 做 4 周平滑后，换手率降至 14.77%，"
                  "成本拖累降至 3.06 个百分点，净年化收益回升至 31.87%，与基准模型基本持平，Sharpe 比率"
                  "（1.20 对 1.19）与最大回撤（−25.62% 对 −26.81%）略优。经典形态事件因子的换手率"
                  "仅 8.22%，年化收益 31.24%，信息比率 1.55，是表现最好的研究模型之一。这两点共同"
                  "说明，形态信息更适合低频、事件驱动的用法，而不是每周重估的高频打分。")
    add_para(doc, "第五，经典形态的“反转”含义具有状态依赖性。顶部十字星、顶部倒锤头线叠加放量后，"
                  "未来 20 日超额收益分别为 −3.53% 与 −0.98%，显著为负，验证了原报告的看空结论；"
                  "但底部十字星、底部锤头线的超额收益分别为 0.37% 与 −1.54%，未能复现原报告的正超额。"
                  "在 2021—2025 年小市值占优、但指数整体走弱的环境中，“底部”信号更容易演化为下跌中继，"
                  "机械套用经典形态可能带来损失。")
    add_para(doc, "本文的研究仍存在若干局限。其一，样本区间较短且风格单一。2020—2025 年仅覆盖一轮"
                  "完整的小市值价值行情与一次微盘股流动性危机，结论对其他风格环境的适用性有待检验。"
                  "其二，形态得分的估计存在误差。41 472 种形态对应约 1 978 万个样本，但样本分布高度不均，"
                  "单个形态的样本中位数仅 15 个；尽管使用了收缩估计，稀疏形态的得分仍不稳定。其三，行业分类为静态表，未做历史"
                  "成分回溯，行业中性化可能引入轻微的前视偏差。其四，交易成本采用固定的单边 0.15%，"
                  "未考虑冲击成本、涨跌停无法成交与流动性分层，对高换手策略偏乐观。其五，本文只用"
                  "相对全市场等权一个标签维度，未复刻原报告的行业相对与风格残差维度，这也是 IC 强度"
                  "低于原报告的原因之一。其六，回测只做多头，未检验多空组合在扣除融券成本后的净收益。")
    add_para(doc, "后续研究可以从四个方向改进。第一，扩展数据频率与维度，引入分钟级行情、行业指数"
                  "与财务数据，复刻原报告的行业轮动与多维度形态评价。第二，改进形态得分的估计方法，"
                  "例如用贝叶斯层次模型或机器学习模型对形态—收益关系做非线性建模，替代简单的历史"
                  "均值映射。第三，建立更精细的交易成本模型，把冲击成本、成交概率与调仓阈值纳入"
                  "组合优化，直接以净收益为目标求解最优换手。第四，把 CSP 用作风险剔除或负面清单，"
                  "与反转、低波等因子结合，检验其在“排除坏股票”而非“挑选好股票”上的价值。")
    add_para(doc, "总体而言，本文完成了课程要求的因子复刻、因子检验与多因子回测，并在复刻基础上对"
                  "原报告结论进行了验证与反思。CSP 因子的统计有效性是真实的，但其经济价值高度依赖"
                  "样本环境、基准强弱与交易成本。对量化研究而言，一个因子“显著”只是起点，"
                  "能否在扣除成本后改善组合，才是更接近实务的检验标准。")


# ---------------------------------------------------------------------------
# 六、参考文献
# ---------------------------------------------------------------------------
def add_references(doc) -> None:
    add_heading(doc, "六、参考文献", 1)
    refs = [
        "[1] 缪铃凯, 刘富兵. K线形态信息中的alpha[R]. 上海: 国盛证券金融工程研究, 2024-09-07.",
        "[2] 国盛金工. 【国盛金工】量化专题：K线形态信息中的alpha[EB/OL]. 微信公众号“留富兵法”, "
        "2024-09-10. https://mp.weixin.qq.com/s/IT933xVRmDNL4wP4Y6-2CQ.",
        "[3] Fama E F, MacBeth J D. Risk, Return, and Equilibrium: Empirical Tests[J]. "
        "Journal of Political Economy, 1973, 81(3): 607-636.",
        "[4] Jegadeesh N, Titman S. Returns to Buying Winners and Selling Losers: Implications for "
        "Stock Market Efficiency[J]. Journal of Finance, 1993, 48(1): 65-91.",
        "[5] Amihud Y. Illiquidity and Stock Returns: Cross-Section and Time-Series Effects[J]. "
        "Journal of Financial Markets, 2002, 5(1): 31-56.",
        "[6] Lo A W, Mamaysky H, Wang J. Foundations of Technical Analysis: Computational Algorithms, "
        "Statistical Inference, and Empirical Implementation[J]. Journal of Finance, 2000, 55(4): 1705-1765.",
        "[7] Brock W, Lakonishok J, LeBaron B. Simple Technical Trading Rules and the Stochastic "
        "Properties of Stock Returns[J]. Journal of Finance, 1992, 47(5): 1731-1764.",
        "[8] Newey W K, West K D. A Simple, Positive Semi-Definite, Heteroskedasticity and "
        "Autocorrelation Consistent Covariance Matrix[J]. Econometrica, 1987, 55(3): 703-708.",
        "[9] Sharpe W F. Mutual Fund Performance[J]. Journal of Business, 1966, 39(1): 119-138.",
        "[10] 石川, 刘洋溢, 连祥斌. 因子投资：方法与实践[M]. 北京: 电子工业出版社, 2020.",
        "[11] 数据来源：课程提供的 A 股日频行情数据集（2020-01-02 至 2025-12-31），"
        "包含开高低收、成交量、成交额、复权因子、总市值、PB、PE_TTM、ST 标记及沪深 300 指数行情；"
        "行业分类采用申万一级行业。",
        "[12] 计算与绘图工具：Python 3.14，NumPy、pandas、SciPy、Matplotlib、python-docx。",
    ]
    for r in refs:
        add_para(doc, r, size=10.5, indent=False, align=WD_ALIGN_PARAGRAPH.LEFT,
                 space_after=2, line=1.25)


# ---------------------------------------------------------------------------
# 字数统计与主流程
# ---------------------------------------------------------------------------
def count_chinese(text: str) -> int:
    return len(re.findall(r"[\u4e00-\u9fff]", text))


def run_report() -> None:
    load_results()
    doc = setup_document()
    add_cover(doc)
    add_division(doc)
    add_abstract(doc)
    add_introduction(doc)
    add_factor_section(doc)
    add_data_section(doc)
    add_results_section(doc)
    add_conclusion(doc)
    add_references(doc)
    doc.save(OUT_PATH)

    # 字数统计（正文，不含封面与参考文献之后的工具说明）
    body_text = "\n".join(p.text for p in doc.paragraphs)
    table_text = "\n".join(c.text for t in doc.tables for row in t.rows for c in row.cells)
    total = count_chinese(body_text + table_text)
    print(f"报告已生成: {OUT_PATH}")
    print(f"中文字符数（含表格）: {total}")
    print(f"段落数: {len(doc.paragraphs)}, 表格数: {len(doc.tables)}")







# ===========================================================================
# 第九部分  一键运行
# ===========================================================================
def main() -> None:
    t0 = time.time()
    print("=" * 70)
    print("STEP 1/3  数据构建 → CSP 因子 → 因子检验 → 多因子回测")
    print("=" * 70)
    run_pipeline()
    print("\n" + "=" * 70)
    print("STEP 2/3  生成图表")
    print("=" * 70)
    run_figures()
    print("\n" + "=" * 70)
    print("STEP 3/3  生成 Word 报告")
    print("=" * 70)
    run_report()
    print(f"\n全部完成，总耗时 {time.time() - t0:.1f}s")


if __name__ == "__main__":
    main()
