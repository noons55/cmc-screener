#!/usr/bin/env python3
"""
CMC Invest Stock Screener  -  Strong Buy to Strong Sell ratings
================================================================

Scans the stocks in the 16 markets CMC Invest offers (ASX + 15 international),
scores every stock on technicals AND fundamentals, and writes a ranked,
colour-coded Excel report plus a CSV.

CMC Invest has no public data API for retail accounts, so the stock list,
prices and company financials come from Yahoo Finance via the free `yfinance`
library, filtered to the same exchanges CMC Invest gives you access to.

SETUP (once)
    pip install yfinance pandas numpy openpyxl

RUN
    python cmc_screener.py --markets au us          # ASX + US (good first run)
    python cmc_screener.py                          # all 16 CMC Invest markets
    python cmc_screener.py --tickers mylist.txt     # just your own list / watchlist
    python cmc_screener.py --fundamentals quick     # skip per-stock financials (much faster)
    python cmc_screener.py --markets au --max-per-market 50   # quick test run
    python cmc_screener.py --list-markets           # show market codes
    python cmc_screener.py --help                   # all options

WHAT IT ANALYSES
    Technical (20 signals)
      Trend     : price vs 50/200-day MA, golden/death cross, 200-day slope,
                  20/50 EMA, ADX/DMI
      Momentum  : RSI, MACD, Stochastic, CCI, 3-month return, 12-1 month momentum,
                  relative strength vs local index, distance from 52-week high
      Volume    : Chaikin Money Flow, On-Balance Volume, Money Flow Index,
                  20-day VWAP, volume surge
    Fundamental (up to 22 signals)
      Valuation : P/E, forward P/E, P/B, P/S, EV/EBITDA - each compared with the
                  median of the stock's sector peers in the same market - plus
                  PEG and dividend yield
      Quality   : profitability, ROE, ROA, operating margin, net margin
      Growth    : revenue growth, earnings growth, forecast EPS growth
      Health    : debt/equity, current ratio, free-cash-flow yield, payout ratio
      Sentiment : analyst consensus, analyst price-target upside

OUTPUT (in ./screener_output)
    site/index.html                 interactive screener page with the results built in -
                                    opens in your browser when the run finishes. To put it
                                    online, upload the 'site' folder to any static host
                                    (Netlify Drop, GitHub Pages, Cloudflare Pages, your own site).
    cmc_ratings_YYYY-MM-DD.xlsx     colour-coded spreadsheet
    cmc_ratings_YYYY-MM-DD.csv      same data as CSV

Prices are cached for the day and company financials for 7 days, saved as they
download - if a big run is interrupted, run the same command again and it
carries on where it stopped.

Not financial advice. The rating is a mechanical summary of indicators -
use it to shortlist, then do your own research.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import logging
import math
import re
import shutil
import sys
import threading
import time
import warnings
import zlib
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path

import numpy as np
import pandas as pd

try:
    import yfinance as yf
    from yfinance import EquityQuery
except ImportError:  # pragma: no cover
    sys.exit("yfinance is not installed. Run:  pip install yfinance pandas numpy openpyxl")

warnings.filterwarnings("ignore", category=FutureWarning)
warnings.filterwarnings("ignore", category=RuntimeWarning)
warnings.filterwarnings("ignore", module=r"yfinance(\..*)?")          # pandas deprecation noise from inside yfinance
warnings.filterwarnings("ignore", message=r".*inplace keyword.*")
logging.getLogger("yfinance").setLevel(logging.CRITICAL)


# =============================================================================
# Settings you might want to tweak
# =============================================================================

# Rating bands on the final score (-1 .. +1, shown in the report as -100 .. +100)
STRONG_BUY_AT = 0.50
BUY_AT = 0.15
SELL_AT = -0.15
STRONG_SELL_AT = -0.50

# Weighting of the groups inside the technical and fundamental scores
TECH_WEIGHTS = {"trend": 0.40, "momentum": 0.40, "volume": 0.20}
FUND_WEIGHTS = {"valuation": 0.25, "quality": 0.25, "growth": 0.20, "health": 0.15, "sentiment": 0.15}

DEFAULT_FUND_WEIGHT = 0.40   # share of the final score that comes from fundamentals
MIN_FUND_SIGNALS = 3         # need at least this many fundamental signals to use fundamentals

# Valuation vs sector peers: cheaper than 80% of the peer median = bullish,
# more than 150% of it = bearish. Needs at least SECTOR_MIN_PEERS peers.
CHEAP_VS_PEERS = 0.80
DEAR_VS_PEERS = 1.50
SECTOR_MIN_PEERS = 5
# Fallback absolute cut-offs (bullish at or below, bearish above) when there aren't enough peers
ABS_VALUATION = {"trailing_pe": (15, 40), "forward_pe": (14, 35), "pb": (1.5, 6),
                 "ps": (1.5, 8), "ev_ebitda": (8, 20)}
# Leverage / liquidity / sales ratios don't mean much for banks and insurers
FINANCIAL_SECTORS = {"Financial Services"}

MIN_BARS = 60                # skip stocks with less daily price history than this
STALE_DAYS = 10              # skip stocks whose last price is older than this (halted/delisted)
DOWNLOAD_CHUNK = 150         # tickers per Yahoo price request
INFO_WORKERS = 6             # parallel requests when pulling company financials
FUND_CACHE_DAYS = 7          # re-use downloaded financials for this many days

RATING_ORDER = ["Strong Buy", "Buy", "Neutral", "Sell", "Strong Sell"]


# =============================================================================
# The 16 CMC Invest markets
# =============================================================================

@dataclass(frozen=True)
class Market:
    code: str            # Yahoo region code
    name: str
    exchanges: tuple     # Yahoo exchange codes for the main board(s) CMC trades
    benchmark: str       # index used for relative strength
    brokerage: str       # CMC Invest brokerage tier
    suffixes: tuple      # Yahoo ticker suffixes


ASX_FEE = "ASX - $0 first buy <=$1k/day"
TIER1 = "Tier 1 - $0 brokerage"
TIER2 = "Tier 2 - A$59 / 0.59%"

MARKETS = {m.code: m for m in [
    Market("au", "Australia",      ("ASX",),                              "^AXJO",   ASX_FEE, (".AX",)),
    Market("us", "United States",  ("NMS", "NGM", "NCM", "NYQ", "ASE"),   "^GSPC",   TIER1,   ("",)),
    Market("gb", "United Kingdom", ("LSE",),                              "^FTSE",   TIER1,   (".L",)),
    Market("ca", "Canada",         ("TOR", "VAN"),                        "^GSPTSE", TIER1,   (".TO", ".V")),
    Market("jp", "Japan",          ("JPX",),                              "^N225",   TIER1,   (".T",)),
    Market("de", "Germany",        ("GER",),                              "^GDAXI",  TIER2,   (".DE",)),
    Market("ch", "Switzerland",    ("EBS",),                              "^SSMI",   TIER2,   (".SW",)),
    Market("fr", "France",         ("PAR",),                              "^FCHI",   TIER2,   (".PA",)),
    Market("be", "Belgium",        ("BRU",),                              "^BFX",    TIER2,   (".BR",)),
    Market("nl", "Netherlands",    ("AMS",),                              "^AEX",    TIER2,   (".AS",)),
    Market("es", "Spain",          ("MCE",),                              "^IBEX",   TIER2,   (".MC",)),
    Market("se", "Sweden",         ("STO",),                              "^OMX",    TIER2,   (".ST",)),
    Market("dk", "Denmark",        ("CPH",),                              "^OMXC25", TIER2,   (".CO",)),
    Market("sg", "Singapore",      ("SES",),                              "^STI",    TIER2,   (".SI",)),
    Market("hk", "Hong Kong",      ("HKG",),                              "^HSI",    TIER2,   (".HK",)),
    Market("nz", "New Zealand",    ("NZE",),                              "^NZ50",   TIER2,   (".NZ",)),
]}

SUFFIX_TO_MARKET = {s: m.code for m in MARKETS.values() for s in m.suffixes if s}


# =============================================================================
# Small helpers
# =============================================================================

def num(x) -> float:
    """Coerce Yahoo values (numbers, strings, {'raw': ..}) to float, NaN if unusable."""
    try:
        if x is None or isinstance(x, bool):
            return np.nan
        if isinstance(x, dict):
            x = x.get("raw")
        f = float(x)
        return f if math.isfinite(f) else np.nan
    except (TypeError, ValueError):
        return np.nan


def ok(x) -> bool:
    if x is None:
        return False
    if isinstance(x, (float, np.floating)):
        return not math.isnan(x)
    return True


def first_ok(*vals) -> float:
    for v in vals:
        f = num(v)
        if ok(f):
            return f
    return np.nan


def pct(a, b) -> float:
    return (a / b - 1) * 100 if ok(a) and ok(b) and b > 0 else np.nan


def parse_analyst(x) -> float:
    """'2.1 - Buy' -> 2.1   (1 = Strong Buy ... 5 = Strong Sell)"""
    if x is None:
        return np.nan
    if isinstance(x, (int, float)):
        return num(x)
    m = re.match(r"\s*([\d.]+)", str(x))
    return float(m.group(1)) if m else np.nan


def with_retry(fn, what: str, tries: int = 4, base_wait: float = 3.0, quiet: bool = False):
    for attempt in range(1, tries + 1):
        try:
            return fn()
        except Exception as e:  # noqa: BLE001 - Yahoo throws all sorts
            if attempt == tries:
                if not quiet:
                    print(f"    ! {what} failed after {tries} tries: {e}")
                return None
            text = f"{type(e).__name__} {e}".lower()
            limited = "ratelimit" in text or "rate limit" in text or "429" in text or "too many" in text
            wait = base_wait * (2 ** (attempt - 1)) * (4 if limited else 1)
            if not quiet or limited:
                print(f"    {what}: {'rate limited' if limited else 'error'} - retrying in {wait:.0f}s")
            time.sleep(wait)


def market_for_ticker(t: str) -> str:
    if "." in t:
        suffix = "." + t.rsplit(".", 1)[1].upper()
        return SUFFIX_TO_MARKET.get(suffix, "other")
    return "us"


def naive_day(ts) -> pd.Timestamp:
    ts = pd.Timestamp(ts)
    if ts.tzinfo is not None:
        ts = ts.tz_convert(None)
    return ts.normalize()


# =============================================================================
# 1. Build the stock universe
# =============================================================================

def quote_to_row(q: dict, market: str) -> dict:
    div = num(q.get("trailingAnnualDividendYield"))
    return {
        "ticker": q.get("symbol"),
        "name": q.get("longName") or q.get("shortName") or q.get("displayName"),
        "market": market,
        "exchange": q.get("fullExchangeName") or q.get("exchange"),
        "sector": q.get("sector"),
        "currency": q.get("currency"),
        "market_cap": num(q.get("marketCap")),
        "trailing_pe": num(q.get("trailingPE")),
        "forward_pe": num(q.get("forwardPE")),
        "pb": num(q.get("priceToBook")),
        "eps_ttm": num(q.get("epsTrailingTwelveMonths")),
        "eps_fwd": num(q.get("epsForward")),
        "analyst": parse_analyst(q.get("averageAnalystRating")),
        "div_yield_pct": div * 100 if ok(div) else np.nan,
        "full": False,
    }


def market_universe(m: Market, min_mcap: float, min_volume: float, cap: int | None) -> list[dict]:
    parts = [EquityQuery("eq", ["region", m.code])]
    if len(m.exchanges) == 1:
        parts.append(EquityQuery("eq", ["exchange", m.exchanges[0]]))
    else:
        parts.append(EquityQuery("is-in", ["exchange", *m.exchanges]))
    if min_mcap > 0:
        parts.append(EquityQuery("gte", ["intradaymarketcap", min_mcap]))
    if min_volume > 0:
        parts.append(EquityQuery("gte", ["avgdailyvol3m", min_volume]))
    query = EquityQuery("and", parts)

    rows, offset, sort_field = [], 0, "intradaymarketcap"
    while True:
        size = 250 if cap is None else max(1, min(250, cap - offset))
        res = with_retry(
            lambda: yf.screen(query, offset=offset, size=size,
                              sortField=sort_field, sortAsc=False),
            f"{m.name} stock list (offset {offset})",
        )
        if res is None and offset == 0 and sort_field != "ticker":
            sort_field = "ticker"            # fall back to Yahoo's default ordering
            continue
        quotes = (res or {}).get("quotes") or []
        total = (res or {}).get("total")
        rows.extend(quote_to_row(q, m.code) for q in quotes
                    if q.get("symbol") and q.get("quoteType", "EQUITY") == "EQUITY")
        offset += len(quotes)
        if not quotes or len(quotes) < size:
            break
        if cap is not None and offset >= cap:
            break
        if total and offset >= total:
            break
        time.sleep(0.4)
    return rows


def universe_from_screener(codes, min_mcap, min_volume, cap) -> pd.DataFrame:
    rows = []
    for code in codes:
        m = MARKETS[code]
        print(f"  {m.name:<15}", end="", flush=True)
        got = market_universe(m, min_mcap, min_volume, cap)
        print(f"{len(got):>6} stocks")
        rows.extend(got)
    df = pd.DataFrame(rows)
    if df.empty:
        return df
    return df.drop_duplicates("ticker").reset_index(drop=True)


def read_tickers_file(path: str, suffix: str | None) -> list[str]:
    p = Path(path)
    if not p.exists():
        sys.exit(f"Tickers file not found: {p}")
    if p.suffix.lower() in (".csv", ".xlsx", ".xls"):
        df = pd.read_csv(p) if p.suffix.lower() == ".csv" else pd.read_excel(p)
        col = next((c for c in df.columns
                    if str(c).strip().lower() in ("ticker", "symbol", "code", "stock", "asx code")),
                   df.columns[0])
        raw = df[col].astype(str).tolist()
    else:
        raw = re.split(r"[\s,;]+", p.read_text(encoding="utf-8-sig"))
    out = []
    for t in raw:
        t = t.strip().upper()
        if not t or t.startswith("#") or t in ("NAN", "TICKER", "SYMBOL", "CODE"):
            continue
        if suffix and "." not in t:
            t += suffix if suffix.startswith(".") else "." + suffix
        out.append(t)
    return list(dict.fromkeys(out))


# =============================================================================
# 2. Company financials (one request per stock, cached for FUND_CACHE_DAYS)
# =============================================================================

def map_info(info: dict) -> dict:
    """Pick the fundamentals we score out of Yahoo's quote summary."""
    price = first_ok(info.get("currentPrice"), info.get("regularMarketPrice"), info.get("previousClose"))
    mcap = num(info.get("marketCap"))

    div_rate = num(info.get("dividendRate"))
    tady = num(info.get("trailingAnnualDividendYield"))
    if ok(div_rate) and ok(price) and price > 0:
        div_yield = div_rate / price * 100
    elif ok(tady):
        div_yield = tady * 100
    else:
        div_yield = np.nan

    # FCF is reported in the company's financial currency; only compare it with the
    # market cap when both are in the same currency (e.g. BHP reports in USD, trades in AUD).
    ccy, fin_ccy = info.get("currency"), info.get("financialCurrency")
    fcf = num(info.get("freeCashflow"))
    same_ccy = not fin_ccy or not ccy or fin_ccy == ccy
    fcf_yield = fcf / mcap * 100 if ok(fcf) and ok(mcap) and mcap > 0 and same_ccy else np.nan

    short = num(info.get("shortPercentOfFloat"))
    return {
        "name": info.get("longName") or info.get("shortName"),
        "exchange": info.get("fullExchangeName") or info.get("exchange"),
        "sector": info.get("sector"),
        "industry": info.get("industry"),
        "currency": ccy,
        "market_cap": mcap,
        "trailing_pe": num(info.get("trailingPE")),
        "forward_pe": num(info.get("forwardPE")),
        "peg": first_ok(info.get("pegRatio"), info.get("trailingPegRatio")),
        "pb": num(info.get("priceToBook")),
        "ps": num(info.get("priceToSalesTrailing12Months")),
        "ev_ebitda": num(info.get("enterpriseToEbitda")),
        "eps_ttm": num(info.get("trailingEps")),
        "eps_fwd": num(info.get("forwardEps")),
        "div_yield_pct": div_yield,
        "payout_ratio": num(info.get("payoutRatio")),
        "roe": num(info.get("returnOnEquity")),
        "roa": num(info.get("returnOnAssets")),
        "op_margin": num(info.get("operatingMargins")),
        "profit_margin": num(info.get("profitMargins")),
        "revenue_growth": num(info.get("revenueGrowth")),
        "earnings_growth": num(info.get("earningsGrowth")),
        "debt_to_equity": num(info.get("debtToEquity")),    # Yahoo gives this as a %, 85 = 0.85x
        "current_ratio": num(info.get("currentRatio")),
        "free_cash_flow": fcf,
        "fcf_yield_pct": fcf_yield,
        "analyst": num(info.get("recommendationMean")),
        "analyst_count": num(info.get("numberOfAnalystOpinions")),
        "target_upside_pct": pct(num(info.get("targetMeanPrice")), price),
        "short_float_pct": short * 100 if ok(short) else np.nan,
    }


def fetch_one(ticker: str) -> dict | None:
    info = with_retry(lambda: yf.Ticker(ticker).info, f"financials {ticker}", tries=3, quiet=True)
    if not isinstance(info, dict) or not info:
        return None
    return map_info(info)


class FundCache:
    """Append-only JSON-lines cache so an interrupted run can pick up where it stopped."""

    def __init__(self, path: Path, max_age_days: int):
        self.path, self.lock, self.fh = path, threading.Lock(), None
        self.data: dict[str, dict] = {}
        self.stamp: dict[str, str] = {}
        now = datetime.now()
        if path.exists():
            with open(path, encoding="utf-8") as f:
                for line in f:
                    try:
                        rec = json.loads(line)
                        # spread refreshes over 5 days so they don't all expire on the same run
                        days = max_age_days + zlib.crc32(rec["t"].encode()) % 5 - 2
                        if datetime.fromisoformat(rec["ts"]) >= now - timedelta(days=days):
                            self.data[rec["t"]], self.stamp[rec["t"]] = rec["d"], rec["ts"]
                    except Exception:  # noqa: BLE001 - skip a half-written line
                        continue

    def get(self, ticker: str):
        return self.data.get(ticker)

    def put(self, ticker: str, d: dict) -> None:
        ts = datetime.now().isoformat(timespec="seconds")
        with self.lock:
            self.data[ticker], self.stamp[ticker] = d, ts
            if self.fh is None:
                self.fh = open(self.path, "a", encoding="utf-8")
            self.fh.write(json.dumps({"t": ticker, "ts": ts, "d": d}) + "\n")
            self.fh.flush()

    def close(self) -> None:
        with self.lock:
            if self.fh:
                self.fh.close()
                self.fh = None
            tmp = self.path.with_suffix(".tmp")           # rewrite without expired/duplicate rows
            with open(tmp, "w", encoding="utf-8") as f:
                for t, d in self.data.items():
                    f.write(json.dumps({"t": t, "ts": self.stamp[t], "d": d}) + "\n")
            tmp.replace(self.path)


def fetch_fundamentals(tickers: list[str], cache: FundCache) -> dict[str, dict]:
    todo = [t for t in tickers if cache.get(t) is None]
    print(f"      {len(tickers) - len(todo):,} already saved from recent runs, "
          f"{len(todo):,} to download")
    if todo:
        t0, done = time.time(), 0
        pool = ThreadPoolExecutor(max_workers=INFO_WORKERS)
        try:
            futs = {pool.submit(fetch_one, t): t for t in todo}
            for f in as_completed(futs):
                d = f.result()
                done += 1
                if d is not None:
                    cache.put(futs[f], d)
                if done % 100 == 0 or done == len(todo):
                    rate = done / max(time.time() - t0, 1e-6)
                    left = (len(todo) - done) / rate / 60
                    print(f"      financials {done:,}/{len(todo):,}  (~{left:.0f} min left)", flush=True)
        except KeyboardInterrupt:
            pool.shutdown(wait=False, cancel_futures=True)
            cache.close()
            sys.exit("\nStopped. Financials downloaded so far are saved - "
                     "run the same command again to carry on.")
        pool.shutdown(wait=True)
    return {t: cache.get(t) for t in tickers if cache.get(t) is not None}


def merge_info(rec: dict, info: dict) -> None:
    """Overlay per-stock financials onto a universe row (only values Yahoo actually returned)."""
    got = False
    for k, v in info.items():
        if ok(v):
            rec[k] = v
            got = True
    rec["full"] = got


# =============================================================================
# 3. Download prices (cached per day)
# =============================================================================

PRICE_COLS = ["Open", "High", "Low", "Close", "Volume"]


def split_download(data, batch: list[str]) -> dict[str, pd.DataFrame]:
    out = {}
    if data is None or getattr(data, "empty", True):
        return out
    if isinstance(data.columns, pd.MultiIndex):
        lvl0 = set(data.columns.get_level_values(0))
        lvl1 = set(data.columns.get_level_values(1))
        for t in batch:
            if t in lvl0:
                df = data[t]
            elif t in lvl1:
                df = data.xs(t, axis=1, level=1)
            else:
                continue
            df = df[[c for c in PRICE_COLS if c in df.columns]].dropna(subset=["Close"])
            if len(df):
                out[t] = df
    elif len(batch) == 1 and "Close" in data.columns:
        df = data[[c for c in PRICE_COLS if c in data.columns]].dropna(subset=["Close"])
        if len(df):
            out[batch[0]] = df
    return out


def download_prices(tickers: list[str], cache_dir: Path, refresh: bool, label: str) -> dict:
    frames: dict[str, pd.DataFrame] = {}

    def run(ticks: list[str], tag: str):
        chunks = [ticks[i:i + DOWNLOAD_CHUNK] for i in range(0, len(ticks), DOWNLOAD_CHUNK)]
        for i, batch in enumerate(chunks, 1):
            key = hashlib.md5("|".join(batch).encode()).hexdigest()[:16]
            path = cache_dir / f"px_{key}.pkl"
            data = None
            if path.exists() and not refresh:
                try:
                    data = pd.read_pickle(path)
                except Exception:  # noqa: BLE001
                    data = None
            if data is None:
                print(f"    {label}{tag}: price batch {i}/{len(chunks)}", flush=True)
                data = with_retry(
                    lambda: yf.download(batch, period="2y", interval="1d", group_by="ticker",
                                        auto_adjust=True, threads=True, progress=False),
                    f"price batch {i}",
                )
                if data is not None and not data.empty:
                    data.to_pickle(path)
                time.sleep(1.0)
            frames.update(split_download(data, batch))

    run(tickers, "")
    missing = [t for t in tickers if t not in frames]
    if missing and len(missing) < len(tickers):
        print(f"    retrying {len(missing)} tickers with no data...")
        time.sleep(5)
        run(missing, " (retry)")
    return frames


# =============================================================================
# 4. Technical indicators
# =============================================================================

def rsi(c: pd.Series, n: int = 14) -> pd.Series:
    d = c.diff()
    up = d.clip(lower=0).ewm(alpha=1 / n, adjust=False, min_periods=n).mean()
    dn = (-d.clip(upper=0)).ewm(alpha=1 / n, adjust=False, min_periods=n).mean()
    return 100 - 100 / (1 + up / dn)


def macd(c: pd.Series, fast=12, slow=26, signal=9):
    line = c.ewm(span=fast, adjust=False).mean() - c.ewm(span=slow, adjust=False).mean()
    sig = line.ewm(span=signal, adjust=False).mean()
    return line, sig


def adx_dmi(h, l, c, n: int = 14):
    up, dn = h.diff(), -l.diff()
    plus_dm = pd.Series(np.where((up > dn) & (up > 0), up, 0.0), index=c.index)
    minus_dm = pd.Series(np.where((dn > up) & (dn > 0), dn, 0.0), index=c.index)
    tr = pd.concat([h - l, (h - c.shift()).abs(), (l - c.shift()).abs()], axis=1).max(axis=1)
    atr = tr.ewm(alpha=1 / n, adjust=False, min_periods=n).mean()
    pdi = 100 * plus_dm.ewm(alpha=1 / n, adjust=False, min_periods=n).mean() / atr
    mdi = 100 * minus_dm.ewm(alpha=1 / n, adjust=False, min_periods=n).mean() / atr
    dx = 100 * (pdi - mdi).abs() / (pdi + mdi)
    adx = dx.ewm(alpha=1 / n, adjust=False, min_periods=n).mean()
    return adx, pdi, mdi, atr


def stochastic(h, l, c, n: int = 14, smooth: int = 3):
    lo, hi = l.rolling(n).min(), h.rolling(n).max()
    fast_k = 100 * (c - lo) / (hi - lo).replace(0, np.nan)
    k = fast_k.rolling(smooth).mean()          # slow %K (14,3)
    d = k.rolling(smooth).mean()               # %D (3)
    return k, d


def cci_last(h, l, c, n: int = 20) -> float:
    tp = ((h + l + c) / 3).iloc[-n:]
    if len(tp) < n:
        return np.nan
    mean = tp.mean()
    md = (tp - mean).abs().mean()
    return float((tp.iloc[-1] - mean) / (0.015 * md)) if md > 0 else np.nan


def cmf(h, l, c, v, n: int = 20) -> pd.Series:
    rng = (h - l).replace(0, np.nan)
    mfm = (((c - l) - (h - c)) / rng).fillna(0)
    return (mfm * v).rolling(n).sum() / v.rolling(n).sum().replace(0, np.nan)


def mfi(h, l, c, v, n: int = 14) -> pd.Series:
    tp = (h + l + c) / 3
    flow, d = tp * v, tp.diff()
    pos = flow.where(d > 0, 0.0).rolling(n).sum()
    neg = flow.where(d < 0, 0.0).rolling(n).sum()
    return 100 - 100 / (1 + pos / neg)


def vwap(h, l, c, v, n: int = 20) -> pd.Series:
    tp = (h + l + c) / 3
    return (tp * v).rolling(n).sum() / v.rolling(n).sum().replace(0, np.nan)


def obv(c, v) -> pd.Series:
    return (np.sign(c.diff()).fillna(0) * v).cumsum()


def last(s: pd.Series) -> float:
    return num(s.iloc[-1]) if len(s) else np.nan


def band(x, lo_bear, hi_bull, bull_lbl, bear_lbl):
    """+1 if x > hi_bull, -1 if x < lo_bear, else 0."""
    if x > hi_bull:
        return (1, bull_lbl)
    if x < lo_bear:
        return (-1, bear_lbl)
    return (0, "")


def technicals(df: pd.DataFrame, bench: pd.Series | None, today: pd.Timestamp):
    """Returns (metrics, votes) or (None, reason). Each vote is (+1 / 0 / -1, label)."""
    df = df[df["Close"] > 0]
    if len(df) < MIN_BARS:
        return None, "short history"
    if (today - naive_day(df.index[-1])).days > STALE_DAYS:
        return None, "stale price"

    c = df["Close"].astype(float)
    h = df["High"].fillna(df["Close"]).astype(float)
    l = df["Low"].fillna(df["Close"]).astype(float)
    v = df["Volume"].fillna(0).astype(float)
    n, price = len(c), float(c.iloc[-1])
    trend, mom, vol, caution = [], [], [], []

    # ---- Trend -------------------------------------------------------------
    sma50 = last(c.rolling(50).mean())
    sma200 = last(c.rolling(200).mean()) if n >= 200 else np.nan
    if ok(sma50):
        trend.append((1, "above 50d MA") if price > sma50 else (-1, "below 50d MA"))
    if ok(sma200):
        trend.append((1, "above 200d MA") if price > sma200 else (-1, "below 200d MA"))
        trend.append((1, "golden cross (50d > 200d)") if sma50 > sma200 else (-1, "death cross (50d < 200d)"))
        if n >= 220:
            s200 = c.rolling(200).mean()
            trend.append(band(pct(s200.iloc[-1], s200.iloc[-21]), -0.5, 0.5,
                              "200d MA rising", "200d MA falling"))
    ema20 = last(c.ewm(span=20, adjust=False).mean())
    ema50 = last(c.ewm(span=50, adjust=False).mean())
    trend.append((1, "20 EMA > 50 EMA") if ema20 > ema50 else (-1, "20 EMA < 50 EMA"))

    adx_s, pdi_s, mdi_s, atr_s = adx_dmi(h, l, c)
    adx_v, pdi, mdi, atr_v = last(adx_s), last(pdi_s), last(mdi_s), last(atr_s)
    if ok(adx_v):
        if adx_v >= 20 and pdi > mdi:
            trend.append((1, f"strong uptrend (ADX {adx_v:.0f})"))
        elif adx_v >= 20 and mdi > pdi:
            trend.append((-1, f"strong downtrend (ADX {adx_v:.0f})"))
        else:
            trend.append((0, ""))

    # ---- Momentum ----------------------------------------------------------
    rsi_v = last(rsi(c))
    if ok(rsi_v):
        if 50 < rsi_v < 70:
            mom.append((1, f"RSI bullish ({rsi_v:.0f})"))
        elif 30 < rsi_v <= 50:
            mom.append((-1, f"RSI weak ({rsi_v:.0f})"))
        else:
            mom.append((0, ""))
            caution.append(f"RSI {'overbought' if rsi_v >= 70 else 'oversold'} ({rsi_v:.0f})")

    m_line, m_sig = macd(c)
    ml, ms = last(m_line), last(m_sig)
    if ml > ms and ml > 0:
        mom.append((1, "MACD bullish"))
    elif ml < ms and ml < 0:
        mom.append((-1, "MACD bearish"))
    else:
        mom.append((0, ""))

    k_s, d_s = stochastic(h, l, c)
    stoch_k, stoch_d = last(k_s), last(d_s)
    if ok(stoch_k) and ok(stoch_d):
        if stoch_k >= 80:
            mom.append((0, ""))
            caution.append(f"Stochastic overbought ({stoch_k:.0f})")
        elif stoch_k <= 20:
            mom.append((0, ""))
            caution.append(f"Stochastic oversold ({stoch_k:.0f})")
        elif stoch_k > stoch_d:
            mom.append((1, f"Stochastic rising ({stoch_k:.0f})"))
        elif stoch_k < stoch_d:
            mom.append((-1, f"Stochastic falling ({stoch_k:.0f})"))
        else:
            mom.append((0, ""))

    cci_v = cci_last(h, l, c)
    if ok(cci_v):
        mom.append(band(cci_v, -100, 100, f"CCI strong ({cci_v:.0f})", f"CCI weak ({cci_v:.0f})"))

    ret_3m = pct(price, c.iloc[-64]) if n > 64 else np.nan
    if ok(ret_3m):
        mom.append(band(ret_3m, -2, 2, f"3M +{ret_3m:.0f}%", f"3M {ret_3m:.0f}%"))

    ret_12m = pct(price, c.iloc[-253]) if n > 253 else np.nan
    if n > 253:
        mom_lt, tag = pct(c.iloc[-22], c.iloc[-253]), "12-1M"      # classic 12-1 month momentum
    elif n > 127:
        mom_lt, tag = pct(price, c.iloc[-127]), "6M"
    else:
        mom_lt, tag = np.nan, ""
    if ok(mom_lt):
        mom.append(band(mom_lt, -5, 5, f"{tag} momentum +{mom_lt:.0f}%", f"{tag} momentum {mom_lt:.0f}%"))

    rel_6m = np.nan
    if bench is not None and n > 127:
        b = bench[bench.index <= df.index[-1]].dropna()
        if len(b) > 127:
            rel_6m = pct(price, c.iloc[-127]) - pct(b.iloc[-1], b.iloc[-127])
            mom.append(band(rel_6m, -5, 5, f"beating index by {rel_6m:.0f}pts (6M)",
                            f"lagging index by {abs(rel_6m):.0f}pts (6M)"))

    from_high = pct(price, float(h.iloc[-252:].max()))
    if from_high >= -10:
        mom.append((1, "near 52w high"))
    elif from_high <= -35:
        mom.append((-1, f"{abs(from_high):.0f}% below 52w high"))
    else:
        mom.append((0, ""))

    # ---- Volume ------------------------------------------------------------
    mfi_v = vol_ratio = np.nan
    if v.iloc[-60:].sum() > 0:
        cmf_v = last(cmf(h, l, c, v))
        if ok(cmf_v):
            vol.append(band(cmf_v, -0.05, 0.05, "buying pressure (CMF)", "selling pressure (CMF)"))
        o = obv(c, v)
        if len(o) > 21:
            vol.append((1, "OBV rising") if o.iloc[-1] > o.iloc[-21] else (-1, "OBV falling"))
        mfi_v = last(mfi(h, l, c, v))
        if ok(mfi_v):
            if mfi_v >= 80:
                vol.append((0, ""))
                caution.append(f"MFI overbought ({mfi_v:.0f})")
            elif mfi_v <= 20:
                vol.append((0, ""))
                caution.append(f"MFI oversold ({mfi_v:.0f})")
            else:
                vol.append((1, f"MFI {mfi_v:.0f}") if mfi_v > 50 else (-1, f"MFI {mfi_v:.0f}"))
        vw = last(vwap(h, l, c, v))
        if ok(vw):
            vol.append((1, "above 20d VWAP") if price > vw else (-1, "below 20d VWAP"))
        if n >= 70:
            base = v.iloc[-70:-20].mean()
            vol_ratio = v.iloc[-20:].mean() / base if base > 0 else np.nan
            chg20 = pct(price, c.iloc[-21])
            if ok(vol_ratio) and vol_ratio >= 1.3 and chg20 > 0:
                vol.append((1, f"volume surge on rise (x{vol_ratio:.1f})"))
            elif ok(vol_ratio) and vol_ratio >= 1.3 and chg20 < 0:
                vol.append((-1, f"volume surge on fall (x{vol_ratio:.1f})"))
            else:
                vol.append((0, ""))

    metrics = {
        "price": price,
        "last_date": naive_day(df.index[-1]).date(),
        "bars": n,
        "rsi": rsi_v, "stoch_k": stoch_k, "cci": cci_v, "mfi": mfi_v, "adx": adx_v,
        "atr_pct": atr_v / price * 100 if ok(atr_v) else np.nan,
        "vol_ratio": vol_ratio,
        "ret_3m": ret_3m, "ret_12m": ret_12m, "rel_6m": rel_6m, "from_high": from_high,
        "caution": caution,
    }
    return metrics, {"trend": trend, "momentum": mom, "volume": vol}


# =============================================================================
# 5. Fundamental scoring
# =============================================================================

VAL_METRICS = ["trailing_pe", "forward_pe", "pb", "ps", "ev_ebitda"]
VAL_LABELS = {"trailing_pe": "P/E", "forward_pe": "fwd P/E", "pb": "P/B", "ps": "P/S", "ev_ebitda": "EV/EBITDA"}


def build_peers(recs: list[dict]) -> dict:
    """Median valuation ratios per (market, sector), with a whole-market fallback."""
    df = pd.DataFrame(recs)
    for col in ["market", "sector", *VAL_METRICS]:
        if col not in df:
            df[col] = np.nan
    vals = df[VAL_METRICS].apply(pd.to_numeric, errors="coerce")
    vals = vals.where(vals > 0)                       # ignore negative / meaningless ratios
    df = pd.concat([df[["market", "sector"]], vals], axis=1)

    def med(s):
        return s.median() if s.count() >= SECTOR_MIN_PEERS else np.nan

    sector = df.dropna(subset=["sector"]).groupby(["market", "sector"])[VAL_METRICS].agg(med)
    market = df.groupby("market")[VAL_METRICS].agg(med)
    return {"sector": sector, "market": market}


def peer_median(peers: dict, rec: dict, metric: str):
    sec, mkt = peers["sector"], peers["market"]
    key = (rec.get("market"), rec.get("sector"))
    if ok(rec.get("sector")) and key in sec.index:
        v = sec.at[key, metric]
        if ok(v):
            return float(v), "sector"
    if rec.get("market") in mkt.index:
        v = mkt.at[rec["market"], metric]
        if ok(v):
            return float(v), "market"
    return np.nan, ""


def fundamental_votes(r: dict, peers: dict):
    V = {g: [] for g in FUND_WEIGHTS}
    caution = []
    fin = r.get("sector") in FINANCIAL_SECTORS
    g = lambda key: r.get(key) if ok(r.get(key)) else np.nan   # noqa: E731

    # ---- Valuation: ratios vs sector peers ----------------------------------
    for key in VAL_METRICS:
        if fin and key in ("ps", "ev_ebitda"):
            continue
        val, label = g(key), VAL_LABELS[key]
        if key == "ev_ebitda" and ok(val) and val < 0:
            V["valuation"].append((-1, "negative EBITDA"))
            continue
        if not ok(val) or val <= 0:
            continue                                   # loss-makers are caught under Quality
        med, scope = peer_median(peers, r, key)
        if ok(med):
            ratio = val / med
            if ratio <= CHEAP_VS_PEERS:
                V["valuation"].append((1, f"{label} {val:.1f} vs {scope} {med:.1f}"))
            elif ratio >= DEAR_VS_PEERS:
                V["valuation"].append((-1, f"{label} {val:.1f} vs {scope} {med:.1f}"))
            else:
                V["valuation"].append((0, ""))
        else:
            cheap, dear = ABS_VALUATION[key]
            V["valuation"].append((1, f"low {label} {val:.1f}") if val <= cheap else
                                  (-1, f"high {label} {val:.1f}") if val > dear else (0, ""))
    peg = g("peg")
    if ok(peg) and peg > 0:
        V["valuation"].append((1, f"PEG {peg:.1f}") if peg <= 1 else
                              (-1, f"PEG {peg:.1f}") if peg > 2.5 else (0, ""))
    dy, payout = g("div_yield_pct"), g("payout_ratio")
    if ok(dy) and dy >= 3 and (not ok(payout) or payout <= 0.8):
        V["valuation"].append((1, f"dividend yield {dy:.1f}%"))

    # ---- Quality ------------------------------------------------------------
    eps = g("eps_ttm")
    if ok(eps):
        V["quality"].append((1, "profitable") if eps > 0 else (-1, "loss-making"))
    if ok(g("roe")):
        V["quality"].append(band(g("roe"), 0, 0.15, f"ROE {g('roe')*100:.0f}%", "negative ROE"))
    if ok(g("roa")):
        V["quality"].append(band(g("roa"), 0, 0.05, f"ROA {g('roa')*100:.0f}%", "negative ROA"))
    if ok(g("op_margin")):
        V["quality"].append(band(g("op_margin"), 0, 0.15, f"op margin {g('op_margin')*100:.0f}%",
                                 "operating loss"))
    if ok(g("profit_margin")):
        V["quality"].append(band(g("profit_margin"), 0, 0.10, f"net margin {g('profit_margin')*100:.0f}%",
                                 "negative net margin"))

    # ---- Growth -------------------------------------------------------------
    if ok(g("revenue_growth")):
        rg = g("revenue_growth")
        V["growth"].append(band(rg, 0, 0.10, f"revenue +{rg*100:.0f}%", f"revenue {rg*100:.0f}%"))
    if ok(g("earnings_growth")):
        eg = g("earnings_growth")
        V["growth"].append(band(eg, 0, 0.10, f"earnings +{eg*100:.0f}%", f"earnings {eg*100:.0f}%"))
    eps_f = g("eps_fwd")
    if ok(eps) and ok(eps_f):
        if eps > 0:
            fg = (eps_f / eps - 1) * 100
            V["growth"].append(band(fg, -5, 10, f"EPS forecast +{fg:.0f}%", f"EPS forecast {fg:.0f}%"))
        elif eps_f > 0:
            V["growth"].append((1, "forecast return to profit"))
        elif eps_f < eps:
            V["growth"].append((-1, "losses forecast to widen"))
        else:
            V["growth"].append((0, ""))

    # ---- Financial health ---------------------------------------------------
    if not fin:
        de = g("debt_to_equity")
        if ok(de):
            V["health"].append((1, f"low debt (D/E {de/100:.1f}x)") if de < 50 else
                               (-1, f"high debt (D/E {de/100:.1f}x)") if de > 200 else (0, ""))
        cr = g("current_ratio")
        if ok(cr):
            V["health"].append((1, f"current ratio {cr:.1f}") if cr >= 1.5 else
                               (-1, f"current ratio {cr:.1f}") if cr < 1 else (0, ""))
    fy = g("fcf_yield_pct")
    if ok(fy):
        V["health"].append(band(fy, 0, 5, f"FCF yield {fy:.0f}%", "cash burn (negative FCF)"))
    elif ok(g("free_cash_flow")):
        V["health"].append((1, "positive free cash flow") if g("free_cash_flow") > 0 else
                           (-1, "cash burn (negative FCF)"))
    if ok(payout) and payout > 1.0:
        V["health"].append((-1, f"paying out {payout*100:.0f}% of earnings"))

    # ---- Sentiment ----------------------------------------------------------
    n_an = g("analyst_count")
    enough = not ok(n_an) or n_an >= 3
    a = g("analyst")
    if ok(a) and enough:
        V["sentiment"].append((1, f"analysts Buy ({a:.1f})") if a <= 2.0 else
                              (-1, f"analysts Hold/Sell ({a:.1f})") if a >= 3.0 else (0, ""))
    up = g("target_upside_pct")
    if ok(up) and enough:
        V["sentiment"].append(band(up, 0, 15, f"analyst target +{up:.0f}%", f"analyst target {up:.0f}%"))
    sf = g("short_float_pct")
    if ok(sf) and sf >= 10:
        caution.append(f"short interest {sf:.0f}% of float")
    return V, caution


# =============================================================================
# 6. Combine into a rating
# =============================================================================

def avg_vote(votes: list) -> float:
    vals = [v for v, _ in votes]
    return float(np.mean(vals)) if vals else np.nan


def weighted(groups: dict, weights: dict):
    means = {k: avg_vote(groups.get(k, [])) for k in weights}
    w = {k: weights[k] for k in weights if ok(means[k])}
    total = sum(means[k] * w[k] for k in w) / sum(w.values()) if w else np.nan
    return total, means


def to_rating(s: float) -> str:
    if not ok(s):
        return "No data"
    if s >= STRONG_BUY_AT:
        return "Strong Buy"
    if s >= BUY_AT:
        return "Buy"
    if s > SELL_AT:
        return "Neutral"
    if s > STRONG_SELL_AT:
        return "Sell"
    return "Strong Sell"


def score_record(rec: dict, tm: dict, tv: dict, peers: dict | None, fund_weight: float) -> dict:
    tech, tech_groups = weighted(tv, TECH_WEIGHTS)

    fv, fcaution = ({g: [] for g in FUND_WEIGHTS}, []) if peers is None else fundamental_votes(rec, peers)
    n_fund = sum(len(x) for x in fv.values())
    if n_fund >= MIN_FUND_SIGNALS:
        fund, fund_groups = weighted(fv, FUND_WEIGHTS)
    else:
        fund, fund_groups = np.nan, {g: np.nan for g in FUND_WEIGHTS}
    total = (1 - fund_weight) * tech + fund_weight * fund if ok(fund) else tech

    all_votes = [x for g in tv.values() for x in g] + [x for g in fv.values() for x in g]
    bull = [lbl for v, lbl in all_votes if v > 0 and lbl]
    bear = [lbl for v, lbl in all_votes if v < 0 and lbl]
    bear += [f"{c} - caution" for c in tm["caution"] + fcaution]

    out = dict(rec)
    out.update(tm)
    out.update(tech_groups)
    out.update(fund_groups)
    out.update({
        "technical": tech, "fundamentals": fund, "score": total, "rating": to_rating(total),
        "peer_pe": peer_median(peers, rec, "trailing_pe")[0] if peers is not None else np.nan,
        "bullish": ", ".join(bull), "bearish": ", ".join(bear), "signals_used": len(all_votes),
    })
    return out


# =============================================================================
# 7. Report
# =============================================================================

# (key, header, section) - section sets the header colour in Excel
REPORT_COLUMNS = [
    ("rank", "Rank", "id"), ("rating", "Rating", "id"), ("score", "Score", "id"),
    ("ticker", "Ticker", "id"), ("name", "Name", "id"), ("market_name", "Market", "id"),
    ("brokerage", "CMC Brokerage", "id"), ("exchange", "Exchange", "id"), ("sector", "Sector", "id"),
    ("industry", "Industry", "id"), ("currency", "Ccy", "id"), ("price", "Price", "id"),
    ("mcap_bn", "Mkt Cap (bn)", "id"),
    # scores (-100 .. +100)
    ("technical", "Technical", "score"), ("trend", "Trend", "score"), ("momentum", "Momentum", "score"),
    ("volume", "Volume", "score"), ("fundamentals", "Fundamentals", "score"),
    ("valuation", "Valuation", "score"), ("quality", "Quality", "score"), ("growth", "Growth", "score"),
    ("health", "Health", "score"), ("sentiment", "Sentiment", "score"),
    # fundamentals
    ("trailing_pe", "P/E", "fund"), ("peer_pe", "Peer P/E (median)", "fund"),
    ("forward_pe", "Fwd P/E", "fund"), ("peg", "PEG", "fund"), ("pb", "P/B", "fund"),
    ("ps", "P/S", "fund"), ("ev_ebitda", "EV/EBITDA", "fund"), ("div_yield_pct", "Div Yield %", "fund"),
    ("payout_pct", "Payout %", "fund"), ("roe_pct", "ROE %", "fund"), ("roa_pct", "ROA %", "fund"),
    ("op_margin_pct", "Op Margin %", "fund"), ("net_margin_pct", "Net Margin %", "fund"),
    ("rev_growth_pct", "Revenue Growth %", "fund"), ("earn_growth_pct", "Earnings Growth %", "fund"),
    ("de_x", "Debt/Equity (x)", "fund"), ("current_ratio", "Current Ratio", "fund"),
    ("fcf_yield_pct", "FCF Yield %", "fund"), ("analyst", "Analyst (1=SB, 5=SS)", "fund"),
    ("analyst_count", "# Analysts", "fund"), ("target_upside_pct", "Target Upside %", "fund"),
    ("short_float_pct", "Short % Float", "fund"),
    # technicals
    ("rsi", "RSI 14", "tech"), ("stoch_k", "Stoch %K", "tech"), ("cci", "CCI 20", "tech"),
    ("mfi", "MFI 14", "tech"), ("adx", "ADX 14", "tech"), ("atr_pct", "ATR %", "tech"),
    ("vol_ratio", "Volume 20d vs 50d", "tech"), ("ret_3m", "3M %", "tech"), ("ret_12m", "12M %", "tech"),
    ("rel_6m", "vs Index 6M (pts)", "tech"), ("from_high", "From 52W High %", "tech"),
    # reasons
    ("bullish", "Bullish signals", "why"), ("bearish", "Bearish signals", "why"),
    ("signals_used", "Signals used", "why"), ("fund_source", "Financials", "why"),
    ("last_date", "Price date", "why"),
]
SCORE_HEADERS = [h for _, h, s in REPORT_COLUMNS if s == "score"] + ["Score"]
ALWAYS_KEEP = {"Rank", "Rating", "Score", "Ticker", "Name", "Market", "Price", "Bullish signals", "Bearish signals"}

METHOD_ROWS = [
    ("TREND (40% of technical)", "Price vs 50-day MA", "Above", "Below", ""),
    ("", "Price vs 200-day MA", "Above", "Below", ""),
    ("", "50-day vs 200-day MA", "Golden cross (50 above 200)", "Death cross (50 below 200)", ""),
    ("", "200-day MA slope (20 days)", "Rising > 0.5%", "Falling > 0.5%", "Flat = 0"),
    ("", "20 EMA vs 50 EMA", "20 above 50", "20 below 50", ""),
    ("", "ADX / DMI (14)", "ADX >= 20 and +DI > -DI", "ADX >= 20 and -DI > +DI", "ADX < 20 (no trend) = 0"),
    ("MOMENTUM (40% of technical)", "RSI (14)", "50 - 70", "30 - 50", ">= 70 or <= 30 = 0, flagged as caution"),
    ("", "MACD (12,26,9)", "Above signal and above zero", "Below signal and below zero", "Mixed = 0"),
    ("", "Stochastic (14,3,3)", "%K above %D (20 - 80)", "%K below %D (20 - 80)", ">= 80 or <= 20 = 0, flagged as caution"),
    ("", "CCI (20)", "> +100", "< -100", ""),
    ("", "3-month return", "> +2%", "< -2%", ""),
    ("", "12-1 month momentum", "> +5%", "< -5%", "Uses 6M return if < 1 year of history"),
    ("", "Relative strength vs local index (6M)", "Beating by > 5 pts", "Lagging by > 5 pts", ""),
    ("", "Distance from 52-week high", "Within 10%", "More than 35% below", ""),
    ("VOLUME (20% of technical)", "Chaikin Money Flow (20)", "> +0.05", "< -0.05", ""),
    ("", "On-Balance Volume (20 days)", "Rising", "Falling", ""),
    ("", "Money Flow Index (14)", "50 - 80", "20 - 50", ">= 80 or <= 20 = 0, flagged as caution"),
    ("", "20-day VWAP", "Price above", "Price below", ""),
    ("", "Volume surge (20d avg vs prior 50d)", ">= 1.3x on a 20-day rise", ">= 1.3x on a 20-day fall", ""),
    ("VALUATION (25% of fundamental)", "P/E, forward P/E, P/B, P/S, EV/EBITDA",
     f"<= {CHEAP_VS_PEERS:.0%} of peer median", f">= {DEAR_VS_PEERS:.0%} of peer median",
     f"Peers = same sector in the same market (min {SECTOR_MIN_PEERS}), else whole market, "
     "else fixed cut-offs. P/S and EV/EBITDA skipped for financials. Negative EBITDA = bearish."),
    ("", "PEG", "<= 1.0", "> 2.5", ""),
    ("", "Dividend yield", ">= 3% with payout <= 80%", "", "No penalty for not paying dividends"),
    ("QUALITY (25% of fundamental)", "Profitability (EPS)", "Positive", "Negative", ""),
    ("", "Return on equity", "> 15%", "< 0%", ""),
    ("", "Return on assets", "> 5%", "< 0%", ""),
    ("", "Operating margin", "> 15%", "< 0%", ""),
    ("", "Net margin", "> 10%", "< 0%", ""),
    ("GROWTH (20% of fundamental)", "Revenue growth (YoY)", "> 10%", "< 0%", ""),
    ("", "Earnings growth (YoY)", "> 10%", "< 0%", ""),
    ("", "Forecast EPS vs trailing", "> +10% / return to profit", "< -5% / losses widening", ""),
    ("HEALTH (15% of fundamental)", "Debt / equity", "< 0.5x", "> 2.0x", "Skipped for financials"),
    ("", "Current ratio", ">= 1.5", "< 1.0", "Skipped for financials"),
    ("", "Free-cash-flow yield", "> 5%", "< 0% (cash burn)", "Sign only if reporting currency differs"),
    ("", "Payout ratio", "", "> 100% of earnings", ""),
    ("SENTIMENT (15% of fundamental)", "Analyst consensus (1-5)", "<= 2.0", ">= 3.0", "Needs 3+ analysts"),
    ("", "Analyst price-target upside", "> +15%", "< 0%", "Short interest >= 10% flagged as caution"),
    ("SCORING", "Each signal votes +1 / 0 / -1", "", "", "Group score = average vote in the group"),
    ("", "Technical score", "", "", "Trend 40%, Momentum 40%, Volume 20%"),
    ("", "Fundamental score", "", "", "Valuation 25%, Quality 25%, Growth 20%, Health 15%, Sentiment 15%. "
                                      f"Needs {MIN_FUND_SIGNALS}+ fundamental signals."),
    ("", "Final score", "", "", "Technical x (1 - F) + Fundamental x F, F = fundamentals weight "
                                  f"(default {DEFAULT_FUND_WEIGHT:.0%}). Technical only if fundamentals missing."),
    ("RATING", "Strong Buy", f">= {STRONG_BUY_AT*100:.0f}", "", ""),
    ("", "Buy", f"{BUY_AT*100:.0f} to {STRONG_BUY_AT*100:.0f}", "", ""),
    ("", "Neutral", f"{SELL_AT*100:.0f} to {BUY_AT*100:.0f}", "", ""),
    ("", "Sell", f"{STRONG_SELL_AT*100:.0f} to {SELL_AT*100:.0f}", "", ""),
    ("", "Strong Sell", f"<= {STRONG_SELL_AT*100:.0f}", "", ""),
    ("DATA", "Source", "Yahoo Finance via yfinance", "", "Unofficial, free, for personal use. Can contain gaps or errors."),
    ("", "Universe", "Main boards of the 16 CMC Invest markets", "", "Filtered by market cap / volume unless --no-filters"),
    ("", "Prices", "Daily, split & dividend adjusted, 2 years", "", "Stale (>10 days) or short-history (<60 days) stocks skipped"),
    ("", "Financials", f"Per-stock, cached {FUND_CACHE_DAYS} days", "",
     "'Full' = full financials; 'Quick' = only P/E, P/B, EPS & analyst rating from the stock list"),
    ("NOTE", "Not financial advice", "", "", "A mechanical summary of indicators. Use it to shortlist, then do your own research."),
]


def build_report_frame(results: list[dict]) -> pd.DataFrame:
    df = pd.DataFrame(results)
    for key, _, _ in REPORT_COLUMNS:
        if key not in df:
            df[key] = np.nan
    for raw in ("payout_ratio", "roe", "roa", "op_margin", "profit_margin", "revenue_growth",
                "earnings_growth", "debt_to_equity", "market_cap", "full"):
        if raw not in df:
            df[raw] = np.nan
    df["market_name"] = df["market"].map(lambda m: MARKETS[m].name if m in MARKETS else "Other")
    df["brokerage"] = df["market"].map(lambda m: MARKETS[m].brokerage if m in MARKETS else "")
    df["mcap_bn"] = pd.to_numeric(df["market_cap"], errors="coerce") / 1e9
    for src, dst in (("payout_ratio", "payout_pct"), ("roe", "roe_pct"), ("roa", "roa_pct"),
                     ("op_margin", "op_margin_pct"), ("profit_margin", "net_margin_pct"),
                     ("revenue_growth", "rev_growth_pct"), ("earnings_growth", "earn_growth_pct")):
        df[dst] = pd.to_numeric(df[src], errors="coerce") * 100
    df["de_x"] = pd.to_numeric(df["debt_to_equity"], errors="coerce") / 100
    df["fund_source"] = np.where(df["full"].fillna(False).astype(bool), "Full", "Quick")
    df.loc[df["fundamentals"].isna(), "fund_source"] = "Missing"

    score_keys = ["score"] + [k for k, _, s in REPORT_COLUMNS if s == "score"]
    for col in score_keys:
        df[col] = (pd.to_numeric(df[col], errors="coerce") * 100).round(0)
    two_dp = {"price", "mcap_bn", "analyst", "peg", "pb", "ps", "de_x", "current_ratio"}
    numeric = [k for k, _, s in REPORT_COLUMNS if s in ("fund", "tech")] + ["price", "mcap_bn"]
    for col in numeric:
        df[col] = pd.to_numeric(df[col], errors="coerce").round(2 if col in two_dp else 1)

    df = df.sort_values(["score", "fundamentals", "momentum"], ascending=False,
                        na_position="last").reset_index(drop=True)
    df["rank"] = np.arange(1, len(df) + 1)
    out = df[[k for k, _, _ in REPORT_COLUMNS]].rename(columns={k: h for k, h, _ in REPORT_COLUMNS})
    empty = [c for c in out.columns if c not in ALWAYS_KEEP and out[c].isna().all()]
    return out.drop(columns=empty)


# =============================================================================
# 7b. Top picks and plain-English explanations
# =============================================================================

GROUP_PHRASES = {
    #  report column: (bullish phrase, bearish phrase)
    "Trend": ("in a strong uptrend", "in a downtrend"),
    "Momentum": ("strong price momentum", "losing momentum"),
    "Volume": ("steady buying volume", "heavy selling volume"),
    "Valuation": ("cheap compared with its peers", "expensive compared with its peers"),
    "Quality": ("highly profitable", "unprofitable or low quality"),
    "Growth": ("growing sales and earnings", "shrinking sales and earnings"),
    "Health": ("a solid balance sheet", "a stretched balance sheet"),
    "Sentiment": ("analysts are bullish", "analysts are bearish"),
}


def _v(r, col):
    v = r.get(col)
    if v is None:
        return None
    try:
        f = float(v)
    except (TypeError, ValueError):
        return None
    return None if math.isnan(f) else f


def _pc(x, d=0):
    return f"{x:+.{d}f}%".replace("+-", "-")


def _join(parts):
    parts = [p for p in parts if p]
    if not parts:
        return ""
    return parts[0] if len(parts) == 1 else ", ".join(parts[:-1]) + " and " + parts[-1]


def headline(r) -> str:
    score = _v(r, "Score") or 0
    d = 1 if score >= 0 else -1
    groups = [(g, _v(r, g)) for g in GROUP_PHRASES]
    lead = sorted([(g, v) for g, v in groups if v is not None and v * d >= 40], key=lambda x: -abs(x[1]))[:3]
    if not lead:
        return "Mixed signals - no strong case either way"
    words = [GROUP_PHRASES[g][0 if d > 0 else 1] for g, _ in lead]
    text = _join(words)
    return text[0].upper() + text[1:]


def explain(r) -> str:
    """A short paragraph, in plain English, on why the stock has its rating."""
    out = []
    score, rating = _v(r, "Score"), r.get("Rating", "")
    tech, fund = _v(r, "Technical"), _v(r, "Fundamentals")

    # 1. Price trend and momentum
    trend = _v(r, "Trend")
    r3, r12, rel, hi = _v(r, "3M %"), _v(r, "12M %"), _v(r, "vs Index 6M (pts)"), _v(r, "From 52W High %")
    moves = []
    if r3 is not None:
        moves.append(f"{_pc(r3)} over 3 months")
    if r12 is not None:
        moves.append(f"{_pc(r12)} over 12 months")
    s = ""
    if trend is not None and trend >= 50:
        s = "The share price is in a clear uptrend, above its 50- and 200-day averages"
    elif trend is not None and trend <= -50:
        s = "The share price is in a clear downtrend, below its 50- and 200-day averages"
    elif trend is not None:
        s = "The price trend is mixed"
    if moves:
        s = (s + f" ({_join(moves)})") if s else "The share price is " + _join(moves)
    if s:
        if rel is not None and abs(rel) >= 5:
            s += f", {'beating' if rel > 0 else 'lagging'} its market index by {abs(rel):.0f} points over 6 months"
        if hi is not None and hi > -5:
            s += ", and it is trading near its 52-week high"
        elif hi is not None and hi <= -35:
            s += f", and it is {abs(hi):.0f}% below its 52-week high"
        out.append(s + ".")

    # 2. Valuation
    pe, peer, fpe, dy = _v(r, "P/E"), _v(r, "Peer P/E (median)"), _v(r, "Fwd P/E"), _v(r, "Div Yield %")
    if pe is not None and pe > 0 and peer:
        diff = (pe / peer - 1) * 100
        if abs(diff) < 10:
            s = f"It is valued about in line with its sector, on a P/E of {pe:.1f} (peer median {peer:.1f})"
        else:
            s = (f"On a P/E of {pe:.1f} it is {abs(diff):.0f}% {'cheaper' if diff < 0 else 'more expensive'} "
                 f"than its sector peers ({peer:.1f})")
        if fpe and fpe > 0 and abs(fpe - pe) / pe > 0.1:
            s += f", falling to {fpe:.1f} on next year's forecast earnings" if fpe < pe else f", rising to {fpe:.1f} on forecast earnings"
        if dy and dy >= 3:
            s += f", and it pays a {dy:.1f}% dividend"
        out.append(s + ".")
    elif (_v(r, "Net Margin %") or 0) < 0 or (_v(r, "ROE %") or 0) < 0:
        out.append("It is currently losing money, so it has no P/E ratio.")
    if not (pe is not None and pe > 0 and peer):
        ps, pb = _v(r, "P/S"), _v(r, "P/B")
        vals = [f"{ps:.1f} times its annual sales" if ps else "", f"{pb:.1f} times its book value" if pb else ""]
        if any(vals):
            out.append("It is priced at " + _join(vals) + ".")

    # 3. Business quality and growth
    bits = []
    roe, nm = _v(r, "ROE %"), _v(r, "Net Margin %")
    rg, eg = _v(r, "Revenue Growth %"), _v(r, "Earnings Growth %")
    if roe is not None:
        bits.append(f"return on equity of {roe:.0f}%")
    if nm is not None and abs(nm) >= 0.5:
        bits.append(f"a net profit margin of {nm:.0f}%")
    if rg is not None:
        bits.append(f"revenue {'up' if rg >= 0 else 'down'} {abs(rg):.0f}%")
    if eg is not None:
        bits.append(f"earnings {'up' if eg >= 0 else 'down'} {abs(eg):.0f}% on last year")
    if bits:
        out.append("The business has " + _join(bits) + ".")

    # 4. Balance sheet
    de, fy = _v(r, "Debt/Equity (x)"), _v(r, "FCF Yield %")
    good, bad = [], []
    if de is not None:
        word = "low debt" if de < 0.5 else "high debt" if de > 2 else "moderate debt"
        (bad if de > 2 else good).append(f"{word} (debt/equity {de:.1f}x)")
    if fy is not None:
        if fy >= 0:
            good.append(f"free cash flow of {fy:.0f}% of its market value")
        else:
            bad.append("negative free cash flow (it is burning cash)")
    cr = _v(r, "Current Ratio")
    if cr is not None and not good and not bad:
        (good if cr >= 1.5 else bad if cr < 1 else good).append(f"a current ratio of {cr:.1f} (short-term assets vs short-term bills)")
    if good and bad:
        out.append("It has " + _join(good) + ", but " + _join(bad) + ".")
    elif good or bad:
        out.append("It has " + _join(good or bad) + ".")

    # 5. Analysts
    an, n_an, tgt = _v(r, "Analyst (1=SB, 5=SS)"), _v(r, "# Analysts"), _v(r, "Target Upside %")
    if an is not None and (n_an is None or n_an >= 3):
        word = "Strong Buy" if an <= 1.5 else "Buy" if an <= 2.5 else "Hold" if an <= 3.5 else "Sell"
        s = f"{int(n_an) if n_an else 'Covering'} analysts rate it {word} on average"
        if tgt is not None:
            s += f", with a price target {abs(tgt):.0f}% {'above' if tgt >= 0 else 'below'} the current price"
        out.append(s + ".")

    # 6. What to watch (the counter-case)
    watch = []
    rsi, short = _v(r, "RSI 14"), _v(r, "Short % Float")
    if rsi is not None and rsi >= 70:
        watch.append(f"the RSI is {rsi:.0f}, so it looks overbought and could pull back short term")
    if rsi is not None and rsi <= 30:
        watch.append(f"the RSI is {rsi:.0f}, so it looks oversold and could bounce short term")
    if short is not None and short >= 10:
        watch.append(f"{short:.0f}% of its shares are sold short")
    if score is not None and score >= 0 and fund is not None and fund < 0:
        watch.append("the rating rests on the chart; the company's numbers are weak")
    if score is not None and score >= 0 and tech is not None and tech < 0:
        watch.append("the rating rests on the company's numbers; the chart is weak")
    if score is not None and score < 0 and fund is not None and fund > 30:
        watch.append("the company's numbers are reasonable; the rating is mostly the weak chart")
    if r.get("Financials") != "Full":
        watch.append("full company financials were not available, so the fundamental view is limited")
    if watch:
        out.append("Worth watching: " + "; ".join(watch) + ".")
    return " ".join(out)


def mark_top_picks(report: pd.DataFrame, top_n: int) -> pd.DataFrame:
    """Flag at most top_n buys and top_n sells where technicals and fundamentals strongly agree.
    Picks are spread across markets (round robin by rank) so one market can't take every slot."""
    df = report.copy()
    df["Top Pick"] = ""
    if top_n <= 0 or "Fundamentals" not in df:
        return df
    mcap = pd.to_numeric(df.get("Mkt Cap (bn)"), errors="coerce").fillna(0)
    full = df.get("Financials", pd.Series("", index=df.index)) == "Full"
    tech, fund = df["Technical"], df["Fundamentals"]
    sides = {
        "Top Buy": (df["Rating"] == "Strong Buy") & (tech >= 40) & (fund >= 40) & full & (mcap >= 0.3),
        "Top Sell": (df["Rating"] == "Strong Sell") & (tech <= -40) & (fund <= -40) & full & (mcap >= 0.3),
    }
    for label, ok_mask in sides.items():
        cand = df[ok_mask].sort_values(["Score", "Fundamentals"], ascending=label == "Top Sell")
        queues = [list(g.index) for _, g in cand.groupby("Market", sort=False)]
        queues.sort(key=lambda q: cand.index.get_loc(q[0]))
        chosen = []
        while len(chosen) < top_n and any(queues):
            for q in queues:
                if q and len(chosen) < top_n:
                    chosen.append(q.pop(0))
        df.loc[chosen, "Top Pick"] = label
    return df


def add_explanations(report: pd.DataFrame, top_n: int) -> pd.DataFrame:
    df = mark_top_picks(report, top_n)
    recs = df.to_dict("records")
    df["Headline"] = [headline(r) for r in recs]
    df["Why"] = [explain(r) for r in recs]
    # keep the readable columns right after the name
    cols = list(df.columns)
    for c in ["Why", "Headline", "Top Pick"]:
        cols.remove(c)
        cols.insert(cols.index("Name") + 1, c)
    return df[cols]


def write_excel(report: pd.DataFrame, path: Path, run_note: str) -> None:
    from openpyxl.formatting.rule import ColorScaleRule, FormulaRule
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    summary = (report.pivot_table(index="Market", columns="Rating", values="Ticker",
                                  aggfunc="count", fill_value=0)
               .reindex(columns=[r for r in RATING_ORDER if r in set(report["Rating"])], fill_value=0))
    summary["Total"] = summary.sum(axis=1)
    if "P/E" in report:
        summary["Median P/E"] = report.groupby("Market")["P/E"].median().round(1)
    summary.loc["ALL MARKETS"] = summary.sum(numeric_only=True)
    if "Median P/E" in summary:
        summary.loc["ALL MARKETS", "Median P/E"] = round(report["P/E"].median(), 1)

    method = pd.DataFrame(METHOD_ROWS, columns=["Section", "Signal", "Bullish (+1)", "Bearish (-1)", "Notes"])
    sheets = {
        "All Ratings": report,
        "Top Picks": pd.concat([report[report.get("Top Pick", "") == "Top Buy"],
                               report[report.get("Top Pick", "") == "Top Sell"].iloc[::-1]]),
        "Strong Buy": report[report["Rating"] == "Strong Buy"],
        "Strong Sell": report[report["Rating"] == "Strong Sell"].iloc[::-1],
    }

    section_of = {h: s for _, h, s in REPORT_COLUMNS}
    section_fill = {"id": "1F2A44", "score": "4B3A8C", "fund": "0F5E5A", "tech": "7A4B12", "why": "4A4A4A"}
    fills = {
        "Strong Buy": ("1E7B3C", "FFFFFF"), "Buy": ("C6EFCE", "006100"), "Neutral": ("EDEDED", "3A3A3A"),
        "Sell": ("FFC7CE", "9C0006"), "Strong Sell": ("9C0006", "FFFFFF"),
    }
    header_font = Font(bold=True, color="FFFFFF")
    widths = {"Rank": 6, "Rating": 12, "Score": 7, "Ticker": 11, "Name": 32, "Market": 14,
              "CMC Brokerage": 22, "Exchange": 14, "Sector": 20, "Industry": 26, "Ccy": 6, "Price": 10,
              "Bullish signals": 90, "Bearish signals": 70, "Top Pick": 10, "Headline": 40, "Why": 120, "Price date": 11, "Financials": 10}

    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        for name, frame in sheets.items():
            frame.to_excel(xw, sheet_name=name, index=False)
        summary.to_excel(xw, sheet_name="Summary")
        method.to_excel(xw, sheet_name="How It Works", index=False)

        wb = xw.book
        for name, frame in sheets.items():
            ws = wb[name]
            cols = list(frame.columns)
            nrows = len(frame) + 1
            for j, col in enumerate(cols, 1):
                cell = ws.cell(row=1, column=j)
                cell.fill = PatternFill("solid", fgColor=section_fill[section_of.get(col, "id")])
                cell.font = header_font
                cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
                ws.column_dimensions[get_column_letter(j)].width = widths.get(col, 10)
            ws.row_dimensions[1].height = 34
            ws.freeze_panes = "E2"
            ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{max(nrows, 2)}"
            if nrows < 2:
                continue
            rcol = get_column_letter(cols.index("Rating") + 1)
            for rating, (bg, fg) in fills.items():
                ws.conditional_formatting.add(
                    f"{rcol}2:{rcol}{nrows}",
                    FormulaRule(formula=[f'${rcol}2="{rating}"'],
                                fill=PatternFill("solid", fgColor=bg, bgColor=bg),
                                font=Font(bold=True, color=fg)))
            for col in SCORE_HEADERS:
                if col in cols:
                    L = get_column_letter(cols.index(col) + 1)
                    ws.conditional_formatting.add(
                        f"{L}2:{L}{nrows}",
                        ColorScaleRule(start_type="num", start_value=-100, start_color="F8696B",
                                       mid_type="num", mid_value=0, mid_color="FFFFFF",
                                       end_type="num", end_value=100, end_color="63BE7B"))
            if "Mkt Cap (bn)" in cols:
                L = get_column_letter(cols.index("Mkt Cap (bn)") + 1)
                for r in range(2, nrows + 1):
                    ws[f"{L}{r}"].number_format = "#,##0.00"

        ws = wb["Summary"]
        ws.column_dimensions["A"].width = 18
        for j in range(1, ws.max_column + 1):
            c = ws.cell(row=1, column=j)
            c.fill, c.font = PatternFill("solid", fgColor=section_fill["id"]), header_font
            if j > 1:
                ws.column_dimensions[get_column_letter(j)].width = 12
        last_row = ws.max_row
        for j in range(1, ws.max_column + 1):
            ws.cell(row=last_row, column=j).font = Font(bold=True)
        ws.cell(row=last_row + 2, column=1, value=run_note)
        ws.cell(row=last_row + 3, column=1,
                value="Not financial advice - a mechanical summary of indicators. Do your own research.")

        ws = wb["How It Works"]
        for j, w in enumerate([30, 38, 34, 30, 80], 1):
            ws.column_dimensions[get_column_letter(j)].width = w
            c = ws.cell(row=1, column=j)
            c.fill, c.font = PatternFill("solid", fgColor=section_fill["id"]), header_font
        for row in ws.iter_rows(min_row=2):
            row[0].font = Font(bold=True)
            for c in row:
                c.alignment = Alignment(wrap_text=True, vertical="top")

        wb.move_sheet("Summary", offset=1 - wb.sheetnames.index("Summary"))   # Summary second


# =============================================================================
# 8. Web screener page (one self-contained index.html, ready to deploy)
# =============================================================================

# Report column -> short key used by the page
WEB_KEYS = {
    "Rank": "rank", "Rating": "rating", "Score": "score", "Ticker": "ticker", "Name": "name",
    "Market": "market", "CMC Brokerage": "brokerage", "Exchange": "exchange", "Sector": "sector",
    "Industry": "industry", "Ccy": "ccy", "Price": "price", "Mkt Cap (bn)": "mcap",
    "Technical": "tech", "Trend": "trend", "Momentum": "mom", "Volume": "vol",
    "Fundamentals": "fund", "Valuation": "val", "Quality": "qual", "Growth": "grow",
    "Health": "health", "Sentiment": "sent", "P/E": "pe", "Peer P/E (median)": "peer_pe",
    "Fwd P/E": "fpe", "PEG": "peg", "P/B": "pb", "P/S": "ps", "EV/EBITDA": "ev",
    "Div Yield %": "dy", "Payout %": "payout", "ROE %": "roe", "ROA %": "roa",
    "Op Margin %": "opm", "Net Margin %": "npm", "Revenue Growth %": "revg",
    "Earnings Growth %": "earng", "Debt/Equity (x)": "de", "Current Ratio": "cr",
    "FCF Yield %": "fcfy", "Analyst (1=SB, 5=SS)": "an", "# Analysts": "an_n",
    "Target Upside %": "tgt", "Short % Float": "short", "RSI 14": "rsi", "Stoch %K": "stoch",
    "CCI 20": "cci", "MFI 14": "mfi", "ADX 14": "adx", "ATR %": "atr",
    "Volume 20d vs 50d": "volr", "3M %": "r3m", "12M %": "r12m", "vs Index 6M (pts)": "rel6m",
    "From 52W High %": "hi52", "Bullish signals": "bull", "Bearish signals": "bear",
    "Financials": "fin", "Price date": "date",
    "Top Pick": "pick", "Headline": "head", "Why": "why",
}


def _web_value(v):
    if v is None:
        return None
    if isinstance(v, (bool, np.bool_)):
        return bool(v)
    if isinstance(v, (int, np.integer)):
        return int(v)
    if isinstance(v, (float, np.floating)):
        f = float(v)
        if not math.isfinite(f):
            return None
        return int(f) if f.is_integer() else round(f, 4)
    if isinstance(v, str):
        return v
    return str(v)


def write_site(report: pd.DataFrame, out_dir: Path, meta: dict) -> Path:
    cols = [(h, k) for h, k in WEB_KEYS.items() if h in report.columns]
    rows = [[_web_value(v) for v in rec]
            for rec in report[[h for h, _ in cols]].itertuples(index=False, name=None)]
    payload = json.dumps({"meta": meta, "cols": [k for _, k in cols], "rows": rows},
                         separators=(",", ":"), ensure_ascii=False).replace("</", "<\\/")
    site = out_dir / "site"
    site.mkdir(parents=True, exist_ok=True)
    path = site / "index.html"
    path.write_text(SCREENER_HTML.replace("__SCREENER_DATA__", payload, 1), encoding="utf-8")
    return path


SCREENER_HTML = r"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>Daily Stock Ratings</title>
<meta name="description" content="ASX and international stocks rated from Strong Buy to Strong Sell on technicals and fundamentals.">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Archivo:wdth,wght@62..125,500..800&family=IBM+Plex+Mono:wght@400;500;600&family=Public+Sans:wght@400;500;600;700&display=swap">
<style>
/* Layout: a control deck (rating ladder + filters) above a ranked list - a sortable table on desktop, tappable cards on phones - with a detail sheet per stock */
:root{
  --bg:#F1F4F2; --surface:#FFFFFF; --surface-2:#F6F8F7; --ink:#13201A; --muted:#5B6A63; --line:#D6DED9;
  --accent:#2343BF; --accent-soft:#E6EBFA; --on-accent:#FFFFFF;
  --sb:#0E7543; --b:#3A9E66; --n:#7F8B85; --s:#CF6553; --ss:#AD2419;
  --sb-bg:#DAF0E2; --b-bg:#E8F5EC; --n-bg:#EAEEEC; --s-bg:#FBE8E3; --ss-bg:#F7DAD6;
  --pos:#0E7543; --neg:#AD2419; --warn:#8A5E00; --warn-bg:#FBF1DA;
  --shadow:0 1px 2px rgb(19 32 26 / .05), 0 10px 28px rgb(19 32 26 / .07);
  --font-display:"Archivo","Arial Narrow",system-ui,sans-serif;
  --font-body:"Public Sans",system-ui,-apple-system,"Segoe UI",Roboto,sans-serif;
  --font-data:"IBM Plex Mono",ui-monospace,SFMono-Regular,Menlo,Consolas,monospace;
  --r:12px;
  color-scheme:light;
}
@media (prefers-color-scheme: dark){
  :root:not([data-theme="light"]){
    --bg:#0C1210; --surface:#131B18; --surface-2:#18221E; --ink:#E3EBE7; --muted:#8C9C95; --line:#25312C;
    --accent:#8DA4FF; --accent-soft:#1B2340; --on-accent:#0C1210;
    --sb:#3CC981; --b:#7AD2A1; --n:#808C86; --s:#EE8E7D; --ss:#FF6A5C;
    --sb-bg:#11331F; --b-bg:#15281D; --n-bg:#1B2420; --s-bg:#341C17; --ss-bg:#3E1511;
    --pos:#3CC981; --neg:#FF6A5C; --warn:#E5B44A; --warn-bg:#2E2512;
    --shadow:0 1px 2px rgb(0 0 0 / .3), 0 10px 28px rgb(0 0 0 / .35);
    color-scheme:dark;
  }
}
:root[data-theme="dark"]{
  --bg:#0C1210; --surface:#131B18; --surface-2:#18221E; --ink:#E3EBE7; --muted:#8C9C95; --line:#25312C;
  --accent:#8DA4FF; --accent-soft:#1B2340; --on-accent:#0C1210;
  --sb:#3CC981; --b:#7AD2A1; --n:#808C86; --s:#EE8E7D; --ss:#FF6A5C;
  --sb-bg:#11331F; --b-bg:#15281D; --n-bg:#1B2420; --s-bg:#341C17; --ss-bg:#3E1511;
  --pos:#3CC981; --neg:#FF6A5C; --warn:#E5B44A; --warn-bg:#2E2512;
  --shadow:0 1px 2px rgb(0 0 0 / .3), 0 10px 28px rgb(0 0 0 / .35);
  color-scheme:dark;
}

*,*::before,*::after{box-sizing:border-box}
html{-webkit-text-size-adjust:100%}
body{margin:0;background:var(--bg);color:var(--ink);font:15px/1.5 var(--font-body);
  padding-top:env(safe-area-inset-top,0px);padding-bottom:env(safe-area-inset-bottom,0px)}
[hidden]{display:none!important}
button,input,select{font:inherit;color:inherit}
:focus-visible{outline:2px solid var(--accent);outline-offset:2px}
.wrap{max-width:1320px;margin:0 auto;padding-left:max(16px,env(safe-area-inset-left,0px));padding-right:max(16px,env(safe-area-inset-right,0px))}
.num{font-family:var(--font-data);font-variant-numeric:tabular-nums}
.pos{color:var(--pos)} .neg{color:var(--neg)} .muted{color:var(--muted)}

/* ---- header ---- */
.top{padding-block:22px 6px}
.top-inner{display:flex;align-items:flex-start;justify-content:space-between;gap:16px}
h1{font-family:var(--font-display);font-stretch:78%;font-weight:800;font-size:clamp(1.7rem,4.2vw,2.4rem);
  line-height:1.05;letter-spacing:-.01em;margin:0;text-wrap:balance}
h1 .mark{color:var(--accent)}
.runline{margin:6px 0 0;color:var(--muted);font-size:.86rem}
.icon-btn{flex:none;width:40px;height:40px;border-radius:999px;border:1px solid var(--line);background:var(--surface);
  display:grid;place-items:center;cursor:pointer}
.icon-btn:hover{border-color:var(--accent)}
.icon-btn svg{width:18px;height:18px}
.banner{margin-top:14px;padding:10px 14px;border-radius:10px;background:var(--warn-bg);color:var(--warn);font-size:.88rem;font-weight:600}

/* ---- control deck ---- */
.deck{margin-top:16px;background:var(--surface);border:1px solid var(--line);border-radius:16px;padding:16px;display:grid;gap:14px;box-shadow:var(--shadow)}
.deck-head{display:flex;align-items:baseline;justify-content:space-between;gap:12px;flex-wrap:wrap}
.label{font-size:.72rem;letter-spacing:.08em;text-transform:uppercase;font-weight:700;color:var(--muted)}
.presets{display:flex;gap:6px;flex-wrap:wrap}
.preset{border:1px solid var(--line);background:var(--surface-2);border-radius:999px;padding:4px 12px;font-size:.82rem;font-weight:600;cursor:pointer}
.preset[aria-pressed="true"]{background:var(--accent);border-color:var(--accent);color:var(--on-accent)}
.ladder-bar{display:flex;height:10px;border-radius:999px;overflow:hidden;background:var(--surface-2);gap:2px}
.ladder-bar span{min-width:0;transition:opacity .2s, flex-grow .3s}
.ladder-bar span.off{opacity:.22}
.ladder-keys{display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:6px}
.lk{border:1px solid var(--line);background:var(--surface);border-radius:10px;padding:8px 8px 7px;text-align:left;cursor:pointer;
  display:grid;gap:2px;min-width:0}
.lk .lk-n{font-family:var(--font-data);font-variant-numeric:tabular-nums;font-weight:600;font-size:1.05rem;line-height:1.1}
.lk .lk-t{font-size:.76rem;font-weight:600;color:var(--muted);line-height:1.2}
.lk[aria-pressed="true"]{border-width:1px}
.lk.sb[aria-pressed="true"]{background:var(--sb-bg);border-color:var(--sb)} .lk.sb .lk-n{color:var(--sb)}
.lk.b[aria-pressed="true"]{background:var(--b-bg);border-color:var(--b)} .lk.b .lk-n{color:var(--b)}
.lk.n[aria-pressed="true"]{background:var(--n-bg);border-color:var(--n)} .lk.n .lk-n{color:var(--n)}
.lk.s[aria-pressed="true"]{background:var(--s-bg);border-color:var(--s)} .lk.s .lk-n{color:var(--s)}
.lk.ss[aria-pressed="true"]{background:var(--ss-bg);border-color:var(--ss)} .lk.ss .lk-n{color:var(--ss)}
.lk[aria-pressed="false"] .lk-n{color:var(--muted)}
.lk[aria-pressed="true"] .lk-t{color:var(--ink)}

.filters{display:grid;grid-template-columns:minmax(0,2fr) repeat(3,minmax(0,1fr)) auto;gap:8px}
.field{display:grid;gap:4px;min-width:0}
.field label{font-size:.74rem;font-weight:600;color:var(--muted)}
.ctl{width:100%;min-width:0;height:40px;border:1px solid var(--line);border-radius:10px;background:var(--surface-2);padding:0 10px}
.ctl:focus{border-color:var(--accent);outline:none;box-shadow:0 0 0 3px var(--accent-soft)}
.more-btn{align-self:end;height:40px;border:1px solid var(--line);border-radius:10px;background:var(--surface);padding:0 14px;font-weight:600;cursor:pointer;white-space:nowrap}
.more-btn[aria-expanded="true"]{border-color:var(--accent);color:var(--accent)}
.more{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px 12px;align-items:end;
  border-top:1px dashed var(--line);padding-top:12px}
.check{display:flex;align-items:center;gap:8px;font-size:.88rem;min-height:40px}
.check input{width:18px;height:18px;accent-color:var(--accent)}
.text-btn{background:none;border:none;color:var(--accent);font-weight:600;cursor:pointer;padding:8px 0;justify-self:start}

/* ---- results ---- */
.resultbar{display:flex;justify-content:space-between;align-items:baseline;gap:12px;flex-wrap:wrap;margin:22px 0 10px}
.resultbar h2{font-family:var(--font-display);font-stretch:85%;font-weight:750;font-size:1.25rem;margin:0}
.resultbar .count{color:var(--muted);font-size:.88rem}

.chip{display:inline-flex;align-items:center;gap:6px;border-radius:999px;padding:2px 10px 2px 8px;font-size:.78rem;font-weight:700;white-space:nowrap}
.chip::before{content:"";width:7px;height:7px;border-radius:50%;background:currentColor}
.chip.sb{background:var(--sb-bg);color:var(--sb)} .chip.b{background:var(--b-bg);color:var(--b)}
.chip.n{background:var(--n-bg);color:var(--n)} .chip.s{background:var(--s-bg);color:var(--s)} .chip.ss{background:var(--ss-bg);color:var(--ss)}
.score-sb{color:var(--sb)} .score-b{color:var(--b)} .score-n{color:var(--n)} .score-s{color:var(--s)} .score-ss{color:var(--ss)}

.meter{position:relative;height:6px;border-radius:999px;background:var(--surface-2);overflow:hidden}
.meter::after{content:"";position:absolute;left:50%;top:-2px;bottom:-2px;width:1px;background:var(--muted);opacity:.5}
.meter i{position:absolute;top:0;bottom:0;border-radius:999px}
.meter i.up{background:var(--pos)} .meter i.dn{background:var(--neg)}

/* cards (phones) */
.cards{display:grid;gap:10px}
.card{display:grid;grid-template-columns:minmax(0,1fr) auto;gap:4px 14px;width:100%;text-align:left;cursor:pointer;
  background:var(--surface);border:1px solid var(--line);border-radius:var(--r);padding:14px}
.card:hover{border-color:var(--accent)}
.card .c-top{display:flex;align-items:center;gap:8px;min-width:0}
.card .rank{font-family:var(--font-data);font-size:.78rem;color:var(--muted)}
.card .tk{display:block;font-family:var(--font-display);font-stretch:85%;font-weight:750;font-size:1.2rem;line-height:1.15;margin-top:6px}
.card .nm{display:block;color:var(--ink);font-size:.9rem;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.card .mt{display:block;color:var(--muted);font-size:.8rem;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.card .sc{grid-column:2;grid-row:1 / span 4;display:grid;align-content:start;justify-items:end;gap:6px;min-width:72px}
.card .sc b{font-family:var(--font-data);font-size:1.7rem;line-height:1;font-weight:600}
.card .sc .meter{width:72px}
.card .sc small{font-size:.68rem;color:var(--muted);letter-spacing:.06em;text-transform:uppercase}
.card .st{grid-column:1 / -1;display:flex;flex-wrap:wrap;gap:4px 16px;margin-top:8px;padding-top:8px;border-top:1px solid var(--line);font-size:.8rem}
.card .st span{white-space:nowrap}
.card .st em{font-style:normal;color:var(--muted);margin-right:4px;font-family:var(--font-body)}

/* table (desktop) */
.table-wrap{overflow-x:auto;background:var(--surface);border:1px solid var(--line);border-radius:14px}
table{width:100%;border-collapse:collapse;font-size:.88rem}
th,td{padding:8px 8px;text-align:right;white-space:nowrap;border-bottom:1px solid var(--line)}
th{position:relative;font-size:.72rem;letter-spacing:.06em;text-transform:uppercase;color:var(--muted);font-weight:700;background:var(--surface-2)}
th.l,td.l{text-align:left}
th button{all:unset;cursor:pointer;display:inline-flex;gap:4px;align-items:center}
th button:focus-visible{outline:2px solid var(--accent);outline-offset:2px}
th[aria-sort="descending"] button::after{content:"\25BE"} th[aria-sort="ascending"] button::after{content:"\25B4"}
th[aria-sort] {color:var(--accent)}
tbody tr{cursor:pointer}
tbody tr:hover td{background:var(--surface-2)}
tbody tr:last-child td{border-bottom:none}
td.tkc button{all:unset;cursor:pointer;font-family:var(--font-display);font-stretch:85%;font-weight:750;font-size:1rem}
td.tkc button:focus-visible{outline:2px solid var(--accent)}
td.co{max-width:180px;overflow:hidden;text-overflow:ellipsis}
td.sec{max-width:128px;overflow:hidden;text-overflow:ellipsis}
td.scr b{font-family:var(--font-data);font-weight:600;font-size:1rem}
td.scr .meter{width:44px;display:inline-block;vertical-align:middle;margin-left:8px}

.more-row{display:flex;justify-content:center;margin:16px 0 8px}
.btn{border:1px solid var(--accent);background:var(--accent);color:var(--on-accent);border-radius:10px;padding:10px 18px;font-weight:700;cursor:pointer}
.btn.ghost{background:var(--surface);color:var(--accent)}
.empty{padding:36px 16px;text-align:center;color:var(--muted);background:var(--surface);border:1px dashed var(--line);border-radius:14px}

/* ---- detail sheet ---- */
dialog{border:none;padding:0;background:var(--surface);color:var(--ink);width:min(780px,calc(100% - 32px));
  max-width:none;max-height:calc(100% - 48px);border-radius:18px;box-shadow:var(--shadow);overflow:auto;overscroll-behavior:contain}
dialog::backdrop{background:rgb(8 14 12 / .5)}
body:has(dialog[open]){overflow:hidden}
.sheet-head{position:sticky;top:0;z-index:2;background:var(--surface);border-bottom:1px solid var(--line);
  display:flex;justify-content:space-between;align-items:flex-start;gap:12px;padding:16px 18px 12px}
.sheet-head .tk{font-family:var(--font-display);font-stretch:80%;font-weight:800;font-size:1.9rem;line-height:1;margin:6px 0 2px}
.sheet-head .nm{font-weight:600}
.sheet-head .mt{color:var(--muted);font-size:.84rem}
.sheet-score{text-align:right;display:grid;justify-items:end;gap:4px}
.sheet-score b{font-family:var(--font-data);font-size:2.3rem;line-height:1;font-weight:600}
.close{flex:none}
.sheet-body{padding:16px 18px calc(22px + env(safe-area-inset-bottom,0px));display:grid;gap:22px}
.sheet-body h3{margin:0 0 10px;font-size:.74rem;letter-spacing:.08em;text-transform:uppercase;color:var(--muted)}
.facts{display:flex;flex-wrap:wrap;gap:6px 18px;font-size:.86rem}
.facts em{font-style:normal;color:var(--muted);margin-right:6px}
.breakdown{display:grid;gap:7px}
.bd{display:grid;grid-template-columns:118px minmax(0,1fr) 44px;gap:10px;align-items:center;font-size:.86rem}
.bd.main{font-weight:700}
.bd.sub span:first-child{padding-left:14px;color:var(--muted)}
.bd .meter{height:8px}
.bd .v{text-align:right}
.figs{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:16px}
.figs section{background:var(--surface-2);border-radius:12px;padding:12px 14px}
.figs h4{margin:0 0 6px;font-size:.8rem}
.figs dl{margin:0;display:grid;grid-template-columns:minmax(0,1fr) auto;gap:3px 12px;font-size:.84rem}
.figs dt{color:var(--muted)} .figs dd{margin:0;text-align:right}
.sig{display:grid;grid-template-columns:repeat(auto-fit,minmax(230px,1fr));gap:16px}
.sig ul{list-style:none;margin:0;padding:0;display:grid;gap:5px;font-size:.86rem}
.sig li{display:grid;grid-template-columns:16px minmax(0,1fr);gap:6px}
.sig li::before{font-weight:700}
.sig .bull li::before{content:"+";color:var(--pos)}
.sig .bear li::before{content:"\2212";color:var(--neg)}
.sig .watch li::before{content:"!";color:var(--warn)}
.sig h4{margin:0 0 8px;font-size:.84rem}
.links{display:flex;gap:10px;flex-wrap:wrap}
.links a{color:var(--accent);font-weight:600}
.fine{color:var(--muted);font-size:.78rem}

footer{margin:34px 0 28px;color:var(--muted);font-size:.82rem;display:grid;gap:10px}
footer p{margin:0;max-width:75ch}
.loader{display:flex;gap:10px;align-items:center;flex-wrap:wrap}


/* ---- top picks ---- */
.picks{margin-top:16px;display:grid;gap:12px}
.picks-head{display:flex;justify-content:space-between;align-items:flex-end;gap:12px;flex-wrap:wrap}
.picks-head h2{font-family:var(--font-display);font-stretch:85%;font-weight:800;font-size:1.45rem;margin:0;text-wrap:balance}
.picks-head p{margin:2px 0 0;color:var(--muted);font-size:.86rem;max-width:60ch}
.tabs{display:inline-flex;background:var(--surface);border:1px solid var(--line);border-radius:999px;padding:3px;gap:2px}
.tab{border:none;background:none;border-radius:999px;padding:6px 14px;font-weight:700;font-size:.88rem;cursor:pointer;color:var(--muted)}
.tab[aria-selected="true"].buy{background:var(--sb);color:var(--surface)}
.tab[aria-selected="true"].sell{background:var(--ss);color:var(--surface)}
.pick-list{display:grid;grid-template-columns:repeat(auto-fill,minmax(min(100%,360px),1fr));gap:12px}
.pick{background:var(--surface);border:1px solid var(--line);border-radius:14px;padding:14px 16px;display:flex;flex-direction:column;gap:8px;min-width:0;box-shadow:var(--shadow)}
.pick.buy{border-top:3px solid var(--sb)} .pick.sell{border-top:3px solid var(--ss)}
.pick-top{display:grid;grid-template-columns:auto minmax(0,1fr) auto;gap:12px;align-items:start}
.pick-n{font-family:var(--font-data);font-weight:600;color:var(--muted);font-size:.85rem;padding-top:4px}
.pick .tk{font-family:var(--font-display);font-stretch:85%;font-weight:800;font-size:1.3rem;line-height:1.1}
.pick .nm{font-size:.88rem;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.pick .mt{font-size:.78rem;color:var(--muted)}
.pick .sc{text-align:right;font-family:var(--font-data);font-weight:600;font-size:1.5rem;line-height:1}
.pick .sc small{display:block;font-family:var(--font-body);font-size:.66rem;letter-spacing:.06em;text-transform:uppercase;color:var(--muted);font-weight:600;margin-top:4px}
.pick .hl{font-weight:700;font-size:.95rem;line-height:1.35}
.pick .why{margin:0;font-size:.88rem;line-height:1.55;color:var(--ink)}
.pick .st{margin-top:auto;display:flex;flex-wrap:wrap;gap:4px 14px;font-size:.78rem;padding-top:8px;border-top:1px solid var(--line)}
.pick .st em{font-style:normal;color:var(--muted);margin-right:4px}
.pick .more{align-self:flex-start;background:none;border:none;color:var(--accent);font-weight:700;cursor:pointer;padding:2px 0}
.why-box{background:var(--surface-2);border-radius:12px;padding:12px 14px}
.why-box .hl{font-weight:700;margin:0 0 4px}
.why-box p{margin:0;line-height:1.6;font-size:.9rem}
.card .hl{grid-column:1 / -1;display:block;font-size:.84rem;font-weight:600;margin-top:4px}

@media (max-width:860px){
  .filters{grid-template-columns:minmax(0,1fr) minmax(0,1fr)}
  .filters .search{grid-column:1 / -1}
}
@media (max-width:640px){
  dialog{width:100%;max-height:92%;margin:auto 0 0;border-radius:18px 18px 0 0}
  .bd{grid-template-columns:96px minmax(0,1fr) 40px}
  .ladder-keys{gap:4px}
  .lk{padding:7px 6px 6px}
  .lk .lk-t{font-size:.7rem}
}
@media (prefers-reduced-motion: reduce){*{transition:none!important;animation:none!important}}
</style>
</head>
<body>
<header class="top">
  <div class="wrap">
    <div class="top-inner">
      <div>
        <h1>Daily Stock <span class="mark">Ratings</span></h1>
        <p class="runline" id="runline">Loading results...</p>
      </div>
      <button class="icon-btn" id="themeBtn" type="button" aria-label="Switch light or dark mode">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><circle cx="12" cy="12" r="8"/><path d="M12 4a8 8 0 0 1 0 16z" fill="currentColor"/></svg>
      </button>
    </div>
    <div class="banner" id="exampleBanner" hidden>Example data. These are made-up stocks that show how the screener works. Run cmc_screener.py to generate your real results, or load a results file at the bottom of the page.</div>
  </div>
</header>

<main class="wrap">
  <section class="picks" id="picks" aria-label="Top picks" hidden>
    <div class="picks-head">
      <div>
        <h2 id="picksTitle">Top picks</h2>
        <p id="picksNote">The strongest buys and sells, where the chart and the company's numbers agree.</p>
      </div>
      <div class="tabs" role="tablist" aria-label="Top buys or sells">
        <button class="tab buy" type="button" role="tab" data-side="buy" aria-selected="true">Top buys</button>
        <button class="tab sell" type="button" role="tab" data-side="sell" aria-selected="false">Top sells</button>
      </div>
    </div>
    <div class="pick-list" id="pickList"></div>
  </section>

  <section class="deck" aria-label="Screener filters">
    <div class="deck-head">
      <span class="label">Rating</span>
      <div class="presets" role="group" aria-label="Quick picks">
        <button class="preset" type="button" data-preset="buys" aria-pressed="false">Recommended buys</button>
        <button class="preset" type="button" data-preset="sells" aria-pressed="false">Sells to avoid</button>
        <button class="preset" type="button" data-preset="all" aria-pressed="false">All ratings</button>
      </div>
    </div>
    <div class="ladder-bar" id="ladderBar" aria-hidden="true"></div>
    <div class="ladder-keys" id="ladderKeys" role="group" aria-label="Show ratings"></div>

    <div class="filters">
      <div class="field search">
        <label for="q">Search</label>
        <input class="ctl" id="q" type="search" placeholder="Ticker or company" autocomplete="off">
      </div>
      <div class="field">
        <label for="market">Market</label>
        <select class="ctl" id="market"><option value="">All markets</option></select>
      </div>
      <div class="field">
        <label for="sector">Sector</label>
        <select class="ctl" id="sector"><option value="">All sectors</option></select>
      </div>
      <div class="field">
        <label for="sort">Sort by</label>
        <select class="ctl" id="sort">
          <option value="score">Overall score</option>
          <option value="fund">Fundamental score</option>
          <option value="tech">Technical score</option>
          <option value="val">Valuation score</option>
          <option value="pe">Lowest P/E</option>
          <option value="dy">Dividend yield</option>
          <option value="tgt">Analyst target upside</option>
          <option value="r3m">3-month return</option>
          <option value="mcap">Market cap</option>
        </select>
      </div>
      <button class="more-btn" id="moreBtn" type="button" aria-expanded="false" aria-controls="more">More filters</button>
    </div>
    <div class="more" id="more" hidden>
      <div class="field"><label for="maxPe">Max P/E</label><input class="ctl" id="maxPe" type="number" min="0" step="1" inputmode="decimal" placeholder="Any"></div>
      <div class="field"><label for="minDy">Min dividend yield %</label><input class="ctl" id="minDy" type="number" min="0" step="0.5" inputmode="decimal" placeholder="Any"></div>
      <div class="field"><label for="minFund">Min fundamental score</label><input class="ctl" id="minFund" type="number" min="-100" max="100" step="5" inputmode="numeric" placeholder="Any"></div>
      <div class="field"><label for="minMcap">Min market cap (bn)</label><input class="ctl" id="minMcap" type="number" min="0" step="0.5" inputmode="decimal" placeholder="Any"></div>
      <label class="check"><input id="zeroFee" type="checkbox"> $0 brokerage markets only</label>
      <label class="check"><input id="fullOnly" type="checkbox"> Full financials only</label>
      <button class="text-btn" id="resetBtn" type="button">Reset filters</button>
    </div>
  </section>

  <div class="resultbar">
    <h2 id="listTitle">Recommended buys</h2>
    <span class="count" id="count"></span>
  </div>
  <div id="results"></div>
  <div class="more-row" id="moreRow" hidden><button class="btn ghost" id="showMore" type="button">Show more</button></div>

  <footer>
    <p><strong>How the rating works.</strong> Each stock gets up to 20 technical signals (trend, momentum, volume) and up to 22 fundamental signals (valuation against sector peers, quality, growth, financial health, analyst sentiment). Every signal votes +1, 0 or &minus;1. The score runs from &minus;100 to +100: Strong Buy at 50 and above, Buy 15 to 50, Neutral &minus;15 to 15, Sell &minus;50 to &minus;15, Strong Sell below &minus;50. <span id="weightNote"></span></p>
    <p><span id="genNote"></span> Covers the main exchanges you can trade on CMC Invest; this is a personal tool, not affiliated with CMC Markets. Data from Yahoo Finance. Not financial advice: the rating is a mechanical summary of indicators, so use it to build a shortlist and do your own research before you trade.</p>
    <div class="loader">
      <label class="btn ghost" for="fileIn" tabindex="0" id="fileLabel">Load a results file</label>
      <input id="fileIn" type="file" accept=".csv,.json" hidden>
      <span class="fine" id="fileNote">Opens a cmc_ratings CSV from the screener script on this device.</span>
    </div>
  </footer>
</main>

<dialog id="detail" aria-labelledby="dTk"></dialog>

<script id="screener-data" type="application/json">__SCREENER_DATA__</script>
<script>
(() => {
"use strict";

const RATINGS = ["Strong Buy", "Buy", "Neutral", "Sell", "Strong Sell"];
const RC = {"Strong Buy": "sb", "Buy": "b", "Neutral": "n", "Sell": "s", "Strong Sell": "ss"};
const PAGE = 60;
const TEXT_KEYS = new Set(["rating","ticker","name","market","brokerage","exchange","sector","industry","ccy","bull","bear","fin","date","pick","head","why"]);
// CSV header (from cmc_screener.py) -> short key
const HEADER_MAP = {
  "Rank":"rank","Rating":"rating","Score":"score","Ticker":"ticker","Name":"name","Market":"market",
  "CMC Brokerage":"brokerage","Exchange":"exchange","Sector":"sector","Industry":"industry","Ccy":"ccy",
  "Price":"price","Mkt Cap (bn)":"mcap","Technical":"tech","Trend":"trend","Momentum":"mom","Volume":"vol",
  "Fundamentals":"fund","Valuation":"val","Quality":"qual","Growth":"grow","Health":"health","Sentiment":"sent",
  "P/E":"pe","Peer P/E (median)":"peer_pe","Fwd P/E":"fpe","PEG":"peg","P/B":"pb","P/S":"ps","EV/EBITDA":"ev",
  "Div Yield %":"dy","Payout %":"payout","ROE %":"roe","ROA %":"roa","Op Margin %":"opm","Net Margin %":"npm",
  "Revenue Growth %":"revg","Earnings Growth %":"earng","Debt/Equity (x)":"de","Current Ratio":"cr",
  "FCF Yield %":"fcfy","Analyst (1=SB, 5=SS)":"an","# Analysts":"an_n","Target Upside %":"tgt",
  "Short % Float":"short","RSI 14":"rsi","Stoch %K":"stoch","CCI 20":"cci","MFI 14":"mfi","ADX 14":"adx",
  "ATR %":"atr","Volume 20d vs 50d":"volr","3M %":"r3m","12M %":"r12m","vs Index 6M (pts)":"rel6m",
  "From 52W High %":"hi52","Bullish signals":"bull","Bearish signals":"bear","Financials":"fin","Price date":"date","Top Pick":"pick","Headline":"head","Why":"why"
};
const SORTS = {
  score:{dir:-1}, fund:{dir:-1}, tech:{dir:-1}, val:{dir:-1}, pe:{dir:1, pos:true}, dy:{dir:-1},
  tgt:{dir:-1}, r3m:{dir:-1}, mcap:{dir:-1}
};

// ---------- helpers ----------
const $ = s => document.querySelector(s);
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const has = v => v !== null && v !== undefined && v !== "" && !Number.isNaN(v);
const fmt = (v, d = 1) => has(v) ? Number(v).toLocaleString("en-AU", {minimumFractionDigits: d, maximumFractionDigits: d}) : "–";
const sgn = (v, d = 0, suf = "") => has(v) ? (v > 0 ? "+" : v < 0 ? "−" : "") + fmt(Math.abs(v), d) + suf : "–";
const cls = v => has(v) ? (v > 0 ? "pos" : v < 0 ? "neg" : "") : "";
const niceDate = d => { const t = new Date(String(d).slice(0, 10) + "T00:00:00"); return isNaN(t) ? String(d) : t.toLocaleDateString("en-AU", {day: "numeric", month: "short", year: "numeric"}); };
const price = v => !has(v) ? "–" : v >= 1000 ? fmt(v, 0) : v >= 1 ? fmt(v, 2) : fmt(v, 3);
const store = {
  get(k) { try { return JSON.parse(localStorage.getItem(k)); } catch (e) { return null; } },
  set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { /* storage unavailable */ } }
};
function meter(v) {
  if (!has(v)) return '<span class="meter" aria-hidden="true"></span>';
  const w = Math.min(Math.abs(v), 100) / 2;
  const pos = v >= 0 ? `left:50%;width:${w}%` : `right:50%;width:${w}%`;
  return `<span class="meter" aria-hidden="true"><i class="${v >= 0 ? "up" : "dn"}" style="${pos}"></i></span>`;
}

// ---------- data ----------
let ALL = [], META = {}, IS_EXAMPLE = false;

function unpack(d) {
  const cols = d.cols, out = [];
  for (const r of d.rows) { const o = {}; for (let i = 0; i < cols.length; i++) o[cols[i]] = r[i]; out.push(o); }
  return out;
}

function loadEmbedded() {
  const txt = ($("#screener-data").textContent || "").trim();
  if (!txt || txt.startsWith("__")) return null;
  try { return JSON.parse(txt); } catch (e) { return null; }
}

function parseCSV(text) {
  if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
  const rows = []; let row = [], f = "", q = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (q) {
      if (c === '"') { if (text[i + 1] === '"') { f += '"'; i++; } else q = false; }
      else f += c;
    } else if (c === '"') q = true;
    else if (c === ",") { row.push(f); f = ""; }
    else if (c === "\n") { row.push(f); rows.push(row); row = []; f = ""; }
    else if (c !== "\r") f += c;
  }
  if (f !== "" || row.length) { row.push(f); rows.push(row); }
  return rows;
}

function fromCSV(text) {
  const rows = parseCSV(text);
  const head = rows.shift().map(h => HEADER_MAP[h.trim()] || null);
  if (!head.includes("ticker") || !head.includes("rating")) throw new Error("This file doesn't look like a cmc_ratings CSV from the screener.");
  return rows.filter(r => r.length > 1).map(r => {
    const o = {};
    head.forEach((k, i) => {
      if (!k) return;
      const v = r[i];
      o[k] = TEXT_KEYS.has(k) ? (v || null) : (v === "" || v === undefined ? null : Number(v));
    });
    return o;
  });
}

function exampleData() {
  let seed = 20261028;
  const rnd = () => (seed = (seed * 16807) % 2147483647) / 2147483647;
  const pm = () => rnd() * 2 - 1;
  const clamp = v => Math.max(-100, Math.min(100, Math.round(v)));
  const markets = [
    ["Australia", "ASX - $0 first buy <=$1k/day", "AUD", ".AX"], ["United States", "Tier 1 - $0 brokerage", "USD", ""],
    ["United Kingdom", "Tier 1 - $0 brokerage", "GBP", ".L"], ["Japan", "Tier 1 - $0 brokerage", "JPY", ".T"],
    ["Germany", "Tier 2 - A$59 / 0.59%", "EUR", ".DE"]];
  const sectors = ["Basic Materials", "Financial Services", "Technology", "Healthcare", "Energy", "Industrials", "Consumer Defensive", "Real Estate"];
  const words = ["Resources", "Holdings", "Minerals", "Systems", "Capital", "Health", "Energy", "Logistics", "Foods", "Property"];
  const rows = [];
  for (let i = 0; i < 160; i++) {
    const m = markets[i % 7 < 4 ? 0 : i % 7 < 6 ? 1 : 2 + (i % 3)];
    const sector = sectors[Math.floor(rnd() * sectors.length)];
    const tb = pm(), fb = pm() * 0.8 + tb * 0.2;
    const trend = clamp((tb * 0.85 + pm() * 0.35) * 100), mom = clamp((tb * 0.75 + pm() * 0.45) * 100), vol = clamp((tb * 0.5 + pm() * 0.6) * 100);
    const val = clamp((fb * 0.4 + pm() * 0.7) * 100), qual = clamp((fb * 0.8 + pm() * 0.3) * 100), grow = clamp((fb * 0.7 + pm() * 0.4) * 100);
    const health = clamp((fb * 0.6 + pm() * 0.5) * 100), sent = clamp((fb * 0.6 + tb * 0.2 + pm() * 0.4) * 100);
    const tech = Math.round(0.4 * trend + 0.4 * mom + 0.2 * vol);
    const fund = Math.round(0.25 * val + 0.25 * qual + 0.2 * grow + 0.15 * health + 0.15 * sent);
    const score = Math.round(0.6 * tech + 0.4 * fund);
    const rating = score >= 50 ? "Strong Buy" : score >= 15 ? "Buy" : score > -15 ? "Neutral" : score > -50 ? "Sell" : "Strong Sell";
    const peer = 12 + rnd() * 14, pe = qual > -40 ? Math.max(4, peer * (1 - val / 160) + pm() * 3) : null;
    const r3m = tb * 18 + pm() * 8, r12m = tb * 40 + pm() * 20, rsi = 50 + tb * 16 + pm() * 8;
    const dy = rnd() < 0.6 ? Math.max(0, 1 + rnd() * 5 + fb) : 0, tgt = sent * 0.25 + pm() * 8;
    const roe = qual * 0.25 + 8, revg = grow * 0.2 + 3;
    const bull = [], bear = [];
    (trend > 0 ? bull : bear).push(trend > 0 ? "above 200d MA" : "below 200d MA");
    (trend > 20 ? bull : trend < -20 ? bear : []).push(trend > 20 ? "golden cross (50d > 200d)" : "death cross (50d < 200d)");
    (rsi > 50 ? bull : bear).push(rsi > 50 ? `RSI bullish (${rsi.toFixed(0)})` : `RSI weak (${rsi.toFixed(0)})`);
    (r3m > 2 ? bull : r3m < -2 ? bear : []).push(`3M ${r3m > 0 ? "+" : ""}${r3m.toFixed(0)}%`);
    if (pe) (pe < peer * 0.8 ? bull : pe > peer * 1.5 ? bear : []).push(`P/E ${pe.toFixed(1)} vs sector ${peer.toFixed(1)}`);
    (roe > 15 ? bull : roe < 0 ? bear : []).push(roe > 15 ? `ROE ${roe.toFixed(0)}%` : "negative ROE");
    (revg > 10 ? bull : revg < 0 ? bear : []).push(`revenue ${revg > 0 ? "+" : ""}${revg.toFixed(0)}%`);
    (tgt > 15 ? bull : tgt < 0 ? bear : []).push(`analyst target ${tgt > 0 ? "+" : ""}${tgt.toFixed(0)}%`);
    if (rsi > 70) bear.push(`RSI overbought (${rsi.toFixed(0)}) - caution`);
    rows.push({
      rating, score, ticker: `EX${String(i + 1).padStart(3, "0")}${m[3]}`, name: `Example ${words[i % words.length]} ${i + 1}`,
      market: m[0], brokerage: m[1], exchange: "Example", sector, industry: null, ccy: m[2],
      price: 0.5 + rnd() * 80, mcap: 0.2 + rnd() * rnd() * 120, tech, trend, mom, vol, fund, val, qual, grow, health, sent,
      pe, peer_pe: peer, fpe: pe ? pe * (1 - grow / 400) : null, peg: pe ? Math.max(0.3, 2 - grow / 60) : null,
      pb: 0.8 + rnd() * 4, ps: 0.5 + rnd() * 6, ev: pe ? pe * 0.6 : null, dy, payout: dy ? 30 + rnd() * 60 : null,
      roe, roa: roe / 2.5, opm: qual * 0.15 + 12, npm: qual * 0.1 + 8, revg, earng: grow * 0.3 + 4,
      de: Math.max(0, 1 - health / 120 + pm() * 0.3), cr: Math.max(0.4, 1.5 + health / 120), fcfy: health * 0.05 + 3,
      an: Math.min(5, Math.max(1, 2.6 - sent / 70)), an_n: Math.round(3 + rnd() * 18), tgt, short: rnd() * 6,
      rsi, stoch: 50 + tb * 25 + pm() * 15, cci: tb * 120 + pm() * 50, mfi: 50 + tb * 15 + pm() * 10, adx: 15 + rnd() * 20,
      atr: 1 + rnd() * 3, volr: 0.7 + rnd() * 0.8, r3m, r12m, rel6m: r12m / 2 - 3, hi52: Math.min(0, -2 - (1 - tb) * 18 + pm() * 4),
      bull: bull.join(", "), bear: bear.join(", "), fin: "Full", date: "2026-10-28"
    });
  }
  rows.sort((a, b) => b.score - a.score).forEach((r, i) => r.rank = i + 1);
  return rows;
}

// ---------- state ----------
const DEFAULTS = {ratings: ["Strong Buy", "Buy"], q: "", market: "", sector: "", sort: "score", dir: -1,
  maxPe: null, minDy: null, minFund: null, minMcap: null, zeroFee: false, fullOnly: false};
let S = Object.assign({}, DEFAULTS, store.get("cmcScreenerFilters") || {});
S.ratings = new Set(S.ratings);
let shown = PAGE;
const save = () => store.set("cmcScreenerFilters", Object.assign({}, S, {ratings: [...S.ratings], q: ""}));

function passesBase(s) {
  if (S.market && s.market !== S.market) return false;
  if (S.sector && s.sector !== S.sector) return false;
  if (S.q) {
    const q = S.q.toLowerCase();
    if (!String(s.ticker).toLowerCase().includes(q) && !String(s.name || "").toLowerCase().includes(q)) return false;
  }
  if (has(S.maxPe) && !(has(s.pe) && s.pe > 0 && s.pe <= S.maxPe)) return false;
  if (has(S.minDy) && !(has(s.dy) && s.dy >= S.minDy)) return false;
  if (has(S.minFund) && !(has(s.fund) && s.fund >= S.minFund)) return false;
  if (has(S.minMcap) && !(has(s.mcap) && s.mcap >= S.minMcap)) return false;
  if (S.zeroFee && !/^(ASX|Tier 1)/.test(s.brokerage || "")) return false;
  if (S.fullOnly && s.fin !== "Full") return false;
  return true;
}

function sorted(list) {
  const key = S.sort, dir = S.dir, posOnly = SORTS[key] && SORTS[key].pos;
  const val = s => { const v = s[key]; return has(v) && (!posOnly || v > 0) ? v : null; };
  return list.slice().sort((a, b) => {
    const x = val(a), y = val(b);
    if (x === null && y === null) return a.rank - b.rank;
    if (x === null) return 1;
    if (y === null) return -1;
    return x === y ? a.rank - b.rank : (x - y) * dir;
  });
}

// ---------- rendering ----------
const desktop = window.matchMedia("(min-width: 861px)");

function renderLadder(base) {
  const counts = Object.fromEntries(RATINGS.map(r => [r, 0]));
  for (const s of base) if (counts[s.rating] !== undefined) counts[s.rating]++;
  $("#ladderBar").innerHTML = RATINGS.map(r =>
    `<span class="${S.ratings.has(r) ? "" : "off"}" style="flex-grow:${counts[r]};background:var(--${RC[r]})"></span>`).join("");
  $("#ladderKeys").innerHTML = RATINGS.map(r =>
    `<button type="button" class="lk ${RC[r]}" data-r="${r}" aria-pressed="${S.ratings.has(r)}">
       <span class="lk-n">${counts[r].toLocaleString("en-AU")}</span><span class="lk-t">${r}</span></button>`).join("");
  const sel = [...S.ratings].sort((a, b) => RATINGS.indexOf(a) - RATINGS.indexOf(b)).join("|");
  const presets = {buys: "Strong Buy|Buy", sells: "Sell|Strong Sell", all: RATINGS.join("|")};
  document.querySelectorAll(".preset").forEach(b => b.setAttribute("aria-pressed", presets[b.dataset.preset] === sel));
  const title = sel === presets.buys ? "Recommended buys" : sel === presets.sells ? "Sells to avoid"
    : sel === presets.all ? "All rated stocks" : sel ? sel.replace(/\|/g, ", ") : "No ratings selected";
  $("#listTitle").textContent = title;
}

function statLine(s) {
  const pe = has(s.pe) ? `${fmt(s.pe, 1)}${has(s.peer_pe) ? ` <span class="muted">vs ${fmt(s.peer_pe, 1)}</span>` : ""}` : "–";
  return `<span><em>Price</em>${has(s.price) ? `${esc(s.ccy || "")} ${price(s.price)}` : "–"}</span>
    <span><em>Tech</em><b class="${cls(s.tech)}">${sgn(s.tech)}</b></span>
    <span><em>Fund</em><b class="${cls(s.fund)}">${sgn(s.fund)}</b></span>
    <span><em>P/E</em>${pe}</span>
    <span><em>Div</em>${has(s.dy) ? fmt(s.dy, 1) + "%" : "–"}</span>
    <span><em>3M</em><b class="${cls(s.r3m)}">${sgn(s.r3m, 0, "%")}</b></span>`;
}

function renderCards(items) {
  return `<div class="cards">${items.map(s => `
    <button type="button" class="card" data-t="${esc(s.ticker)}">
      <span class="c-top"><span class="chip ${RC[s.rating] || "n"}">${esc(s.rating)}</span><span class="rank">#${s.rank}</span></span>
      <span class="tk">${esc(s.ticker)}</span>
      <span class="nm">${esc(s.name || "")}</span>
      ${s.pick ? `<span class="hl">${s.pick === "Top Buy" ? "\u2605 Top buy" : "\u2605 Top sell"}: ${esc(s.head || "")}</span>` : ""}
      <span class="mt">${esc([s.market, s.sector].filter(Boolean).join(" · "))}</span>
      <span class="sc"><small>Score</small><b class="score-${RC[s.rating] || "n"}">${sgn(s.score)}</b>${meter(s.score)}</span>
      <span class="st num">${statLine(s)}</span>
    </button>`).join("")}</div>`;
}

const COLS = [
  ["rank", "#"], ["rating", "Rating", "l"], ["score", "Score"], ["ticker", "Ticker", "l"], ["name", "Company", "l"],
  ["market", "Market", "l"], ["sector", "Sector", "l"], ["price", "Price"], ["tech", "Tech"], ["fund", "Fund"],
  ["pe", "P/E"], ["peer_pe", "Peer P/E"], ["dy", "Div %"], ["r3m", "3M %"], ["tgt", "Target"]
];
const SORTABLE = new Set(["score", "tech", "fund", "pe", "dy", "r3m", "tgt"]);

function renderTable(items) {
  const head = COLS.map(([k, label, align]) => {
    const active = S.sort === k;
    const aria = active ? ` aria-sort="${S.dir < 0 ? "descending" : "ascending"}"` : "";
    const inner = SORTABLE.has(k) ? `<button type="button" data-sort="${k}">${label}</button>` : label;
    return `<th class="${align || ""}" scope="col"${aria}>${inner}</th>`;
  }).join("");
  const body = items.map(s => `<tr data-t="${esc(s.ticker)}">
    <td class="num muted">${s.rank}</td>
    <td class="l"><span class="chip ${RC[s.rating] || "n"}">${esc(s.rating)}</span></td>
    <td class="scr"><b class="score-${RC[s.rating] || "n"}">${sgn(s.score)}</b>${meter(s.score)}</td>
    <td class="l tkc"><button type="button" data-t="${esc(s.ticker)}">${esc(s.ticker)}</button></td>
    <td class="l co" title="${esc(s.name || "")}">${esc(s.name || "")}</td>
    <td class="l">${esc(s.market || "")}</td>
    <td class="l muted sec" title="${esc(s.sector || "")}">${esc(s.sector || "–")}</td>
    <td class="num">${has(s.price) ? `<span class="muted">${esc(s.ccy || "")}</span> ${price(s.price)}` : "–"}</td>
    <td class="num ${cls(s.tech)}">${sgn(s.tech)}</td>
    <td class="num ${cls(s.fund)}">${sgn(s.fund)}</td>
    <td class="num">${fmt(s.pe, 1)}</td>
    <td class="num muted">${fmt(s.peer_pe, 1)}</td>
    <td class="num">${has(s.dy) ? fmt(s.dy, 1) : "–"}</td>
    <td class="num ${cls(s.r3m)}">${sgn(s.r3m, 0)}</td>
    <td class="num ${cls(s.tgt)}">${sgn(s.tgt, 0, "%")}</td></tr>`).join("");
  return `<div class="table-wrap"><table><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table></div>`;
}

let current = [];
function render() {
  const base = ALL.filter(passesBase);
  renderLadder(base);
  current = sorted(base.filter(s => S.ratings.has(s.rating)));
  const items = current.slice(0, shown);
  $("#count").textContent = current.length
    ? `Showing ${items.length.toLocaleString("en-AU")} of ${current.length.toLocaleString("en-AU")} stocks`
    : "";
  $("#results").innerHTML = current.length
    ? (desktop.matches ? renderTable(items) : renderCards(items))
    : `<div class="empty">No stocks match these filters. Try another rating, a wider P/E limit or <button class="text-btn" type="button" data-reset>reset the filters</button>.</div>`;
  $("#moreRow").hidden = current.length <= shown;
  $("#showMore").textContent = `Show ${Math.min(PAGE, current.length - shown).toLocaleString("en-AU")} more`;
  $("#sort").value = S.sort;
}

// ---------- top picks ----------
let pickSide = "buy";
function renderPicks() {
  const want = pickSide === "buy" ? "Top Buy" : "Top Sell";
  const all = ALL.filter(s => s.pick);
  $("#picks").hidden = !all.length;
  if (!all.length) return;
  const list = ALL.filter(s => s.pick === want).sort((a, b) => pickSide === "buy" ? b.score - a.score : a.score - b.score);
  const nb = all.filter(s => s.pick === "Top Buy").length, ns = all.length - nb;
  $("#picksTitle").textContent = `Today's top ${nb} buys and ${ns} sells`;
  document.querySelectorAll(".tab").forEach(t => t.setAttribute("aria-selected", t.dataset.side === pickSide));
  $("#pickList").innerHTML = list.length ? list.map((s, i) => `
    <article class="pick ${pickSide}">
      <div class="pick-top">
        <span class="pick-n">${i + 1}</span>
        <div style="min-width:0"><div class="tk">${esc(s.ticker)}</div><div class="nm">${esc(s.name || "")}</div>
          <div class="mt">${esc([s.market, s.sector].filter(Boolean).join(" \u00b7 "))}</div></div>
        <div class="sc score-${RC[s.rating] || "n"}">${sgn(s.score)}<small>${esc(s.rating)}</small></div>
      </div>
      <div class="hl">${esc(s.head || "")}</div>
      <p class="why">${esc(s.why || "")}</p>
      <div class="st num">
        <span><em>Price</em>${esc(s.ccy || "")} ${price(s.price)}</span>
        <span><em>P/E</em>${has(s.pe) ? fmt(s.pe, 1) + (has(s.peer_pe) ? ` <span class="muted">vs ${fmt(s.peer_pe, 1)}</span>` : "") : "loss-making"}</span>
        <span><em>3M</em><b class="${cls(s.r3m)}">${sgn(s.r3m, 0, "%")}</b></span>
        <span><em>12M</em><b class="${cls(s.r12m)}">${sgn(s.r12m, 0, "%")}</b></span>
      </div>
      <button class="more" type="button" data-t="${esc(s.ticker)}">All figures and signals</button>
    </article>`).join("") : `<div class="empty">No stocks met the bar for a top ${pickSide} today.</div>`;
}

// ---------- detail sheet ----------
function dl(rows) {
  return `<dl>${rows.map(([k, v]) => `<dt>${k}</dt><dd class="num">${v}</dd>`).join("")}</dl>`;
}
function signalList(str) {
  return String(str || "").split(", ").map(x => x.trim()).filter(Boolean);
}

function openDetail(ticker) {
  const s = ALL.find(x => x.ticker === ticker);
  if (!s) return;
  const rc = RC[s.rating] || "n";
  const bear = signalList(s.bear), bull = signalList(s.bull);
  const watch = bear.filter(x => / - caution$/.test(x)).map(x => x.replace(/ - caution$/, ""));
  const bad = bear.filter(x => !/ - caution$/.test(x));
  const bd = (label, v, kind) => `<div class="bd ${kind}"><span>${label}</span>${meter(v)}<span class="v num ${cls(v)}">${sgn(v)}</span></div>`;
  const pct = (v, d = 1) => has(v) ? fmt(v, d) + "%" : "–";
  const an = has(s.an) ? `${fmt(s.an, 1)}${has(s.an_n) ? ` <span class="muted">(${fmt(s.an_n, 0)})</span>` : ""}` : "–";
  const yahoo = `https://finance.yahoo.com/quote/${encodeURIComponent(s.ticker)}`;
  $("#detail").innerHTML = `
    <div class="sheet-head">
      <div>
        <span class="chip ${rc}">${esc(s.rating)}</span>
        <div class="tk" id="dTk">${esc(s.ticker)}</div>
        <div class="nm">${esc(s.name || "")}</div>
        <div class="mt">${esc([s.market, s.exchange, s.sector, s.industry].filter(Boolean).join(" · "))}</div>
      </div>
      <div class="sheet-score">
        <button class="icon-btn close" type="button" data-close aria-label="Close">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" aria-hidden="true"><path d="M6 6l12 12M18 6L6 18"/></svg></button>
        <b class="score-${rc}">${sgn(s.score)}</b><span class="fine">score, rank #${s.rank}</span>
      </div>
    </div>
    <div class="sheet-body">
      <div class="facts">
        <span><em>Price</em><span class="num">${esc(s.ccy || "")} ${price(s.price)}</span></span>
        <span><em>Mkt cap</em><span class="num">${has(s.mcap) ? fmt(s.mcap, 2) + " bn " + esc(s.ccy || "") : "–"}</span></span>
        <span><em>CMC brokerage</em>${esc((s.brokerage || "–").replace("<=", "≤"))}</span>
      </div>
      ${s.why ? `<section><h3>Why it's rated ${esc(s.rating)}${s.pick ? ` \u00b7 ${esc(s.pick)}` : ""}</h3>
        <div class="why-box"><p class="hl">${esc(s.head || "")}</p><p>${esc(s.why)}</p></div></section>` : ""}
      <section>
        <h3>Score breakdown</h3>
        <div class="breakdown">
          ${bd("Technical", s.tech, "main")}${bd("Trend", s.trend, "sub")}${bd("Momentum", s.mom, "sub")}${bd("Volume", s.vol, "sub")}
          ${bd("Fundamental", s.fund, "main")}${bd("Valuation", s.val, "sub")}${bd("Quality", s.qual, "sub")}
          ${bd("Growth", s.grow, "sub")}${bd("Health", s.health, "sub")}${bd("Sentiment", s.sent, "sub")}
        </div>
      </section>
      <section class="sig">
        <div class="bull"><h4>Bullish (${bull.length})</h4><ul>${bull.map(x => `<li><span>${esc(x)}</span></li>`).join("") || "<li><span class=\"muted\">None</span></li>"}</ul></div>
        <div class="bear"><h4>Bearish (${bad.length})</h4><ul>${bad.map(x => `<li><span>${esc(x)}</span></li>`).join("") || "<li><span class=\"muted\">None</span></li>"}</ul>
          ${watch.length ? `<h4 style="margin-top:14px">Watch</h4><ul class="watch">${watch.map(x => `<li><span>${esc(x)}</span></li>`).join("")}</ul>` : ""}
        </div>
      </section>
      <section>
        <h3>Key figures</h3>
        <div class="figs">
          <section><h4>Valuation</h4>${dl([["P/E", fmt(s.pe, 1)], ["Sector peer P/E", fmt(s.peer_pe, 1)], ["Forward P/E", fmt(s.fpe, 1)],
            ["PEG", fmt(s.peg, 2)], ["Price / book", fmt(s.pb, 2)], ["Price / sales", fmt(s.ps, 2)], ["EV / EBITDA", fmt(s.ev, 1)],
            ["Dividend yield", pct(s.dy)], ["Payout ratio", pct(s.payout, 0)]])}</section>
          <section><h4>Quality &amp; growth</h4>${dl([["Return on equity", pct(s.roe)], ["Return on assets", pct(s.roa)],
            ["Operating margin", pct(s.opm)], ["Net margin", pct(s.npm)], ["Revenue growth", `<span class="${cls(s.revg)}">${sgn(s.revg, 1, "%")}</span>`],
            ["Earnings growth", `<span class="${cls(s.earng)}">${sgn(s.earng, 1, "%")}</span>`]])}</section>
          <section><h4>Health &amp; analysts</h4>${dl([["Debt / equity", has(s.de) ? fmt(s.de, 2) + "x" : "–"], ["Current ratio", fmt(s.cr, 2)],
            ["FCF yield", pct(s.fcfy)], ["Analyst rating (1 best)", an], ["Target upside", `<span class="${cls(s.tgt)}">${sgn(s.tgt, 1, "%")}</span>`],
            ["Short interest", pct(s.short)]])}</section>
          <section><h4>Price action</h4>${dl([["RSI 14", fmt(s.rsi, 0)], ["Stochastic %K", fmt(s.stoch, 0)], ["CCI 20", fmt(s.cci, 0)],
            ["MFI 14", fmt(s.mfi, 0)], ["ADX 14", fmt(s.adx, 0)], ["ATR", pct(s.atr)], ["Volume vs avg", has(s.volr) ? fmt(s.volr, 2) + "x" : "–"],
            ["3-month return", `<span class="${cls(s.r3m)}">${sgn(s.r3m, 1, "%")}</span>`], ["12-month return", `<span class="${cls(s.r12m)}">${sgn(s.r12m, 1, "%")}</span>`],
            ["vs index (6M)", `<span class="${cls(s.rel6m)}">${sgn(s.rel6m, 1, " pts")}</span>`], ["From 52-week high", `<span class="${cls(s.hi52)}">${sgn(s.hi52, 1, "%")}</span>`]])}</section>
        </div>
      </section>
      <div class="links"><a href="${yahoo}" target="_blank" rel="noopener">View ${esc(s.ticker)} on Yahoo Finance</a></div>
      <p class="fine">Prices to ${esc(s.date ? niceDate(s.date) : "–")}. Financials: ${esc(s.fin || "–")}. ${IS_EXAMPLE ? "Example data, not a real stock. " : ""}Not financial advice.</p>
    </div>`;
  const dlg = $("#detail");
  if (!dlg.open) { try { dlg.showModal(); } catch (e) { dlg.setAttribute("open", ""); } }
  dlg.scrollTop = 0;
}

// ---------- setup ----------
function fillSelect(sel, values, first) {
  const cur = sel.value;
  sel.innerHTML = `<option value="">${first}</option>` + values.map(v => `<option>${esc(v)}</option>`).join("");
  sel.value = values.includes(cur) ? cur : "";
}

function setData(rows, meta, example) {
  ALL = rows.filter(r => r.ticker && r.rating);
  META = meta || {};
  IS_EXAMPLE = !!example;
  $("#exampleBanner").hidden = !IS_EXAMPLE;
  const markets = [...new Set(ALL.map(s => s.market).filter(Boolean))].sort();
  const sectors = [...new Set(ALL.map(s => s.sector).filter(Boolean))].sort();
  fillSelect($("#market"), markets, "All markets");
  fillSelect($("#sector"), sectors, "All sectors");
  if (!markets.includes(S.market)) S.market = "";
  if (!sectors.includes(S.sector)) S.sector = "";
  $("#market").value = S.market; $("#sector").value = S.sector;
  const dates = ALL.map(s => s.date).filter(Boolean).sort();
  const last = META.price_date || dates[dates.length - 1];
  const parts = [`${ALL.length.toLocaleString("en-AU")} stocks rated`];
  if (last) parts.push(`prices to ${niceDate(last)}`);
  if (markets.length) parts.push(markets.length > 4 ? `${markets.length} markets` : markets.join(", "));
  $("#runline").textContent = (IS_EXAMPLE ? "Example data · " : "") + parts.join(" · ");
  const fw = META.fund_weight;
  $("#genNote").textContent = META.generated ? `Results generated ${niceDate(META.generated.slice(0, 10))} at ${META.generated.slice(11)}.` : "";
  $("#weightNote").textContent = has(fw) ? `This run weights fundamentals at ${Math.round(fw * 100)}% and technicals at ${Math.round((1 - fw) * 100)}%.` : "";
  shown = PAGE;
  renderPicks();
  render();
}

function bind() {
  document.querySelectorAll(".tab").forEach(t => t.addEventListener("click", () => { pickSide = t.dataset.side; renderPicks(); }));
  $("#pickList").addEventListener("click", e => { const b = e.target.closest("[data-t]"); if (b) openDetail(b.dataset.t); });
  $("#ladderKeys").addEventListener("click", e => {
    const b = e.target.closest("[data-r]"); if (!b) return;
    const r = b.dataset.r;
    S.ratings.has(r) ? S.ratings.delete(r) : S.ratings.add(r);
    shown = PAGE; save(); render();
  });
  document.querySelectorAll(".preset").forEach(b => b.addEventListener("click", () => {
    const p = b.dataset.preset;
    S.ratings = new Set(p === "buys" ? ["Strong Buy", "Buy"] : p === "sells" ? ["Sell", "Strong Sell"] : RATINGS);
    if (p === "sells" && S.sort === "score") S.dir = 1; else if (S.sort === "score") S.dir = -1;
    shown = PAGE; save(); render();
  }));
  let t;
  $("#q").addEventListener("input", e => { clearTimeout(t); t = setTimeout(() => { S.q = e.target.value.trim(); shown = PAGE; render(); }, 150); });
  $("#market").addEventListener("change", e => { S.market = e.target.value; shown = PAGE; save(); render(); });
  $("#sector").addEventListener("change", e => { S.sector = e.target.value; shown = PAGE; save(); render(); });
  $("#sort").addEventListener("change", e => { S.sort = e.target.value; S.dir = (SORTS[S.sort] || {dir: -1}).dir; shown = PAGE; save(); render(); });
  const numIn = (id, key) => $(id).addEventListener("input", e => {
    const v = e.target.value === "" ? null : Number(e.target.value);
    S[key] = Number.isFinite(v) ? v : null; shown = PAGE; save(); render();
  });
  numIn("#maxPe", "maxPe"); numIn("#minDy", "minDy"); numIn("#minFund", "minFund"); numIn("#minMcap", "minMcap");
  $("#zeroFee").addEventListener("change", e => { S.zeroFee = e.target.checked; shown = PAGE; save(); render(); });
  $("#fullOnly").addEventListener("change", e => { S.fullOnly = e.target.checked; shown = PAGE; save(); render(); });
  $("#moreBtn").addEventListener("click", () => {
    const open = $("#more").hidden; $("#more").hidden = !open; $("#moreBtn").setAttribute("aria-expanded", open);
  });
  const reset = () => {
    S = Object.assign({}, DEFAULTS, {ratings: new Set(DEFAULTS.ratings)});
    ["#q", "#maxPe", "#minDy", "#minFund", "#minMcap"].forEach(id => $(id).value = "");
    $("#market").value = ""; $("#sector").value = ""; $("#zeroFee").checked = false; $("#fullOnly").checked = false;
    shown = PAGE; save(); render();
  };
  $("#resetBtn").addEventListener("click", reset);
  $("#showMore").addEventListener("click", () => { shown += PAGE; render(); });
  $("#results").addEventListener("click", e => {
    if (e.target.closest("[data-reset]")) return reset();
    const sb = e.target.closest("[data-sort]");
    if (sb) {
      const k = sb.dataset.sort;
      if (S.sort === k) S.dir *= -1; else { S.sort = k; S.dir = (SORTS[k] || {dir: -1}).dir; }
      shown = PAGE; save(); render(); return;
    }
    const row = e.target.closest("[data-t]");
    if (row) openDetail(row.dataset.t);
  });
  $("#detail").addEventListener("click", e => {
    if (e.target.closest("[data-close]") || e.target === $("#detail")) $("#detail").close();
  });
  desktop.addEventListener("change", render);

  $("#fileLabel").addEventListener("keydown", e => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); $("#fileIn").click(); } });
  $("#fileIn").addEventListener("change", async e => {
    const f = e.target.files && e.target.files[0]; if (!f) return;
    $("#fileNote").textContent = `Reading ${f.name}...`;
    try {
      const text = await f.text();
      let rows, meta = {};
      if (/\.json$/i.test(f.name)) { const d = JSON.parse(text); rows = d.cols ? unpack(d) : d; meta = d.meta || {}; }
      else rows = fromCSV(text);
      setData(rows, meta, false);
      $("#fileNote").textContent = `Showing ${f.name}. Reload the page to go back to the built-in results.`;
    } catch (err) {
      $("#fileNote").textContent = `Couldn't read ${f.name}: ${err.message || err}`;
    }
    e.target.value = "";
  });

  $("#themeBtn").addEventListener("click", () => {
    const root = document.documentElement;
    const dark = root.dataset.theme ? root.dataset.theme === "dark" : window.matchMedia("(prefers-color-scheme: dark)").matches;
    root.dataset.theme = dark ? "light" : "dark";
    store.set("cmcScreenerTheme", root.dataset.theme);
  });
}

function restoreInputs() {
  const set = (id, v) => { if (has(v)) $(id).value = v; };
  set("#maxPe", S.maxPe); set("#minDy", S.minDy); set("#minFund", S.minFund); set("#minMcap", S.minMcap);
  $("#zeroFee").checked = !!S.zeroFee; $("#fullOnly").checked = !!S.fullOnly;
  if (S.maxPe || S.minDy || S.minFund || S.minMcap || S.zeroFee || S.fullOnly) { $("#more").hidden = false; $("#moreBtn").setAttribute("aria-expanded", "true"); }
}

const theme = store.get("cmcScreenerTheme");
if (theme === "dark" || theme === "light") document.documentElement.dataset.theme = theme;
bind();
restoreInputs();
const embedded = loadEmbedded();
if (embedded) setData(unpack(embedded), embedded.meta, false);
else setData(exampleData(), {}, true);
})();
</script>
</body>
</html>
"""


# =============================================================================
# Main
# =============================================================================

def parse_args():
    codes = list(MARKETS)
    p = argparse.ArgumentParser(
        description="Rate every stock on CMC Invest from Strong Buy to Strong Sell.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="Market codes: " + ", ".join(f"{c}={MARKETS[c].name}" for c in codes))
    p.add_argument("--markets", nargs="+", default=["all"], metavar="CODE",
                   help="Markets to scan: 'all' (default) or codes, e.g. --markets au us")
    p.add_argument("--tickers", metavar="FILE",
                   help="Scan only the tickers in this .txt/.csv/.xlsx file (Yahoo format, e.g. BHP.AX, AAPL)")
    p.add_argument("--suffix", metavar=".AX",
                   help="With --tickers: add this suffix to every ticker without a dot (e.g. .AX for a "
                        "list of ASX codes). For mixed lists, write full tickers instead: BHP.AX, AAPL")
    p.add_argument("--fundamentals", choices=["full", "quick", "off"], default="full",
                   help="full (default): pull every stock's financials - P/E vs sector, PEG, margins, "
                        "growth, debt, cash flow, analyst targets (one request per stock, cached 7 days). "
                        "quick: only P/E, P/B, EPS and analyst rating from the stock list (fast). "
                        "off: technical only.")
    p.add_argument("--top", type=int, default=10, metavar="N",
                   help="How many Top Buys and Top Sells to flag (default 10 each; 0 = none)")
    p.add_argument("--fund-weight", type=float, default=DEFAULT_FUND_WEIGHT,
                   help=f"Share of the final score from fundamentals, 0-1 (default {DEFAULT_FUND_WEIGHT})")
    p.add_argument("--min-mcap", type=float, default=50e6,
                   help="Minimum market cap passed to Yahoo's screener (default 50,000,000)")
    p.add_argument("--min-volume", type=float, default=20000,
                   help="Minimum 3-month average daily volume in shares (default 20,000)")
    p.add_argument("--no-filters", action="store_true", help="Include every stock, however small or illiquid")
    p.add_argument("--max-per-market", type=int, help="Only take the N largest stocks per market (quick tests)")
    p.add_argument("--out", default="screener_output", help="Output folder (default ./screener_output)")
    p.add_argument("--refresh", action="store_true",
                   help="Ignore cached prices and financials and download everything again")
    p.add_argument("--no-open", action="store_true", help="Don't open the screener page in your browser when done")
    p.add_argument("--list-markets", action="store_true", help="Show the CMC Invest markets and exit")
    a = p.parse_args()

    if a.list_markets:
        print(f"{'Code':<6}{'Market':<16}{'CMC brokerage':<31}Index")
        for m in MARKETS.values():
            print(f"{m.code:<6}{m.name:<16}{m.brokerage:<31}{m.benchmark}")
        sys.exit(0)

    a.markets = [x.lower() for x in a.markets]
    if "all" in a.markets:
        a.markets = codes
    bad = [x for x in a.markets if x not in MARKETS]
    if bad:
        p.error(f"Unknown market code(s): {', '.join(bad)}. Valid: all, {', '.join(codes)}")
    if not 0 <= a.fund_weight <= 1:
        p.error("--fund-weight must be between 0 and 1")
    if a.no_filters:
        a.min_mcap = a.min_volume = 0
    return a


def main():
    try:                                   # don't crash on odd company names in a Windows console
        sys.stdout.reconfigure(errors="replace")
    except Exception:  # noqa: BLE001
        pass
    args = parse_args()
    t0 = time.time()
    today = pd.Timestamp(datetime.now().date())
    out_dir = Path(args.out)
    cache_root = out_dir / "cache"
    cache_dir = cache_root / today.strftime("%Y-%m-%d")
    cache_dir.mkdir(parents=True, exist_ok=True)
    for old in cache_root.iterdir():                 # keep only today's price cache
        if old.is_dir() and old != cache_dir:
            shutil.rmtree(old, ignore_errors=True)

    # ---- 1. Universe --------------------------------------------------------
    if args.tickers:
        tickers = read_tickers_file(args.tickers, args.suffix)
        print(f"\n[1/5] {len(tickers)} tickers from {args.tickers}")
        universe = pd.DataFrame({"ticker": tickers, "market": [market_for_ticker(t) for t in tickers],
                                 "full": False})
    else:
        print(f"\n[1/5] Building stock list for {len(args.markets)} CMC Invest market(s)")
        universe = universe_from_screener(args.markets, args.min_mcap, args.min_volume, args.max_per_market)
    if universe.empty:
        sys.exit("\nNo stocks found. Check your internet connection, or try again later "
                 "(Yahoo sometimes rate-limits).")
    print(f"      {len(universe):,} stocks to analyse")

    # ---- 2. Prices ----------------------------------------------------------
    print("\n[2/5] Downloading 2 years of daily prices (cached for today)")
    bench_syms = sorted({MARKETS[m].benchmark for m in universe["market"].unique() if m in MARKETS})
    bench_frames = download_prices(bench_syms, cache_dir, args.refresh, "Indexes")
    bench_close = {sym: f["Close"] for sym, f in bench_frames.items()}
    prices = {}
    for code, grp in universe.groupby("market", sort=False):
        label = MARKETS[code].name if code in MARKETS else "Other"
        prices.update(download_prices(grp["ticker"].tolist(), cache_dir, args.refresh, label))
    print(f"      prices for {len(prices):,} of {len(universe):,}")

    # ---- 3. Technicals ------------------------------------------------------
    print("\n[3/5] Technical analysis")
    recs, tech, skipped = [], {}, {}
    for i, rec in enumerate(universe.to_dict("records"), 1):
        if i % 2000 == 0:
            print(f"      {i:,}/{len(universe):,}", flush=True)
        df = prices.get(rec["ticker"])
        if df is None:
            skipped["no price data"] = skipped.get("no price data", 0) + 1
            continue
        bench = bench_close.get(MARKETS[rec["market"]].benchmark) if rec["market"] in MARKETS else None
        metrics, votes = technicals(df, bench, today)
        if metrics is None:
            skipped[votes] = skipped.get(votes, 0) + 1
            continue
        tech[rec["ticker"]] = (metrics, votes)
        recs.append(rec)
    if not recs:
        sys.exit("No stocks could be scored.")
    print(f"      {len(recs):,} stocks with usable price history")

    # ---- 4. Fundamentals ----------------------------------------------------
    peers, peer_pool = None, list(recs)
    if args.fundamentals == "full":
        print(f"\n[4/5] Company financials ({INFO_WORKERS} at a time; saved as they download)")
        cache_file = cache_root / "fundamentals.jsonl"
        if args.refresh and cache_file.exists():
            cache_file.unlink()
        cache = FundCache(cache_file, FUND_CACHE_DAYS)
        infos = fetch_fundamentals([r["ticker"] for r in recs], cache)
        cache.close()
        for r in recs:
            if r["ticker"] in infos:
                merge_info(r, infos[r["ticker"]])
        print(f"      full financials for {sum(1 for r in recs if r.get('full')):,} of {len(recs):,}")
        if args.tickers:
            # A short watchlist has too few peers of its own - also compare against
            # sector peers saved from earlier full-market runs (if any).
            mine = {r["ticker"] for r in recs}
            peer_pool += [dict(d, market=market_for_ticker(t)) for t, d in cache.data.items() if t not in mine]
            if len(peer_pool) > len(recs):
                print(f"      comparing valuations with {len(peer_pool) - len(recs):,} saved peers "
                      "from earlier runs")
    else:
        print(f"\n[4/5] Company financials: {args.fundamentals}")
    if args.fundamentals != "off":
        peers = build_peers(peer_pool)

    # ---- 5. Score & write ---------------------------------------------------
    print("\n[5/5] Scoring and writing report")
    fund_weight = args.fund_weight if args.fundamentals != "off" else 0.0
    results = [score_record(r, *tech[r["ticker"]], peers, fund_weight) for r in recs]
    report = build_report_frame(results)
    report = add_explanations(report, args.top if args.fundamentals != "off" else 0)
    if args.fundamentals == "off" and "Financials" in report:
        report["Financials"] = "Off"

    stamp = today.strftime("%Y-%m-%d")
    xlsx = out_dir / f"cmc_ratings_{stamp}.xlsx"
    csv = out_dir / f"cmc_ratings_{stamp}.csv"
    run_note = (f"Run {datetime.now():%Y-%m-%d %H:%M} | {len(report):,} stocks rated | "
                f"markets: {', '.join(sorted(universe['market'].unique()))} | "
                f"financials: {args.fundamentals} | fundamentals weight {fund_weight:.0%}")
    try:
        report.to_csv(csv, index=False, encoding="utf-8-sig")
        write_excel(report, xlsx, run_note)
    except PermissionError:
        suffix = datetime.now().strftime("%H%M%S")
        xlsx = out_dir / f"cmc_ratings_{stamp}_{suffix}.xlsx"
        csv = out_dir / f"cmc_ratings_{stamp}_{suffix}.csv"
        print("      (today's file is open in Excel - writing a new copy)")
        report.to_csv(csv, index=False, encoding="utf-8-sig")
        write_excel(report, xlsx, run_note)

    dates = [d for d in report.get("Price date", pd.Series(dtype=object)).dropna()]
    site_meta = {
        "generated": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "price_date": str(max(dates)) if dates else None,
        "fund_weight": fund_weight,
        "financials": args.fundamentals,
        "markets": sorted({MARKETS[m].name for m in universe["market"].unique() if m in MARKETS}),
    }
    site_path = write_site(report, out_dir, site_meta)

    # ---- Console summary ----------------------------------------------------
    counts = report["Rating"].value_counts()
    print("\n" + "=" * 72)
    print(f" {len(report):,} stocks rated in {(time.time() - t0) / 60:.1f} min")
    print("  " + "   ".join(f"{r}: {counts.get(r, 0):,}" for r in RATING_ORDER))
    if skipped:
        print("  skipped: " + ", ".join(f"{k} {v:,}" for k, v in skipped.items()))
    for label in ("Top Buy", "Top Sell"):
        picks = report[report["Top Pick"] == label] if "Top Pick" in report else report.iloc[0:0]
        if len(picks):
            print(f"\n {label}s")
            for _, r in picks.iterrows():
                name = str(r["Name"])[:28] if pd.notna(r["Name"]) else ""
                print(f"  {int(r['Score']):>4}  {r['Ticker']:<11}{name:<30}{r['Headline']}")
    size_mb = site_path.stat().st_size / 1e6
    print(f"\n Screener page: {site_path.resolve()}  ({size_mb:.1f} MB)")
    print("   To deploy, upload the 'site' folder (just index.html) to any static host -")
    print("   e.g. drag it onto app.netlify.com/drop, or push it to GitHub Pages.")
    print(f" Excel: {xlsx.resolve()}")
    print(f" CSV:   {csv.resolve()}")
    print(" Not financial advice - a mechanical summary of indicators.")
    print("=" * 72)
    if not args.no_open:
        try:
            import webbrowser
            webbrowser.open(site_path.resolve().as_uri())
        except Exception:  # noqa: BLE001 - no browser available is fine
            pass


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit("\nStopped. Anything downloaded so far is cached - run the same command again to carry on.")
