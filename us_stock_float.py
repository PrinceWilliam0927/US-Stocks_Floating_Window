# -*- coding: utf-8 -*-
"""
美股悬浮行情窗。

运行:
  python us_stock_float.py

说明:
  - 使用 Python 标准库实现，不需要安装第三方包。
  - 报价来自 Yahoo Finance 的公开 quote 接口，实际延迟取决于数据源。
  - 自选股、窗口位置和刷新间隔保存在 us_stock_float_config.json。
"""

from __future__ import annotations

import ctypes
import json
import logging
import queue
import re
import threading
import urllib.parse
import urllib.request
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Iterable

import tkinter as tk


APP_DIR = Path(__file__).resolve().parent
CONFIG_PATH = APP_DIR / "us_stock_float_config.json"
LOG_PATH = APP_DIR / "us_stock_float_error.log"
QUOTE_ENDPOINT = "https://query1.finance.yahoo.com/v7/finance/quote"
CHART_ENDPOINT = "https://query1.finance.yahoo.com/v8/finance/chart"
YAHOO_QUOTE_PAGE = "https://finance.yahoo.com/quote/{symbol}/"

WINDOW_W = 460
WINDOW_H = 455
MIN_WINDOW_W = 360
MIN_WINDOW_H = 320
MIN_REFRESH_SECONDS = 1
MAX_REFRESH_SECONDS = 300
TOPMOST_REFRESH_MS = 1200
HWND_TOPMOST = -1
SWP_NOSIZE = 0x0001
SWP_NOMOVE = 0x0002
SWP_NOACTIVATE = 0x0010
SWP_SHOWWINDOW = 0x0040
GEOMETRY_RE = re.compile(r"^(\d+)x(\d+)([+-]\d+)([+-]\d+)$")

BG = "#111820"
PANEL = "#19222d"
PANEL_2 = "#202b37"
TEXT = "#eef3f7"
MUTED = "#95a3b3"
BORDER = "#2d3a48"
ACCENT = "#4b8df7"
UP = "#20b26b"
DOWN = "#e45555"
INPUT_BG = "#0d1319"

STATE_LABELS = {
    "REGULAR": "盘中",
    "PRE": "盘前",
    "PREPRE": "盘前",
    "POST": "盘后",
    "POSTPOST": "盘后",
    "OVERNIGHT": "隔夜",
    "CLOSED": "休市",
}


logging.basicConfig(
    filename=str(LOG_PATH),
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
)


@dataclass
class Quote:
    symbol: str
    name: str = ""
    price: float | None = None
    change: float | None = None
    change_pct: float | None = None
    currency: str = ""
    market_state: str = ""
    market_time: int | None = None
    display_time: str = ""
    source: str = ""
    error: str = ""
    extended_state: str = ""
    extended_price: float | None = None
    extended_change: float | None = None
    extended_change_pct: float | None = None
    extended_time: int | None = None


def normalize_symbol(raw: str) -> str:
    """把用户输入整理成 Yahoo Finance 常用的股票代码格式。"""
    symbol = raw.strip().upper().replace(" ", "")
    if symbol.endswith(".US"):
        symbol = symbol[:-3]
    symbol = symbol.replace(".", "-")
    return "".join(ch for ch in symbol if ch.isalnum() or ch in "-^=")


def split_symbols(text: str) -> list[str]:
    text = text.replace("，", ",").replace("；", ",").replace("、", ",")
    parts = re.split(r"[,\s;]+", text)
    symbols: list[str] = []
    for part in parts:
        symbol = normalize_symbol(part)
        if symbol and symbol not in symbols:
            symbols.append(symbol)
    return symbols


def sanitize_geometry(geometry: str, screen_w: int, screen_h: int) -> str:
    match = GEOMETRY_RE.match(geometry.strip())
    if not match:
        return ""

    width = max(int(match.group(1)), MIN_WINDOW_W)
    height = max(int(match.group(2)), MIN_WINDOW_H)
    x = int(match.group(3))
    y = int(match.group(4))
    return f"{width}x{height}{x:+d}{y:+d}"


def force_tk_window_topmost(window: tk.Tk | tk.Toplevel, activate: bool = False) -> None:
    try:
        window.attributes("-topmost", True)
        window.lift()
        if activate:
            window.focus_force()
    except tk.TclError:
        return

    try:
        hwnd = int(window.winfo_id())
        user32 = ctypes.windll.user32  # type: ignore[attr-defined]
        flags = SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW
        if not activate:
            flags |= SWP_NOACTIVATE
        user32.SetWindowPos(
            hwnd,
            HWND_TOPMOST,
            0,
            0,
            0,
            0,
            flags,
        )
        if activate:
            user32.BringWindowToTop(hwnd)
            user32.SetForegroundWindow(hwnd)
    except (AttributeError, OSError, TypeError, ValueError, tk.TclError):
        return


def load_config() -> dict:
    default = {
        "symbols": [],
        "refresh_seconds": 15,
        "topmost": True,
        "alpha": 0.86,
        "alert_threshold_pct": 1.0,
        "alert_cooldown_seconds": 300,
        "geometry": "",
        "selected_symbol": "",
    }
    if not CONFIG_PATH.exists():
        return default

    try:
        data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        logging.exception("Failed to read config")
        return default

    if not isinstance(data, dict):
        return default

    merged = default | data
    merged["symbols"] = [
        symbol
        for symbol in (normalize_symbol(str(item)) for item in merged.get("symbols", []))
        if symbol
    ]
    try:
        merged["refresh_seconds"] = int(merged.get("refresh_seconds", 15))
    except (TypeError, ValueError):
        merged["refresh_seconds"] = 15
    merged["refresh_seconds"] = min(
        max(merged["refresh_seconds"], MIN_REFRESH_SECONDS),
        MAX_REFRESH_SECONDS,
    )
    merged["topmost"] = bool(merged.get("topmost", True))
    try:
        merged["alpha"] = float(merged.get("alpha", 0.86))
    except (TypeError, ValueError):
        merged["alpha"] = 0.86
    merged["alpha"] = min(max(merged["alpha"], 0.35), 1.0)
    try:
        merged["alert_threshold_pct"] = float(merged.get("alert_threshold_pct", 1.0))
    except (TypeError, ValueError):
        merged["alert_threshold_pct"] = 1.0
    merged["alert_threshold_pct"] = min(max(merged["alert_threshold_pct"], 0.1), 50.0)
    try:
        merged["alert_cooldown_seconds"] = int(merged.get("alert_cooldown_seconds", 300))
    except (TypeError, ValueError):
        merged["alert_cooldown_seconds"] = 300
    merged["alert_cooldown_seconds"] = min(max(merged["alert_cooldown_seconds"], 0), 86400)
    merged["selected_symbol"] = normalize_symbol(str(merged.get("selected_symbol", "")))
    return merged


def save_config(config: dict) -> None:
    try:
        CONFIG_PATH.write_text(
            json.dumps(config, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    except OSError:
        logging.exception("Failed to save config")


def _read_float(item: dict, key: str) -> float | None:
    value = item.get(key)
    if isinstance(value, dict):
        value = value.get("raw")
    if value is None:
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _read_int(item: dict, key: str) -> int | None:
    value = item.get(key)
    if isinstance(value, dict):
        value = value.get("raw")
    if value is None:
        return None
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def _quote_from_yahoo_item(item: dict) -> Quote:
    symbol = str(item.get("symbol") or "").upper()
    state = str(item.get("marketState") or "").upper()
    current_state, current_prefix = _pick_yahoo_current_prefix(item, state)
    price = _read_float(item, "regularMarketPrice")
    change = _read_float(item, "regularMarketChange")
    change_pct = _read_float(item, "regularMarketChangePercent")
    market_time = _read_int(item, "regularMarketTime")
    extended_state = ""
    extended_price = None
    extended_change = None
    extended_change_pct = None
    extended_time = None
    if current_prefix != "regularMarket":
        extended_state = current_state
        extended_price = _read_float(item, f"{current_prefix}Price")
        extended_change = _read_float(item, f"{current_prefix}Change")
        extended_change_pct = _read_float(item, f"{current_prefix}ChangePercent")
        extended_time = _read_int(item, f"{current_prefix}Time")

    name = str(item.get("shortName") or item.get("longName") or "").strip()
    return Quote(
        symbol=symbol,
        name=name,
        price=price,
        change=change,
        change_pct=change_pct,
        currency=str(item.get("currency") or ""),
        market_state=current_state,
        market_time=market_time,
        source="Yahoo",
        extended_state=extended_state,
        extended_price=extended_price,
        extended_change=extended_change,
        extended_change_pct=extended_change_pct,
        extended_time=extended_time,
    )


def _pick_yahoo_current_prefix(item: dict, state: str) -> tuple[str, str]:
    """Pick Yahoo's current active stream, matching the quote page state."""
    if state == "OVERNIGHT" and _read_float(item, "overnightMarketPrice") is not None:
        return "OVERNIGHT", "overnightMarket"
    if state.startswith("PRE") and _read_float(item, "preMarketPrice") is not None:
        return "PRE", "preMarket"
    if state.startswith("POST") and _read_float(item, "postMarketPrice") is not None:
        return "POST", "postMarket"

    candidates: list[tuple[int, int, str, str]] = []
    for priority, (prefix, display_state) in enumerate(
        (
            ("overnightMarket", "OVERNIGHT"),
            ("preMarket", "PRE"),
            ("postMarket", "POST"),
        )
    ):
        if _read_float(item, f"{prefix}Price") is None:
            continue
        market_time = _read_int(item, f"{prefix}Time") or 0
        candidates.append((market_time, priority, prefix, display_state))

    regular_time = _read_int(item, "regularMarketTime") or 0
    candidates = [item for item in candidates if item[0] > regular_time]

    if not candidates:
        return state or "REGULAR", "regularMarket"

    _, _, prefix, display_state = max(candidates)
    return display_state, prefix


def fetch_quotes(symbols: Iterable[str], timeout: int = 10) -> dict[str, Quote]:
    clean_symbols = [normalize_symbol(symbol) for symbol in symbols]
    clean_symbols = [symbol for symbol in clean_symbols if symbol]
    if not clean_symbols:
        return {}

    result: dict[str, Quote] = {}
    remaining = list(clean_symbols)
    page_error: Exception | None = None
    quote_error: Exception | None = None

    try:
        page_quotes = fetch_quotes_yahoo_page(clean_symbols, timeout=timeout)
    except Exception as yahoo_page_exc:
        page_error = yahoo_page_exc
        logging.warning("Yahoo quote page failed, trying Yahoo quote API: %s", yahoo_page_exc)
    else:
        result.update(page_quotes)
        remaining = [
            symbol
            for symbol in clean_symbols
            if symbol not in result or not quote_matches_current_session(result[symbol])
        ]

    if remaining:
        try:
            quote_quotes = fetch_quotes_yahoo(remaining, timeout=timeout)
        except Exception as yahoo_exc:
            quote_error = yahoo_exc
            logging.warning("Yahoo quote failed, trying Yahoo chart: %s", yahoo_exc)
        else:
            for symbol, quote in quote_quotes.items():
                if not quote.error or symbol not in result:
                    result[symbol] = quote
            remaining = [
                symbol
                for symbol in clean_symbols
                if symbol not in result or not quote_matches_current_session(result[symbol])
            ]

    if remaining:
        try:
            chart_quotes = fetch_quotes_yahoo_chart(remaining, timeout=timeout)
        except Exception as yahoo_chart_exc:
            if not result or all(quote.error for quote in result.values()):
                first_error = quote_error or page_error or yahoo_chart_exc
                raise RuntimeError(
                    f"Yahoo quote 获取失败: {first_error}; Yahoo chart 获取失败: {yahoo_chart_exc}"
                ) from yahoo_chart_exc
            logging.warning("Yahoo chart failed for fallback symbols: %s", yahoo_chart_exc)
        else:
            for symbol, quote in chart_quotes.items():
                if not quote.error or symbol not in result:
                    result[symbol] = quote

    for symbol in clean_symbols:
        quote = result.get(symbol)
        if quote is None or not quote_matches_current_session(quote):
            result[symbol] = Quote(symbol=symbol, source="Yahoo", error="未取得当前价")

    return {
        symbol: result.get(symbol, Quote(symbol=symbol, source="Yahoo", error="未找到"))
        for symbol in clean_symbols
    }


def _extract_yahoo_prefetched_quote(page_text: str, symbol: str) -> Quote:
    escaped_symbol = f'\\"symbol\\":\\"{symbol}\\"'
    candidates: list[tuple[int, Quote]] = []
    seen_body_starts: set[int] = set()

    for match in re.finditer(re.escape(escaped_symbol), page_text):
        body_key = page_text.rfind('"body":', 0, match.start())
        if body_key < 0:
            continue

        body_start = body_key + len('"body":')
        if body_start in seen_body_starts:
            continue
        seen_body_starts.add(body_start)

        try:
            body_text, _ = json.JSONDecoder().raw_decode(page_text[body_start:])
            payload = json.loads(body_text)
        except (json.JSONDecodeError, TypeError):
            continue

        rows = payload.get("quoteResponse", {}).get("result", [])
        for item in rows:
            if str(item.get("symbol") or "").upper() != symbol:
                continue

            score = 0
            state = str(item.get("marketState") or "").upper()
            if state == "OVERNIGHT":
                score += 100
            if item.get("overnightMarketPrice") is not None:
                score += 50
            if item.get("postMarketPrice") is not None or item.get("preMarketPrice") is not None:
                score += 10
            candidates.append((score, _quote_from_yahoo_item(item)))

    if not candidates:
        raise RuntimeError(f"Yahoo 页面未找到 {symbol} 报价")

    candidates.sort(key=lambda item: item[0], reverse=True)
    return candidates[0][1]


def fetch_quote_yahoo_page(symbol: str, timeout: int = 10) -> Quote:
    clean_symbol = normalize_symbol(symbol)
    if not clean_symbol:
        return Quote(symbol="", source="Yahoo", error="未选择股票")

    url = YAHOO_QUOTE_PAGE.format(symbol=urllib.parse.quote(clean_symbol, safe=""))
    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0 Safari/537.36"
            ),
            "Accept": "text/html,*/*",
            "Accept-Language": "en-US,en;q=0.9",
        },
    )
    with urllib.request.urlopen(request, timeout=timeout) as response:
        page_text = response.read().decode("utf-8", errors="replace")
    return _extract_yahoo_prefetched_quote(page_text, clean_symbol)


def fetch_quotes_yahoo_page(symbols: Iterable[str], timeout: int = 10) -> dict[str, Quote]:
    clean_symbols = [normalize_symbol(symbol) for symbol in symbols]
    clean_symbols = [symbol for symbol in clean_symbols if symbol]
    if not clean_symbols:
        return {}

    result: dict[str, Quote] = {}
    failures: list[str] = []

    def fetch_one(symbol: str) -> tuple[str, Quote | None, Exception | None]:
        try:
            return symbol, fetch_quote_yahoo_page(symbol, timeout=timeout), None
        except Exception as exc:  # noqa: BLE001 - keep per-symbol failures visible.
            return symbol, None, exc

    max_workers = min(max(len(clean_symbols), 1), 8)
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = [executor.submit(fetch_one, symbol) for symbol in clean_symbols]
        for future in as_completed(futures):
            symbol, quote, exc = future.result()
            if quote is not None:
                result[symbol] = quote
            else:
                failures.append(f"{symbol}: {exc}")
                result[symbol] = Quote(symbol=symbol, source="Yahoo", error="Yahoo 页面失败")

    if failures and all(quote.error for quote in result.values()):
        raise RuntimeError("; ".join(failures[:3]))
    return result


def fetch_quotes_yahoo(symbols: Iterable[str], timeout: int = 10) -> dict[str, Quote]:
    clean_symbols = [normalize_symbol(symbol) for symbol in symbols]
    clean_symbols = [symbol for symbol in clean_symbols if symbol]
    if not clean_symbols:
        return {}

    query = urllib.parse.urlencode({"symbols": ",".join(clean_symbols)})
    url = f"{QUOTE_ENDPOINT}?{query}"
    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0 Safari/537.36"
            ),
            "Accept": "application/json,text/plain,*/*",
        },
    )

    with urllib.request.urlopen(request, timeout=timeout) as response:
        payload = json.loads(response.read().decode("utf-8"))

    rows = payload.get("quoteResponse", {}).get("result", [])
    quotes = {
        quote.symbol: quote
        for quote in (_quote_from_yahoo_item(item) for item in rows)
        if quote.symbol
    }

    result: dict[str, Quote] = {}
    for symbol in clean_symbols:
        result[symbol] = quotes.get(symbol, Quote(symbol=symbol, error="未找到"))
    return result


def fetch_quote_yahoo_chart(symbol: str, timeout: int = 10) -> Quote:
    clean_symbol = normalize_symbol(symbol)
    if not clean_symbol:
        return Quote(symbol="", source="Yahoo", error="未选择股票")

    query = urllib.parse.urlencode(
        {
            "range": "1d",
            "interval": "1m",
            "includePrePost": "true",
        }
    )
    url = f"{CHART_ENDPOINT}/{urllib.parse.quote(clean_symbol, safe='')}?{query}"
    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0 Safari/537.36"
            ),
            "Accept": "application/json,text/plain,*/*",
        },
    )

    with urllib.request.urlopen(request, timeout=timeout) as response:
        payload = json.loads(response.read().decode("utf-8"))

    chart = payload.get("chart", {})
    error = chart.get("error")
    if error:
        description = error.get("description") if isinstance(error, dict) else str(error)
        raise RuntimeError(description or "Yahoo chart 返回错误")

    results = chart.get("result") or []
    if not results:
        raise RuntimeError("Yahoo chart 没有返回报价数据")

    row = results[0]
    meta = row.get("meta", {})
    timestamps = row.get("timestamp") or []
    quote_rows = row.get("indicators", {}).get("quote") or []
    closes = quote_rows[0].get("close") if quote_rows else []

    points: list[tuple[int, float]] = []
    for timestamp, price in zip(timestamps, closes or [], strict=False):
        if timestamp is None or price is None:
            continue
        try:
            points.append((int(timestamp), float(price)))
        except (TypeError, ValueError):
            continue

    regular_price = _read_float(meta, "regularMarketPrice")
    current_state = str(meta.get("marketState") or "CLOSED").upper()
    extended_state = ""
    extended_price = None
    extended_change = None
    extended_change_pct = None
    extended_time = None
    if points:
        last_ts, last_price = points[-1]
    previous_close = _read_float(meta, "previousClose") or _read_float(meta, "chartPreviousClose")
    regular_time = _read_int(meta, "regularMarketTime")
    change = None
    change_pct = None
    if regular_price is not None and previous_close not in (None, 0):
        change = regular_price - previous_close
        change_pct = (change / previous_close) * 100

    if points:
        period = meta.get("currentTradingPeriod") or {}
        pre = period.get("pre") or {}
        regular = period.get("regular") or {}
        post = period.get("post") or {}
        if pre.get("start") and pre.get("end") and pre["start"] <= last_ts < pre["end"]:
            current_state = "PRE"
            extended_state = "PRE"
        elif post.get("start") and post.get("end") and post["start"] <= last_ts <= post["end"]:
            current_state = "POST"
            extended_state = "POST"
        elif regular_time and last_ts > regular_time:
            current_state = "POST"
            extended_state = "POST"
        elif regular.get("start") and last_ts < regular["start"]:
            current_state = "PRE"
            extended_state = "PRE"

        if extended_state and regular_price is not None and abs(last_price - regular_price) > 0.000001:
            extended_price = last_price
            extended_time = last_ts
            extended_change = last_price - regular_price
            if regular_price:
                extended_change_pct = extended_change / regular_price * 100

    return Quote(
        symbol=clean_symbol,
        price=regular_price,
        change=change,
        change_pct=change_pct,
        currency=str(meta.get("currency") or ""),
        market_state=current_state,
        market_time=regular_time,
        source="Yahoo",
        extended_state=extended_state,
        extended_price=extended_price,
        extended_change=extended_change,
        extended_change_pct=extended_change_pct,
        extended_time=extended_time,
    )


def fetch_quotes_yahoo_chart(symbols: Iterable[str], timeout: int = 10) -> dict[str, Quote]:
    clean_symbols = [normalize_symbol(symbol) for symbol in symbols]
    clean_symbols = [symbol for symbol in clean_symbols if symbol]
    if not clean_symbols:
        return {}

    result: dict[str, Quote] = {}
    failures: list[str] = []
    for symbol in clean_symbols:
        try:
            quote = fetch_quote_yahoo_chart(symbol, timeout=timeout)
        except Exception as exc:  # noqa: BLE001 - keep per-symbol failures visible.
            failures.append(f"{symbol}: {exc}")
            result[symbol] = Quote(symbol=symbol, source="Yahoo", error="Yahoo chart 失败")
        else:
            result[symbol] = quote

    if failures and all(quote.error for quote in result.values()):
        raise RuntimeError("; ".join(failures[:3]))

    return result


def format_price(value: float | None) -> str:
    if value is None:
        return "--"
    if abs(value) >= 1000:
        return f"{value:,.2f}"
    if abs(value) >= 1:
        return f"{value:.2f}"
    return f"{value:.4f}"


def format_change(value: float | None, pct: float | None) -> str:
    if value is None and pct is None:
        return "--"
    value_text = "" if value is None else f"{value:+.2f}"
    pct_text = "" if pct is None else f"{pct:+.2f}%"
    return "  ".join(part for part in (value_text, pct_text) if part)


def _nth_weekday_day(year: int, month: int, weekday: int, n: int) -> int:
    first = datetime(year, month, 1)
    offset = (weekday - first.weekday()) % 7
    return 1 + offset + (n - 1) * 7


def _eastern_offset_for_utc(utc_dt: datetime) -> timedelta:
    year = utc_dt.year
    dst_start_day = _nth_weekday_day(year, 3, 6, 2)
    dst_end_day = _nth_weekday_day(year, 11, 6, 1)
    dst_start_utc = datetime(year, 3, dst_start_day, 7, tzinfo=timezone.utc)
    dst_end_utc = datetime(year, 11, dst_end_day, 6, tzinfo=timezone.utc)
    if dst_start_utc <= utc_dt < dst_end_utc:
        return timedelta(hours=-4)
    return timedelta(hours=-5)


def eastern_datetime_from_timestamp(timestamp: int | None) -> datetime | None:
    if not timestamp:
        return None
    try:
        utc_dt = datetime.fromtimestamp(int(timestamp), tz=timezone.utc)
    except (OSError, OverflowError, TypeError, ValueError):
        return None
    return (utc_dt + _eastern_offset_for_utc(utc_dt)).replace(tzinfo=None)


def current_eastern_datetime() -> datetime:
    utc_dt = datetime.now(timezone.utc)
    return (utc_dt + _eastern_offset_for_utc(utc_dt)).replace(tzinfo=None)


def _minutes_since_midnight(value: datetime) -> int:
    return value.hour * 60 + value.minute


def current_market_session(now_et: datetime | None = None) -> str:
    now_et = now_et or current_eastern_datetime()
    weekday = now_et.weekday()
    minutes = _minutes_since_midnight(now_et)

    if weekday < 5 and 9 * 60 + 30 <= minutes < 16 * 60:
        return "REGULAR"
    if weekday < 5 and 4 * 60 <= minutes < 9 * 60 + 30:
        return "PRE"
    if weekday < 5 and 16 * 60 <= minutes < 20 * 60:
        return "POST"
    if weekday == 6 and minutes >= 20 * 60:
        return "OVERNIGHT"
    if weekday in (0, 1, 2, 3) and minutes >= 20 * 60:
        return "OVERNIGHT"
    if weekday in (0, 1, 2, 3, 4) and minutes < 4 * 60:
        return "OVERNIGHT"
    return "CLOSED"


def quote_active_time(quote: Quote | None) -> int | None:
    if quote is None:
        return None
    if quote.extended_price is not None and quote.extended_time is not None:
        return quote.extended_time
    if quote.price is not None and quote.market_time is not None:
        return quote.market_time
    return None


def quote_matches_current_session(quote: Quote, now_et: datetime | None = None) -> bool:
    if quote.error:
        return False

    now_et = now_et or current_eastern_datetime()
    session = current_market_session(now_et)
    if session == "CLOSED":
        return False

    if session == "REGULAR":
        if quote.price is None or quote.market_time is None:
            return False
        quote_et = eastern_datetime_from_timestamp(quote.market_time)
        if quote_et is None or quote_et.date() != now_et.date():
            return False
        return _minutes_since_midnight(quote_et) >= 9 * 60 + 30

    if quote.extended_price is None or quote.extended_time is None:
        return False
    if quote.market_time is not None and quote.extended_time <= quote.market_time:
        return False

    quote_et = eastern_datetime_from_timestamp(quote.extended_time)
    if quote_et is None:
        return False

    now_ts = datetime.now(timezone.utc).timestamp()
    if quote.extended_time > now_ts + 300 or now_ts - quote.extended_time > 12 * 60 * 60:
        return False

    quote_session = current_market_session(quote_et)
    if session == "OVERNIGHT":
        return quote_session == "OVERNIGHT"
    return quote_session == session and quote_et.date() == now_et.date()


def format_market_time(timestamp: int | None) -> str:
    market_dt = eastern_datetime_from_timestamp(timestamp)
    if market_dt is None:
        return "--:--"
    return f"{market_dt:%H:%M:%S}"


def short_text(text: str, max_len: int) -> str:
    if len(text) <= max_len:
        return text
    return text[: max_len - 1] + "…"


class StockFloatApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.config = load_config()
        self.symbols: list[str] = list(dict.fromkeys(self.config["symbols"]))
        self.quotes: dict[str, Quote] = {}
        self.selected_symbol = self._initial_selected_symbol()
        self.result_queue: queue.Queue[tuple[dict[str, Quote] | None, str | None]] = queue.Queue()
        self.refresh_after_id: str | None = None
        self.topmost_after_id: str | None = None
        self.fetching = False
        self.previous_alert_prices: dict[str, float] = {}
        self.last_alert_at: dict[str, float] = {}

        self.topmost_var = tk.BooleanVar(value=self.config.get("topmost", True))
        self.interval_var = tk.StringVar(value=str(self.config.get("refresh_seconds", 15)))
        self.alpha_var = tk.StringVar(value=f"{float(self.config.get('alpha', 0.86)):.2f}")
        self.alert_threshold_var = tk.StringVar(
            value=f"{float(self.config.get('alert_threshold_pct', 1.0)):.1f}"
        )
        self.status_var = tk.StringVar(value="输入代码后添加")

        self._setup_window()
        self._build_ui()
        self._render_rows()
        self._schedule_refresh(delay_ms=250)
        self._schedule_topmost_guard()

    def _initial_selected_symbol(self) -> str:
        selected = normalize_symbol(str(self.config.get("selected_symbol", "")))
        if selected in self.symbols:
            return selected
        return self.symbols[0] if self.symbols else ""

    def _setup_window(self) -> None:
        self.root.title("美股悬浮行情")
        self.root.configure(bg=BG)
        self.root.overrideredirect(False)
        self.root.minsize(MIN_WINDOW_W, MIN_WINDOW_H)
        self.root.resizable(True, True)
        self.root.attributes("-topmost", bool(self.topmost_var.get()))
        self.root.attributes("-alpha", float(self.config.get("alpha", 0.96)))

        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        geometry = sanitize_geometry(str(self.config.get("geometry") or ""), screen_w, screen_h)
        if geometry:
            self.root.geometry(geometry)
        else:
            x = max(screen_w - WINDOW_W - 28, 20)
            self.root.geometry(f"{WINDOW_W}x{WINDOW_H}+{x}+80")

        self.root.protocol("WM_DELETE_WINDOW", self.close)
        self.root.bind("<FocusOut>", lambda _event: self._force_topmost())
        self.root.bind("<Map>", lambda _event: self.root.after(80, self._force_topmost))
        self.root.bind_all("<Control-Key-0>", self.restore_default_size)

    def _build_ui(self) -> None:
        self.header = tk.Frame(self.root, bg=PANEL_2, height=36)
        self.header.pack(fill="x")
        self.header.pack_propagate(False)

        title = tk.Label(
            self.header,
            text="行情",
            bg=PANEL_2,
            fg=TEXT,
            font=("Microsoft YaHei UI", 10, "bold"),
        )
        title.pack(side="left", padx=(10, 6))

        self.status_label = tk.Label(
            self.header,
            textvariable=self.status_var,
            bg=PANEL_2,
            fg=MUTED,
            font=("Microsoft YaHei UI", 8),
        )
        self.status_label.pack(side="left", fill="x", expand=True)

        self._make_header_button("还原", self.restore_default_size).pack(side="right", padx=(0, 8))
        self._make_header_button("刷新", self.refresh_now).pack(side="right", padx=(0, 4))

        body = tk.Frame(self.root, bg=BG)
        body.pack(fill="both", expand=True, padx=10, pady=(10, 8))

        entry_row = tk.Frame(body, bg=BG)
        entry_row.pack(fill="x")

        self.symbol_entry = tk.Entry(
            entry_row,
            bg=INPUT_BG,
            fg=TEXT,
            insertbackground=TEXT,
            relief="flat",
            font=("Consolas", 11),
        )
        self.symbol_entry.pack(side="left", fill="x", expand=True, ipady=6)
        self.symbol_entry.bind("<Return>", lambda _event: self.add_symbols())

        add_button = self._make_body_button(entry_row, "添加", self.add_symbols, width=7)
        add_button.pack(side="left", padx=(8, 0), ipady=2)

        hint = tk.Label(
            body,
            text="支持多个代码: AAPL, MSFT, NVDA",
            bg=BG,
            fg=MUTED,
            font=("Microsoft YaHei UI", 8),
            anchor="w",
        )
        hint.pack(fill="x", pady=(5, 8))

        table_header = tk.Frame(body, bg=BG)
        table_header.pack(fill="x", pady=(0, 4))
        self._table_label(table_header, "代码", 12, "w").grid(row=0, column=0, sticky="ew")
        self._table_label(table_header, "价格", 10, "e").grid(row=0, column=1, sticky="ew")
        self._table_label(table_header, "涨跌", 12, "e").grid(row=0, column=2, sticky="ew")
        self._table_label(table_header, "时间", 8, "e").grid(row=0, column=3, sticky="ew")
        self._table_label(table_header, "", 5, "e").grid(row=0, column=4, sticky="ew")
        for col, weight in enumerate((2, 2, 2, 1, 0)):
            table_header.columnconfigure(col, weight=weight)

        list_wrap = tk.Frame(body, bg=BORDER)
        list_wrap.pack(fill="both", expand=True)

        self.canvas = tk.Canvas(
            list_wrap,
            bg=BG,
            bd=0,
            highlightthickness=0,
            yscrollincrement=28,
        )
        self.scrollbar = tk.Scrollbar(list_wrap, orient="vertical", command=self.canvas.yview)
        self.rows_frame = tk.Frame(self.canvas, bg=BG)
        self.rows_window = self.canvas.create_window((0, 0), window=self.rows_frame, anchor="nw")

        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        self.rows_frame.bind("<Configure>", self._on_rows_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.rows_frame.bind("<MouseWheel>", self._on_mousewheel)

        footer = tk.Frame(body, bg=BG)
        footer.pack(fill="x", pady=(8, 0))

        topmost_check = tk.Checkbutton(
            footer,
            text="置顶",
            variable=self.topmost_var,
            command=self.toggle_topmost,
            bg=BG,
            fg=MUTED,
            selectcolor=INPUT_BG,
            activebackground=BG,
            activeforeground=TEXT,
            font=("Microsoft YaHei UI", 9),
            relief="flat",
        )
        topmost_check.pack(side="left")

        tk.Label(
            footer,
            text="刷新秒数",
            bg=BG,
            fg=MUTED,
            font=("Microsoft YaHei UI", 9),
        ).pack(side="left", padx=(14, 5))

        self.interval_spin = tk.Spinbox(
            footer,
            from_=MIN_REFRESH_SECONDS,
            to=MAX_REFRESH_SECONDS,
            increment=1,
            width=5,
            textvariable=self.interval_var,
            command=self.update_interval,
            bg=INPUT_BG,
            fg=TEXT,
            buttonbackground=PANEL_2,
            insertbackground=TEXT,
            relief="flat",
            font=("Consolas", 9),
        )
        self.interval_spin.pack(side="left")
        self.interval_spin.bind("<Return>", lambda _event: self.update_interval())
        self.interval_spin.bind("<FocusOut>", lambda _event: self.update_interval())

        tk.Label(
            footer,
            text="透明度",
            bg=BG,
            fg=MUTED,
            font=("Microsoft YaHei UI", 9),
        ).pack(side="left", padx=(12, 5))

        self.alpha_spin = tk.Spinbox(
            footer,
            from_=0.35,
            to=1.0,
            increment=0.05,
            width=5,
            format="%.2f",
            textvariable=self.alpha_var,
            command=self.update_alpha,
            bg=INPUT_BG,
            fg=TEXT,
            buttonbackground=PANEL_2,
            insertbackground=TEXT,
            relief="flat",
            font=("Consolas", 9),
        )
        self.alpha_spin.pack(side="left")
        self.alpha_spin.bind("<Return>", lambda _event: self.update_alpha())
        self.alpha_spin.bind("<FocusOut>", lambda _event: self.update_alpha())

        tk.Label(
            footer,
            text="报警%",
            bg=BG,
            fg=MUTED,
            font=("Microsoft YaHei UI", 9),
        ).pack(side="left", padx=(12, 5))

        self.alert_threshold_spin = tk.Spinbox(
            footer,
            from_=0.1,
            to=50.0,
            increment=0.1,
            width=5,
            format="%.1f",
            textvariable=self.alert_threshold_var,
            command=self.update_alert_threshold,
            bg=INPUT_BG,
            fg=TEXT,
            buttonbackground=PANEL_2,
            insertbackground=TEXT,
            relief="flat",
            font=("Consolas", 9),
        )
        self.alert_threshold_spin.pack(side="left")
        self.alert_threshold_spin.bind("<Return>", lambda _event: self.update_alert_threshold())
        self.alert_threshold_spin.bind("<FocusOut>", lambda _event: self.update_alert_threshold())

        self._make_body_button(footer, "立即更新", self.refresh_now, width=9).pack(side="right")

    def _make_header_button(
        self,
        text: str,
        command,
        width: int | None = None,
    ) -> tk.Button:
        return tk.Button(
            self.header,
            text=text,
            command=command,
            width=width,
            bg=PANEL_2,
            fg=TEXT,
            activebackground=BORDER,
            activeforeground=TEXT,
            bd=0,
            relief="flat",
            cursor="hand2",
            font=("Microsoft YaHei UI", 9),
        )

    def _make_body_button(
        self,
        parent: tk.Widget,
        text: str,
        command,
        width: int | None = None,
    ) -> tk.Button:
        return tk.Button(
            parent,
            text=text,
            command=command,
            width=width,
            bg=ACCENT,
            fg="#ffffff",
            activebackground="#6ca2ff",
            activeforeground="#ffffff",
            bd=0,
            relief="flat",
            cursor="hand2",
            font=("Microsoft YaHei UI", 9),
        )

    def _table_label(
        self,
        parent: tk.Widget,
        text: str,
        width: int,
        anchor: str,
    ) -> tk.Label:
        return tk.Label(
            parent,
            text=text,
            width=width,
            anchor=anchor,
            bg=BG,
            fg=MUTED,
            font=("Microsoft YaHei UI", 8),
        )

    def restore_default_size(self, _event: tk.Event | None = None) -> str:
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        x = min(max(self.root.winfo_x(), 0), max(screen_w - WINDOW_W, 0))
        y = min(max(self.root.winfo_y(), 0), max(screen_h - WINDOW_H, 0))
        self.root.geometry(f"{WINDOW_W}x{WINDOW_H}+{x}+{y}")
        self._save_current_config()
        self._force_topmost()
        return "break"

    def _force_topmost(self) -> None:
        if not bool(self.topmost_var.get()):
            return
        force_tk_window_topmost(self.root)

    def _schedule_topmost_guard(self) -> None:
        if self.topmost_after_id is not None:
            return
        self.topmost_after_id = self.root.after(TOPMOST_REFRESH_MS, self._topmost_guard)

    def _topmost_guard(self) -> None:
        self.topmost_after_id = None
        self._force_topmost()
        self._schedule_topmost_guard()

    def _on_rows_configure(self, _event: tk.Event) -> None:
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event: tk.Event) -> None:
        self.canvas.itemconfigure(self.rows_window, width=event.width)

    def _on_mousewheel(self, event: tk.Event) -> None:
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _ensure_selected_symbol(self) -> None:
        if self.selected_symbol in self.symbols:
            return
        self.selected_symbol = self.symbols[0] if self.symbols else ""

    def _select_symbol(self, symbol: str) -> None:
        symbol = normalize_symbol(symbol)
        if not symbol or symbol not in self.symbols:
            return
        self.selected_symbol = symbol
        self._save_current_config()
        self._render_rows()

    def add_symbols(self) -> None:
        new_symbols = split_symbols(self.symbol_entry.get())
        if not new_symbols:
            self.status_var.set("请输入股票代码")
            return

        added = 0
        for symbol in new_symbols:
            if symbol not in self.symbols:
                self.symbols.append(symbol)
                added += 1

        self.symbol_entry.delete(0, "end")
        if added:
            if not self.selected_symbol:
                self.selected_symbol = new_symbols[0]
            self._save_current_config()
            self._render_rows()
            self.status_var.set(f"已添加 {added} 个")
            self.refresh_now()
        else:
            self.status_var.set("代码已存在")

    def remove_symbol(self, symbol: str) -> None:
        self.symbols = [item for item in self.symbols if item != symbol]
        self.quotes.pop(symbol, None)
        self.previous_alert_prices.pop(symbol, None)
        self.last_alert_at.pop(symbol, None)
        self._ensure_selected_symbol()
        self._save_current_config()
        self._render_rows()
        self.status_var.set(f"已删除 {symbol}")
        self._schedule_refresh(delay_ms=self._refresh_seconds() * 1000)

    def toggle_topmost(self) -> None:
        self.root.attributes("-topmost", bool(self.topmost_var.get()))
        self._force_topmost()
        self._save_current_config()

    def update_interval(self) -> None:
        seconds = self._refresh_seconds()
        self.interval_var.set(str(seconds))
        self._save_current_config()
        self._schedule_refresh(delay_ms=seconds * 1000)

    def update_alpha(self) -> None:
        alpha = self._alpha()
        self.alpha_var.set(f"{alpha:.2f}")
        self.root.attributes("-alpha", alpha)
        self._save_current_config()

    def update_alert_threshold(self) -> None:
        threshold = self._alert_threshold_pct()
        self.alert_threshold_var.set(f"{threshold:.1f}")
        self._save_current_config()

    def _refresh_seconds(self) -> int:
        try:
            seconds = int(float(self.interval_var.get()))
        except (TypeError, ValueError):
            seconds = 15
        return min(max(seconds, MIN_REFRESH_SECONDS), MAX_REFRESH_SECONDS)

    def _alpha(self) -> float:
        try:
            alpha = float(self.alpha_var.get())
        except (TypeError, ValueError):
            alpha = 0.86
        return min(max(alpha, 0.35), 1.0)

    def _alert_threshold_pct(self) -> float:
        try:
            threshold = float(self.alert_threshold_var.get())
        except (TypeError, ValueError):
            threshold = 1.0
        return min(max(threshold, 0.1), 50.0)

    def refresh_now(self) -> None:
        self._schedule_refresh(delay_ms=1)

    def _schedule_refresh(self, delay_ms: int) -> None:
        if self.refresh_after_id is not None:
            try:
                self.root.after_cancel(self.refresh_after_id)
            except tk.TclError:
                pass
            self.refresh_after_id = None
        self.refresh_after_id = self.root.after(max(delay_ms, 1), self._start_refresh)

    def _start_refresh(self) -> None:
        self.refresh_after_id = None
        if self.fetching:
            self._schedule_refresh(delay_ms=1000)
            return

        if not self.symbols:
            self.status_var.set("输入代码后添加")
            self._schedule_refresh(delay_ms=self._refresh_seconds() * 1000)
            return

        self.fetching = True
        self.status_var.set("更新中...")
        self._ensure_selected_symbol()
        symbols_snapshot = list(self.symbols)
        thread = threading.Thread(
            target=self._fetch_worker,
            args=(symbols_snapshot,),
            daemon=True,
        )
        thread.start()
        self.root.after(100, self._poll_results)

    def _fetch_worker(self, symbols: list[str]) -> None:
        quotes: dict[str, Quote] | None = None
        quote_error: str | None = None
        try:
            quotes = fetch_quotes(symbols)
        except Exception as exc:  # noqa: BLE001 - show network/provider failures in the UI.
            logging.exception("Fetch quotes failed")
            quote_error = str(exc)

        self.result_queue.put((quotes, quote_error))

    def _poll_results(self) -> None:
        try:
            quotes, error = self.result_queue.get_nowait()
        except queue.Empty:
            if self.fetching:
                self.root.after(100, self._poll_results)
            return

        self.fetching = False
        alerts: list[str] = []
        if error:
            self.status_var.set("更新失败，保留上次")
            logging.warning("Keeping previous quotes after fetch error: %s", error)
        elif quotes is not None:
            accepted_quotes, skipped_symbols = self._filter_quotes_for_update(quotes)
            alerts = self._collect_price_alerts(accepted_quotes)
            self.quotes.update(accepted_quotes)
            source = next((quote.source for quote in accepted_quotes.values() if quote.source), "")
            prefix = f"{source} " if source else ""
            if accepted_quotes:
                suffix = f"，保留旧价 {len(skipped_symbols)}" if skipped_symbols else ""
                self.status_var.set(f"{prefix}{datetime.now().strftime('%H:%M:%S')}{suffix}")
            else:
                self.status_var.set("未取得当前价，保留上次")

        self._render_rows()
        if alerts:
            self.root.after(50, lambda messages=alerts: self._show_price_alerts(messages))
        self._schedule_refresh(delay_ms=self._refresh_seconds() * 1000)

    def _filter_quotes_for_update(self, quotes: dict[str, Quote]) -> tuple[dict[str, Quote], list[str]]:
        accepted: dict[str, Quote] = {}
        skipped: list[str] = []
        now_et = current_eastern_datetime()

        for symbol, quote in quotes.items():
            if not quote_matches_current_session(quote, now_et):
                skipped.append(symbol)
                continue

            previous_time = quote_active_time(self.quotes.get(symbol))
            current_time = quote_active_time(quote)
            if previous_time is not None and current_time is not None and current_time < previous_time:
                skipped.append(symbol)
                continue

            accepted[symbol] = quote

        return accepted, skipped

    def _quote_alert_price(self, quote: Quote) -> tuple[float | None, str]:
        if quote.extended_price is not None:
            label = STATE_LABELS.get(quote.extended_state, quote.extended_state or "当前")
            return quote.extended_price, f"{label}价"
        label = STATE_LABELS.get(quote.market_state, quote.market_state or "当前")
        return quote.price, f"{label}价"

    def _collect_price_alerts(self, quotes: dict[str, Quote]) -> list[str]:
        threshold = self._alert_threshold_pct()
        cooldown = int(self.config.get("alert_cooldown_seconds", 300))
        now = datetime.now().timestamp()
        alerts: list[str] = []

        for symbol in self.symbols:
            quote = quotes.get(symbol)
            if quote is None or quote.error:
                continue

            current_price, price_label = self._quote_alert_price(quote)
            if current_price is None:
                continue

            previous_price = self.previous_alert_prices.get(symbol)
            self.previous_alert_prices[symbol] = current_price
            if previous_price in (None, 0):
                continue

            change_pct = (current_price - previous_price) / previous_price * 100
            if abs(change_pct) <= threshold:
                continue

            last_at = self.last_alert_at.get(symbol, 0)
            if cooldown and now - last_at < cooldown:
                continue

            self.last_alert_at[symbol] = now
            direction = "上涨" if change_pct > 0 else "下跌"
            alerts.append(
                f"{symbol} {price_label}{direction} {change_pct:+.2f}%\n"
                f"上次: {format_price(previous_price)}  当前: {format_price(current_price)}"
            )

        return alerts

    def _show_price_alerts(self, alerts: list[str]) -> None:
        if not alerts:
            return
        try:
            alert = tk.Toplevel(self.root)
            alert.withdraw()
            alert.title("美股价格警报")
            alert.configure(bg=BG)
            alert.attributes("-topmost", True)
            alert.transient(self.root)
            alert.resizable(False, False)

            frame = tk.Frame(alert, bg=BG, padx=16, pady=14)
            frame.pack(fill="both", expand=True)

            title = tk.Label(
                frame,
                text="价格变动超过阈值",
                bg=BG,
                fg=TEXT,
                anchor="w",
                font=("Microsoft YaHei UI", 11, "bold"),
            )
            title.pack(fill="x")

            body = tk.Label(
                frame,
                text="\n\n".join(alerts),
                bg=BG,
                fg=TEXT,
                anchor="w",
                justify="left",
                wraplength=380,
                font=("Microsoft YaHei UI", 10),
            )
            body.pack(fill="both", expand=True, pady=(10, 14))

            def close_alert(_event: tk.Event | None = None) -> str:
                try:
                    alert.grab_release()
                except tk.TclError:
                    pass
                try:
                    alert.destroy()
                except tk.TclError:
                    pass
                return "break"

            ok_button = tk.Button(
                frame,
                text="知道了",
                command=close_alert,
                bg=ACCENT,
                fg="#ffffff",
                activebackground="#3d7bd8",
                activeforeground="#ffffff",
                bd=0,
                padx=18,
                pady=6,
            )
            ok_button.pack(anchor="e")

            alert.bind("<Return>", close_alert)
            alert.bind("<Escape>", close_alert)
            alert.protocol("WM_DELETE_WINDOW", close_alert)

            alert.update_idletasks()
            screen_w = alert.winfo_screenwidth()
            window_w = min(max(alert.winfo_reqwidth(), 360), 460)
            window_h = min(max(alert.winfo_reqheight(), 160), 520)
            x = max(screen_w - window_w - 18, 0)
            y = 18
            alert.geometry(f"{window_w}x{window_h}+{x}+{y}")
            alert.deiconify()

            def keep_alert_on_top() -> None:
                try:
                    if not alert.winfo_exists():
                        return
                    force_tk_window_topmost(alert, activate=True)
                    alert.after(1000, keep_alert_on_top)
                except tk.TclError:
                    return

            try:
                alert.grab_set()
            except tk.TclError:
                pass
            keep_alert_on_top()
        except tk.TclError:
            logging.exception("Failed to show price alert")

    def _render_rows(self) -> None:
        self._ensure_selected_symbol()
        for child in self.rows_frame.winfo_children():
            child.destroy()

        if not self.symbols:
            empty = tk.Label(
                self.rows_frame,
                text="还没有自选股",
                bg=BG,
                fg=MUTED,
                font=("Microsoft YaHei UI", 11),
                pady=36,
            )
            empty.pack(fill="x")
            return

        for symbol in self.symbols:
            quote = self.quotes.get(symbol, Quote(symbol=symbol))
            self._render_quote_row(symbol, quote)

    def _render_quote_row(self, symbol: str, quote: Quote) -> None:
        row_bg = PANEL_2 if symbol == self.selected_symbol else PANEL
        has_extended = quote.extended_price is not None
        row_span = 2 if has_extended else 1
        row = tk.Frame(
            self.rows_frame,
            bg=row_bg,
            highlightbackground=ACCENT if symbol == self.selected_symbol else BORDER,
            highlightthickness=1,
        )
        row.pack(fill="x", pady=(0, 5))

        for col, weight in enumerate((2, 2, 2, 1, 0)):
            row.columnconfigure(col, weight=weight)

        name = short_text(quote.name, 16) if quote.name else ""
        symbol_box = tk.Frame(row, bg=row_bg)
        symbol_box.grid(row=0, column=0, rowspan=row_span, sticky="ew", padx=(8, 4), pady=6)
        symbol_label = tk.Label(
            symbol_box,
            text=symbol,
            bg=row_bg,
            fg=TEXT,
            anchor="w",
            font=("Consolas", 11, "bold"),
        )
        symbol_label.pack(fill="x")
        name_label = tk.Label(
            symbol_box,
            text=name or " ",
            bg=row_bg,
            fg=MUTED,
            anchor="w",
            font=("Microsoft YaHei UI", 8),
        )
        name_label.pack(fill="x")

        price_color = TEXT if not quote.error else MUTED
        price_label = tk.Label(
            row,
            text=format_price(quote.price),
            bg=row_bg,
            fg=price_color,
            anchor="e",
            font=("Consolas", 12, "bold"),
        )
        price_label.grid(row=0, column=1, sticky="ew", padx=4, pady=6)

        change_color = MUTED
        if quote.change is not None:
            change_color = UP if quote.change >= 0 else DOWN
        elif quote.change_pct is not None:
            change_color = UP if quote.change_pct >= 0 else DOWN

        change_text = format_change(quote.change, quote.change_pct)
        if quote.error:
            change_text = quote.error
            change_color = MUTED

        change_label = tk.Label(
            row,
            text=short_text(change_text, 18),
            bg=row_bg,
            fg=change_color,
            anchor="e",
            font=("Consolas", 10),
        )
        change_label.grid(row=0, column=2, sticky="ew", padx=4, pady=6)

        state = "收盘" if has_extended else STATE_LABELS.get(quote.market_state, quote.market_state or "")
        time_text = quote.display_time or format_market_time(quote.market_time)
        time_label = tk.Label(
            row,
            text=f"{state}\n{time_text}".strip(),
            bg=row_bg,
            fg=MUTED,
            anchor="e",
            justify="right",
            font=("Microsoft YaHei UI", 8),
        )
        time_label.grid(row=0, column=3, sticky="ew", padx=4, pady=6)

        extended_widgets: list[tk.Widget] = []
        if has_extended:
            extended_color = MUTED
            if quote.extended_change is not None:
                extended_color = UP if quote.extended_change >= 0 else DOWN
            elif quote.extended_change_pct is not None:
                extended_color = UP if quote.extended_change_pct >= 0 else DOWN

            extended_price_label = tk.Label(
                row,
                text=format_price(quote.extended_price),
                bg=row_bg,
                fg=extended_color,
                anchor="e",
                font=("Consolas", 12, "bold"),
            )
            extended_price_label.grid(row=1, column=1, sticky="ew", padx=4, pady=(0, 6))

            extended_change_label = tk.Label(
                row,
                text=short_text(
                    format_change(quote.extended_change, quote.extended_change_pct),
                    18,
                ),
                bg=row_bg,
                fg=extended_color,
                anchor="e",
                font=("Consolas", 10),
            )
            extended_change_label.grid(row=1, column=2, sticky="ew", padx=4, pady=(0, 6))

            extended_state = STATE_LABELS.get(quote.extended_state, quote.extended_state or "")
            extended_time_label = tk.Label(
                row,
                text=f"{extended_state}\n{format_market_time(quote.extended_time)}".strip(),
                bg=row_bg,
                fg=MUTED,
                anchor="e",
                justify="right",
                font=("Microsoft YaHei UI", 8),
            )
            extended_time_label.grid(row=1, column=3, sticky="ew", padx=4, pady=(0, 6))
            extended_widgets.extend(
                (extended_price_label, extended_change_label, extended_time_label)
            )

        delete_button = tk.Button(
            row,
            text="删",
            command=lambda s=symbol: self.remove_symbol(s),
            bg=row_bg,
            fg=MUTED,
            activebackground=BORDER,
            activeforeground=TEXT,
            bd=0,
            relief="flat",
            cursor="hand2",
            font=("Microsoft YaHei UI", 9),
        )
        delete_button.grid(row=0, column=4, rowspan=row_span, sticky="e", padx=(2, 6), pady=6)

        for widget in (
            row,
            symbol_box,
            symbol_label,
            name_label,
            price_label,
            change_label,
            time_label,
            *extended_widgets,
        ):
            widget.bind("<Button-1>", lambda _event, s=symbol: self._select_symbol(s))

        row.bind("<MouseWheel>", self._on_mousewheel)
        for child in row.winfo_children():
            child.bind("<MouseWheel>", self._on_mousewheel)

    def _save_current_config(self) -> None:
        self.config["symbols"] = self.symbols
        self.config["refresh_seconds"] = self._refresh_seconds()
        self.config["topmost"] = bool(self.topmost_var.get())
        self.config["alpha"] = self._alpha()
        self.config["alert_threshold_pct"] = self._alert_threshold_pct()
        self.config["selected_symbol"] = self.selected_symbol
        self.config["geometry"] = sanitize_geometry(
            self.root.geometry(),
            self.root.winfo_screenwidth(),
            self.root.winfo_screenheight(),
        )
        save_config(self.config)

    def close(self) -> None:
        if self.refresh_after_id is not None:
            try:
                self.root.after_cancel(self.refresh_after_id)
            except tk.TclError:
                pass
        if self.topmost_after_id is not None:
            try:
                self.root.after_cancel(self.topmost_after_id)
            except tk.TclError:
                pass
        self._save_current_config()
        self.root.destroy()


def main() -> None:
    root = tk.Tk()
    StockFloatApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
