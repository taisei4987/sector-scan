#!/usr/bin/env python3
"""
東証プライム 33業種・銘柄の値動きスキャン（デイトレ向きセクター／銘柄探し用）

データ元: Yahoo Finance（yfinance）… 東証は約20分遅延
業種区分: JPX「東証上場銘柄一覧」（data_j.xls）の33業種区分

モード
  morning : プライム全銘柄の「前日」データ。売買代金上位リストも作る
  yori    : 寄り付き後の当日データ（9:30実行で9:10ごろまでの値動き）。売買代金上位銘柄のみ
  zenba   : 当日の前場データ（11:50以降推奨）。売買代金上位銘柄のみ
  close   : 当日の大引けデータ（15:50以降推奨）。売買代金上位銘柄のみ
  auto    : 実行時刻（日本時間）から自動判定

出力（data/ フォルダ、UTF-8 BOM付きCSV）
  YYYY-MM-DD_<mode>_sectors.csv : 33業種ごとの集計
  YYYY-MM-DD_<mode>_stocks.csv  : 銘柄ごとの指標
"""
from __future__ import annotations

import argparse
import io
import os
import re
import sys
import time
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

import numpy as np
import pandas as pd

JST = timezone(timedelta(hours=9))
JPX_PAGE = "https://www.jpx.co.jp/markets/statistics-equities/misc/01.html"
JPX_FALLBACK = (
    "https://www.jpx.co.jp/markets/statistics-equities/misc/"
    "tvdivq0000001vg2-att/data_j.xls"
)
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
UNIVERSE_FILE = DATA_DIR / "universe.csv"       # プライム全銘柄＋業種（キャッシュ）
TOPN_FILE = DATA_DIR / "top_turnover.csv"       # 朝に作る売買代金上位リスト
WATCHLIST_FILE = BASE_DIR / "watchlist.txt"     # 毎回必ず対象に含める銘柄コード
SECTOR_TREND_FILE = DATA_DIR / "sector_trend.csv"  # 朝に作る業種の環境（プライム全銘柄）
UNIVERSE_MAX_AGE_DAYS = 7
MODE_LABEL = {"morning": "朝（前日データ）", "yori": "寄り後", "zenba": "前場後", "close": "大引け後"}


# ---------------------------------------------------------------- モード判定
def resolve_mode(mode: str, now: datetime) -> str:
    if mode != "auto":
        return mode
    if now.weekday() >= 5:                       # 土日は直近営業日の振り返り
        return "morning"
    hm = now.hour * 60 + now.minute
    if hm < 9 * 60:
        return "morning"
    if hm < 10 * 60 + 30:
        return "yori"
    if hm < 12 * 60 + 30:
        return "zenba"
    return "close"


def timing_warning(mode: str, now: datetime) -> str | None:
    hm = now.hour * 60 + now.minute
    if now.weekday() >= 5:
        return None
    if mode == "yori" and hm < 9 * 60 + 20:
        return "9:20より前のため、当日のデータがまだほとんどありません（約20分遅延）。"
    if mode == "yori":
        return f"約20分遅延のため、{(now - timedelta(minutes=20)):%H:%M}ごろまでの値動きです。出来高倍率は1日平均との比較なので小さく出ます。"
    if mode == "zenba" and hm < 11 * 60 + 50:
        return "11:50より前のため、前場の最後の約20分が反映されていません。"
    if mode == "close" and hm < 15 * 60 + 50:
        return "15:50より前のため、大引けまでのデータがそろっていません（場中の値）。"
    return None


# ---------------------------------------------------------------- 銘柄一覧
def _find_col(df: pd.DataFrame, *keywords: str) -> str:
    for c in df.columns:
        if all(k in str(c) for k in keywords):
            return c
    raise KeyError(f"列が見つかりません: {keywords} / 列一覧: {list(df.columns)}")


def parse_jpx_list(raw: pd.DataFrame) -> pd.DataFrame:
    """JPXの上場銘柄一覧からプライム（内国株式）だけを取り出す。"""
    c_code = _find_col(raw, "コード")
    c_name = _find_col(raw, "銘柄名")
    c_mkt = _find_col(raw, "市場")
    c_sec = _find_col(raw, "33業種区分")
    df = raw[[c_code, c_name, c_mkt, c_sec]].copy()
    df.columns = ["コード", "銘柄名", "市場", "業種"]
    df = df[df["市場"].astype(str).str.contains("プライム")
            & df["市場"].astype(str).str.contains("内国")]
    df["コード"] = df["コード"].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
    df["ticker"] = df["コード"] + ".T"
    return df[["コード", "銘柄名", "業種", "ticker"]].reset_index(drop=True)


def _read_excel_bytes(content: bytes) -> pd.DataFrame:
    """中身の先頭バイトで形式を判定して読む（.xls / .xlsx）。"""
    if content[:4] == b"\xd0\xcf\x11\xe0":
        return pd.read_excel(io.BytesIO(content), dtype=str, engine="xlrd")
    if content[:2] == b"PK":
        return pd.read_excel(io.BytesIO(content), dtype=str, engine="openpyxl")
    head = content[:200].decode("utf-8", errors="replace").replace("\n", " ")
    raise ValueError(f"Excelファイルではありません（先頭: {head!r}）")


def _local_jpx_file() -> Path | None:
    """リポジトリに置いた data_j.xls / data_j.xlsx があればそれを使う。"""
    for d in (BASE_DIR, DATA_DIR):
        for name in ("data_j.xls", "data_j.xlsx"):
            if (d / name).exists():
                return d / name
    return None


def _save_universe(raw: pd.DataFrame, source: str) -> pd.DataFrame:
    uni = parse_jpx_list(raw)
    uni["取得日"] = datetime.now(JST).date().isoformat()
    DATA_DIR.mkdir(exist_ok=True)
    uni.to_csv(UNIVERSE_FILE, index=False, encoding="utf-8-sig")
    print(f"銘柄一覧: {source} からプライム {len(uni)}銘柄")
    return uni


def load_universe(force: bool = False) -> pd.DataFrame:
    local = _local_jpx_file()
    if local:
        return _save_universe(_read_excel_bytes(local.read_bytes()), local.name)

    cached = pd.read_csv(UNIVERSE_FILE, dtype=str) if UNIVERSE_FILE.exists() else None
    if cached is not None and not force:
        fetched = date.fromisoformat(cached["取得日"].iloc[0])
        if (datetime.now(JST).date() - fetched).days < UNIVERSE_MAX_AGE_DAYS:
            return cached

    import requests
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                             "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
               "Accept-Language": "ja,en;q=0.8"}
    urls = []
    try:
        r = requests.get(JPX_PAGE, headers=headers, timeout=30)
        print(f"JPXページ: HTTP {r.status_code}")
        for m in re.finditer(r'href=["\']([^"\']*data_j\.xlsx?)["\']', r.text):
            urls.append(requests.compat.urljoin(JPX_PAGE, m.group(1)))
    except Exception as e:
        print(f"[注意] JPXページ取得失敗: {e}", file=sys.stderr)
    urls.append(JPX_FALLBACK)

    for url in dict.fromkeys(urls):
        try:
            r = requests.get(url, headers=headers, timeout=60)
            print(f"銘柄一覧を取得中: {url} → HTTP {r.status_code}, {len(r.content)} bytes")
            return _save_universe(_read_excel_bytes(r.content), url)
        except Exception as e:
            print(f"[注意] 取得失敗: {e}", file=sys.stderr)

    if cached is not None:
        print("[注意] 銘柄一覧を更新できなかったため、前回の一覧を使います。", file=sys.stderr)
        return cached
    raise SystemExit(
        "JPXの銘柄一覧を自動で取得できませんでした。\n"
        "JPXサイト「東証上場銘柄一覧」から data_j.xls をダウンロードし、"
        "リポジトリのトップにアップロードしてから再実行してください。")


# ---------------------------------------------------------------- 株価取得
def _split_download(data: pd.DataFrame, tickers: list[str]) -> dict[str, pd.DataFrame]:
    out: dict[str, pd.DataFrame] = {}
    if data is None or data.empty:
        return out
    if isinstance(data.columns, pd.MultiIndex):
        lv0 = set(data.columns.get_level_values(0))
        for t in tickers:
            if t in lv0:
                df = data[t]
            elif t in set(data.columns.get_level_values(1)):
                df = data.xs(t, axis=1, level=1)
            else:
                continue
            df = df.dropna(subset=["Close"])
            if not df.empty:
                out[t] = df
    elif len(tickers) == 1:
        df = data.dropna(subset=["Close"])
        if not df.empty:
            out[tickers[0]] = df
    return out


def download_prices(tickers: list[str], chunk: int = 100, period: str = "3mo") -> dict[str, pd.DataFrame]:
    import yfinance as yf

    result: dict[str, pd.DataFrame] = {}
    pending = list(tickers)
    for attempt in range(2):                       # 失敗分は1回だけ再試行
        failed: list[str] = []
        for i in range(0, len(pending), chunk):
            part = pending[i:i + chunk]
            try:
                data = yf.download(part, period=period, interval="1d", group_by="ticker",
                                   auto_adjust=False, threads=True, progress=False)
            except Exception as e:
                print(f"[注意] 取得エラー: {e}", file=sys.stderr)
                data = None
            got = _split_download(data, part)
            result.update(got)
            failed += [t for t in part if t not in got]
            print(f"  取得 {min(i + chunk, len(pending))}/{len(pending)}（成功累計 {len(result)}）")
            time.sleep(1.0)
        if not failed:
            break
        pending = failed
        if attempt == 0:
            print(f"取得失敗 {len(failed)}銘柄を再試行します…")
            time.sleep(5)
    return result


def load_watchlist() -> list[str]:
    if not WATCHLIST_FILE.exists():
        return []
    codes = []
    for line in WATCHLIST_FILE.read_text(encoding="utf-8").splitlines():
        code = line.split("#")[0].strip()
        if code:
            codes.append(code)
    return codes


def fetch_intraday(tickers: list[str], today: date, cutoff_hm: tuple[int, int] | None,
                   chunk: int = 100) -> dict[str, pd.DataFrame]:
    """当日の5分足を取得する（cutoff_hm を指定するとその時刻より前の足だけ）。"""
    import yfinance as yf

    out: dict[str, pd.DataFrame] = {}
    cut = (datetime(today.year, today.month, today.day, *cutoff_hm, tzinfo=JST)
           if cutoff_hm else None)
    for i in range(0, len(tickers), chunk):
        part = tickers[i:i + chunk]
        try:
            data = yf.download(part, period="1d", interval="5m", group_by="ticker",
                               auto_adjust=False, threads=True, progress=False)
        except Exception as e:
            print(f"[注意] 5分足の取得エラー: {e}", file=sys.stderr)
            continue
        for t, bars in _split_download(data, part).items():
            idx = bars.index
            idx = idx.tz_convert("Asia/Tokyo") if idx.tz is not None else idx.tz_localize("Asia/Tokyo")
            bars = bars.copy()
            bars.index = idx
            b = bars[bars.index.date == today]
            if cut is not None:
                b = b[b.index < cut]
            if not b.empty:
                out[t] = b
        time.sleep(1.0)
    return out


def sanitize_intraday(prices: dict[str, pd.DataFrame], bars_map: dict[str, pd.DataFrame],
                      today: date) -> int:
    """Yahooの5分足は、最初の足の始値・高値・安値が「前日終値」になっていることがある。
    日足の当日始値（正しい寄り付き値）で最初の足を補正する。戻り値は補正した銘柄数。"""
    fixed = 0
    for t, b in bars_map.items():
        df = prices.get(t)
        if df is None or b.empty:
            continue
        dates = [pd.Timestamp(d).date() for d in df.index]
        if today not in dates or dates.index(today) == 0:
            continue
        i = dates.index(today)
        d_open, prev_close = df["Open"].iloc[i], df["Close"].iloc[i - 1]
        if not (d_open == d_open and prev_close == prev_close) or not d_open:
            continue
        near = lambda v: v == v and abs(v / prev_close - 1) < 1e-4
        # 寄り付き前の「前日終値だけの足」（4本値すべて前日終値）は取り除く
        lead = 0
        while lead < len(b) - 1 and all(near(b[k].iloc[lead]) for k in ("Open", "High", "Low", "Close")):
            lead += 1
        if lead:
            b = b.iloc[lead:]
            bars_map[t] = b
            fixed += 1
        first = b.iloc[0].copy()
        if abs(first["Open"] / d_open - 1) < 1e-6 or abs(first["Open"] / d_open - 1) > 0.3:
            continue                                   # 始値は正しい／日足と5分足が別物すぎる
        c = first["Close"]
        if near(first["High"]) and not near(d_open):
            first["High"] = max(d_open, c)
        if near(first["Low"]) and not near(d_open):
            first["Low"] = min(d_open, c)
        first["Open"] = d_open                         # 始値は日足の正式な始値に合わせる
        first["High"] = max(first["High"], d_open)
        first["Low"] = min(first["Low"], d_open)
        b = b.copy()
        b.iloc[0] = first
        bars_map[t] = b
        fixed += 1
    return fixed


def replace_today_row(prices: dict[str, pd.DataFrame], bars_map: dict[str, pd.DataFrame],
                      today: date) -> int:
    """5分足から作った当日OHLCVで、日足の当日行を置き換える（前場後モード用）。"""
    replaced = 0
    for t, b in bars_map.items():
        if t not in prices:
            continue
        row = {"Open": b["Open"].iloc[0], "High": b["High"].max(), "Low": b["Low"].min(),
               "Close": b["Close"].iloc[-1], "Volume": b["Volume"].sum()}
        row["Adj Close"] = row["Close"]
        df = prices[t]
        df = df[[pd.Timestamp(d).date() != today for d in df.index]].copy()
        new = pd.DataFrame([row], index=[pd.Timestamp(today)])
        prices[t] = pd.concat([df, new[[c for c in df.columns if c in new.columns]]]).sort_index()
        replaced += 1
    return replaced


# ---------------------------------------------------------------- 基調の判定
TREND_RULE = ("環境（日足・前日まで）：上昇基調＝終値が25日線より上 かつ 25日線が上向き（5日前より上） "
              "かつ 5日線が25日線より上／下落基調＝その逆／それ以外＝横ばい")
FLOW_MIN_OPEN = 0.3   # 始値比がこれ未満（±%）の小さな動きは方向とみなさない
FLOW_MIN_VWAP = 0.1   # VWAP比がこれ未満（±%）の小さな乖離は方向とみなさない
FLOW_RULE = (f"当日の流れ（5分足）：買い優勢＝始値より{FLOW_MIN_OPEN}%以上上 かつ VWAPより{FLOW_MIN_VWAP}%以上上 "
             "かつ 高値・安値が切り下げでない／売り優勢＝その逆／それ以外＝方向感なし。"
             "高安は前半と後半の高値・安値の比較。指数はVWAPなし")


def daily_trend(df: pd.DataFrame) -> dict:
    """日足（前日までの確定分）から環境を判定する。"""
    c = df["Close"].dropna()
    if len(c) < 30:
        return {"環境": "判定不可", "25日線比%": np.nan, "25日線の向き%": np.nan,
                "5日騰落%": np.nan, "20日騰落%": np.nan}
    ma5, ma25 = c.rolling(5).mean(), c.rolling(25).mean()
    last, m25, m25_prev = c.iloc[-1], ma25.iloc[-1], ma25.iloc[-6]
    slope = (m25 / m25_prev - 1) * 100
    if last > m25 and slope > 0 and ma5.iloc[-1] > m25:
        env = "上昇基調"
    elif last < m25 and slope < 0 and ma5.iloc[-1] < m25:
        env = "下落基調"
    else:
        env = "横ばい"
    return {"環境": env,
            "25日線比%": round((last / m25 - 1) * 100, 2),
            "25日線の向き%": round(slope, 2),
            "5日騰落%": round((last / c.iloc[-6] - 1) * 100, 2),
            "20日騰落%": round((last / c.iloc[-21] - 1) * 100, 2)}


NO_FLOW = {"当日の流れ": "—", "始値比%": np.nan, "VWAP比%": np.nan, "高安": "—", "前日レンジ": "—"}


def intraday_flow(bars: pd.DataFrame | None, prev_high: float, prev_low: float) -> dict:
    """当日の5分足から当日の流れを判定する。"""
    if bars is None or bars.dropna(subset=["Close"]).empty:
        return dict(NO_FLOW, 当日の流れ="判定不可")
    b = bars.dropna(subset=["Close"])
    o, last = b["Open"].iloc[0], b["Close"].iloc[-1]
    vol = b["Volume"].fillna(0)
    tp = (b["High"] + b["Low"] + b["Close"]) / 3
    vwap = (tp * vol).sum() / vol.sum() if vol.sum() > 0 else np.nan
    from_open = (last / o - 1) * 100 if o else np.nan
    vwap_dev = (last / vwap - 1) * 100 if vwap == vwap and vwap else np.nan
    if len(b) >= 4:
        h = len(b) // 2
        a, z = b.iloc[:h], b.iloc[h:]
        up = z["High"].max() > a["High"].max() and z["Low"].min() > a["Low"].min()
        down = z["High"].max() < a["High"].max() and z["Low"].min() < a["Low"].min()
        struct = "切り上げ" if up else "切り下げ" if down else "もみ合い"
    else:
        struct = "データ不足"
    if last > prev_high:
        pos = "前日高値超え"
    elif last < prev_low:
        pos = "前日安値割れ"
    else:
        pos = "前日レンジ内"
    vw_ok_up = np.isnan(vwap_dev) or vwap_dev >= FLOW_MIN_VWAP
    vw_ok_dn = np.isnan(vwap_dev) or vwap_dev <= -FLOW_MIN_VWAP
    if from_open >= FLOW_MIN_OPEN and vw_ok_up and struct != "切り下げ":
        flow = "買い優勢"
    elif from_open <= -FLOW_MIN_OPEN and vw_ok_dn and struct != "切り上げ":
        flow = "売り優勢"
    else:
        flow = "方向感なし"
    return {"当日の流れ": flow,
            "始値比%": round(from_open, 2),
            "VWAP比%": round(vwap_dev, 2) if vwap_dev == vwap_dev else np.nan,
            "高安": struct, "前日レンジ": pos}


# ---------------------------------------------------------------- 指標計算
def calc_metrics(df: pd.DataFrame, mode: str, today: date,
                 bars: pd.DataFrame | None = None) -> dict | None:
    df = df.sort_index()
    idx_dates = pd.Index([pd.Timestamp(d).date() for d in df.index])
    df = df.copy()
    df.index = idx_dates
    flag = ""
    if mode == "morning":
        df = df[df.index < today]                  # 当日の途中データは使わない
    else:
        if len(df) == 0 or df.index[-1] != today:
            flag = "当日データなし"
    if len(df) < 16:
        return None

    cur, prev = df.iloc[-1], df.iloc[-2]
    hist = df.iloc[:-1]                             # 当日を含まない過去
    pc = prev["Close"]
    if not pc or np.isnan(pc):
        return None

    tr = pd.concat([
        hist["High"] - hist["Low"],
        (hist["High"] - hist["Close"].shift()).abs(),
        (hist["Low"] - hist["Close"].shift()).abs(),
    ], axis=1).max(axis=1)
    atr = tr.tail(14).mean()

    typical = (df["High"] + df["Low"] + df["Close"]) / 3
    turnover = typical * df["Volume"] / 1e8          # 売買代金（概算・億円）
    vol_avg = hist["Volume"].tail(20).mean()

    rng_pct = (cur["High"] - cur["Low"]) / pc * 100
    atr_pct = atr / pc * 100
    m = {
        "データ日付": df.index[-1].isoformat(),
        "現在値": round(float(cur["Close"]), 1),
        "前日比%": round((cur["Close"] / pc - 1) * 100, 2),
        "ギャップ%": round((cur["Open"] / pc - 1) * 100, 2) if cur["Open"] else np.nan,
        "始値比_日足%": round((cur["Close"] / cur["Open"] - 1) * 100, 2) if cur["Open"] else np.nan,
        "日中値幅%": round(rng_pct, 2),
        "ATR14%": round(atr_pct, 2),
        "値幅/ATR": round(rng_pct / atr_pct, 2) if atr_pct else np.nan,
        "出来高": int(cur["Volume"]) if not np.isnan(cur["Volume"]) else np.nan,
        "出来高倍率": round(cur["Volume"] / vol_avg, 2) if vol_avg else np.nan,
        "売買代金_億円": round(float(turnover.iloc[-1]), 1),
        "売買代金20日平均_億円": round(float(turnover.iloc[:-1].tail(20).mean()), 1),
        "備考": flag,
    }
    # 環境：朝は前日まで（＝当日行なし）、それ以外は当日行を除いた前日まで
    base = df if mode == "morning" or flag else hist
    m.update(daily_trend(base))
    if mode == "morning" or flag:
        m.update(NO_FLOW)
    else:
        m.update(intraday_flow(bars, float(prev["High"]), float(prev["Low"])))
    return m


def build_stock_table(prices: dict[str, pd.DataFrame], uni: pd.DataFrame, mode: str,
                      today: date, bars_map: dict[str, pd.DataFrame] | None = None) -> pd.DataFrame:
    rows = []
    meta = uni.set_index("ticker")
    bars_map = bars_map or {}
    for t, df in prices.items():
        m = calc_metrics(df, mode, today, bars_map.get(t))
        if m is None or t not in meta.index:
            continue
        rows.append({"コード": meta.at[t, "コード"], "銘柄名": meta.at[t, "銘柄名"],
                     "業種": meta.at[t, "業種"], **m})
    return pd.DataFrame(rows)


MARKET_TICKERS = {"^N225": "日経平均", "1306.T": "TOPIX連動ETF（1306）"}


def build_market_table(mode: str, today: date, cutoff_hm: tuple[int, int] | None) -> pd.DataFrame:
    """日経平均とTOPIX連動ETFの環境・当日の流れ。"""
    tickers = list(MARKET_TICKERS)
    prices = download_prices(tickers)
    bars_map = fetch_intraday(tickers, today, cutoff_hm) if mode != "morning" else {}
    sanitize_intraday(prices, bars_map, today)
    now = datetime.now(JST)
    if cutoff_hm and now.hour * 60 + now.minute >= 12 * 60 + 50:
        replace_today_row(prices, bars_map, today)
    rows = []
    for t, name in MARKET_TICKERS.items():
        m = calc_metrics(prices[t], mode, today, bars_map.get(t)) if t in prices else None
        if m is None:
            rows.append({"指標": name, "環境": "未確認", "当日の流れ": "未確認"})
            continue
        rows.append({"指標": name, **m})
    return pd.DataFrame(rows)


def _label_sector_trend(g: pd.DataFrame) -> str:
    up = (g["環境"] == "上昇基調").mean()
    down = (g["環境"] == "下落基調").mean()
    above = (g["25日線比%"] > 0).mean()
    r20 = g["20日騰落%"].mean()
    is_up = up >= 0.5 or (above >= 0.6 and r20 > 0)
    is_down = down >= 0.5 or (above <= 0.4 and r20 < 0)
    if is_up and not is_down:
        return "上昇基調"
    if is_down and not is_up:
        return "下落基調"
    if is_up and is_down:
        return "上昇基調" if r20 > 0 else "下落基調"
    return "横ばい"


def build_sector_trend(stocks: pd.DataFrame, today: date) -> pd.DataFrame:
    rows = []
    for sec, g in stocks.groupby("業種"):
        g = g[g["環境"].isin(["上昇基調", "下落基調", "横ばい"])]
        if g.empty:
            continue
        rows.append({"業種": sec, "業種の環境": _label_sector_trend(g),
                     "上昇基調の割合%": round((g["環境"] == "上昇基調").mean() * 100),
                     "下落基調の割合%": round((g["環境"] == "下落基調").mean() * 100),
                     "25日線より上の割合%": round((g["25日線比%"] > 0).mean() * 100),
                     "平均5日騰落%": round(g["5日騰落%"].mean(), 2),
                     "平均20日騰落%": round(g["20日騰落%"].mean(), 2),
                     "判定銘柄数": len(g), "作成日": today.isoformat()})
    return pd.DataFrame(rows)


def sector_flow(stocks: pd.DataFrame) -> pd.Series:
    out = {}
    for sec, g in stocks.groupby("業種"):
        f = g["当日の流れ"]
        n = f.isin(["買い優勢", "売り優勢", "方向感なし"]).sum()
        if n == 0:
            out[sec] = "—"
            continue
        buy, sell = (f == "買い優勢").sum(), (f == "売り優勢").sum()
        lbl = ("買い優勢" if buy / n >= 0.5 and buy > sell else
               "売り優勢" if sell / n >= 0.5 and sell > buy else "方向感なし")
        out[sec] = f"{lbl}（買{buy}・売{sell}／{n}）"
    return pd.Series(out)


def build_sector_table(stocks: pd.DataFrame) -> pd.DataFrame:
    g = stocks.groupby("業種")
    sec = pd.DataFrame({
        "構成銘柄数": g.size(),
        "平均騰落率%": g["前日比%"].mean().round(2),
        "中央値騰落率%": g["前日比%"].median().round(2),
        "上昇数": g["前日比%"].apply(lambda s: int((s > 0).sum())),
        "下落数": g["前日比%"].apply(lambda s: int((s < 0).sum())),
        "平均日中値幅%": g["日中値幅%"].mean().round(2),
        "平均値幅/ATR": g["値幅/ATR"].mean().round(2),
        "売買代金合計_億円": g["売買代金_億円"].sum().round(1),
    })
    # 代表銘柄 = 売買代金上位2銘柄
    rep = (stocks.sort_values("売買代金_億円", ascending=False)
           .groupby("業種").head(2)
           .assign(lbl=lambda d: d["銘柄名"] + "(" + d["前日比%"].map("{:+.2f}%".format) + ")")
           .groupby("業種")["lbl"].apply(" / ".join))
    sec["代表銘柄"] = rep
    sec = sec.sort_values("平均騰落率%", ascending=False)
    sec.insert(0, "順位", range(1, len(sec) + 1))
    n = len(sec)
    sec["区分"] = ""
    sec.loc[sec["順位"] <= 3, "区分"] = "上位3"
    sec.loc[sec["順位"] > n - 3, "区分"] = "下位3"
    big = (sec["区分"] == "") & (sec["平均騰落率%"].abs() >= 2)
    sec.loc[big, "区分"] = "±2%以上"
    return sec.reset_index()


def sort_stocks(stocks: pd.DataFrame, sectors: pd.DataFrame) -> pd.DataFrame:
    rank = sectors.set_index("業種")["順位"]
    s = stocks.copy()
    s.insert(0, "業種順位", s["業種"].map(rank))
    return s.sort_values(["業種順位", "日中値幅%"], ascending=[True, False]).reset_index(drop=True)


# ---------------------------------------------------------------- まとめ（Actions画面用）
def markdown_summary(mode: str, now: datetime, sectors: pd.DataFrame, stocks: pd.DataFrame,
                     warn: str | None, min_turnover: float) -> str:
    L = [f"## 33業種スキャン：{MODE_LABEL[mode]}",
         f"取得 {now:%Y-%m-%d %H:%M} JST（Yahoo Finance・約20分遅延）／対象 {len(stocks)}銘柄", ""]
    if warn:
        L += [f"> ⚠ {warn}", ""]
    if (stocks["備考"] == "当日データなし").mean() > 0.5:
        L += ["> ⚠ 多くの銘柄で当日データがありません（休場日の可能性）。", ""]
    L += ["| 順位 | 業種 | 騰落率 | 日中値幅 | 値幅/ATR | 売買代金(億) | 銘柄数 |",
          "|---|---|---|---|---|---|---|"]
    show = sectors[sectors["区分"] != ""]
    for _, r in show.iterrows():
        L.append(f"| {r['順位']} | {r['業種']} | {r['平均騰落率%']:+.2f}% | {r['平均日中値幅%']:.2f}% | "
                 f"{r['平均値幅/ATR']:.2f} | {r['売買代金合計_億円']:,.0f} | {r['構成銘柄数']} |")
    L += ["", f"### 上位・下位業種の値幅が大きい銘柄（売買代金{min_turnover:g}億円以上）"]
    for _, r in show.iterrows():
        cand = stocks[(stocks["業種"] == r["業種"]) & (stocks["売買代金_億円"] >= min_turnover)]
        cand = cand.nlargest(5, "日中値幅%")
        if cand.empty:
            continue
        L.append(f"**{r['業種']}**（{r['平均騰落率%']:+.2f}%）")
        for _, s in cand.iterrows():
            L.append(f"- {s['コード']} {s['銘柄名']}：前日比 {s['前日比%']:+.2f}%／値幅 {s['日中値幅%']:.2f}%"
                     f"／値幅/ATR {s['値幅/ATR']:.2f}／出来高倍率 {s['出来高倍率']:.2f}"
                     f"／売買代金 {s['売買代金_億円']:,.0f}億")
        L.append("")
    return "\n".join(L)


def _f(v, fmt: str = "{}") -> str:
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "—"
    try:
        return fmt.format(v)
    except (ValueError, TypeError):
        return str(v)


def _row(cells: list[str]) -> str:
    return "| " + " | ".join(cells) + " |"


def _head(cols: list[str]) -> list[str]:
    return [_row(cols), _row(["---"] * len(cols))]


STOCK_COLS = ["コード", "銘柄名", "業種", "現在値", "前日比", "ギャップ", "日中値幅", "値幅/ATR", "出来高倍率",
              "売買代金(億)", "環境", "25日線比", "当日の流れ", "始値比", "VWAP比", "高安", "前日レンジ", "備考"]


def _stock_line(s: pd.Series) -> str:
    return _row([str(s["コード"]), str(s["銘柄名"]), str(s["業種"]), _f(s["現在値"], "{:,.1f}"),
                 _f(s["前日比%"], "{:+.2f}%"), _f(s["ギャップ%"], "{:+.2f}%"), _f(s["日中値幅%"], "{:.2f}%"), _f(s["値幅/ATR"], "{:.2f}"),
                 _f(s["出来高倍率"], "{:.2f}"), _f(s["売買代金_億円"], "{:,.0f}"), _f(s["環境"]),
                 _f(s["25日線比%"], "{:+.2f}%"), _f(s["当日の流れ"]), _f(s["始値比%"], "{:+.2f}%"),
                 _f(s["VWAP比%"], "{:+.2f}%"), _f(s["高安"]), _f(s["前日レンジ"]), _f(s["備考"] or "")])


def report_markdown(mode: str, now: datetime, stocks: pd.DataFrame, sectors: pd.DataFrame,
                    market: pd.DataFrame, sec_trend: pd.DataFrame, watch: list[str],
                    basis: str, min_turnover: float) -> str:
    """Claudeのレポートが読む用のまとめ（data/latest_<mode>.md）。"""
    data_date = stocks["データ日付"].mode().iloc[0] if not stocks.empty else ""
    trend_date = sec_trend["作成日"].iloc[0] if not sec_trend.empty else "未作成"
    L = [f"# 33業種スキャン {MODE_LABEL[mode]}",
         "",
         f"- 取得時刻: {now:%Y-%m-%d %H:%M} JST",
         f"- データの日付: {data_date}",
         f"- 数値の時点: {basis}",
         "- 出典: Yahoo Finance（yfinance、東証は約20分遅延）",
         "- 売買代金は (高値+安値+終値)/3×出来高 の概算。出来高倍率は過去20日平均との比較",
         "",
         "## デイトレ視点",
         "",
         "機械的なルールによる分類で、売買の推奨ではない。",
         "",
         f"- {TREND_RULE}",
         f"- {FLOW_RULE}" + ("（朝モードは当日データがないため「—」）" if mode == "morning" else ""),
         "",
         "### 市場全体", ""]
    L += _head(["指標", "現在値", "前日比", "環境", "25日線比", "25日線の向き", "当日の流れ",
                "始値比", "高安", "前日レンジ"])
    for _, r in market.iterrows():
        g = r.get
        L.append(_row([str(g("指標")), _f(g("現在値"), "{:,.1f}"), _f(g("前日比%"), "{:+.2f}%"),
                       _f(g("環境")), _f(g("25日線比%"), "{:+.2f}%"), _f(g("25日線の向き%"), "{:+.2f}%"),
                       _f(g("当日の流れ")), _f(g("始値比%"), "{:+.2f}%"), _f(g("高安")), _f(g("前日レンジ"))]))

    L += ["", "### 業種（平均日中値幅の大きい順）", "",
          f"騰落率は構成銘柄の単純平均（業種の順位は公式の業種別指数を使うこと）。"
          f"業種の環境は {trend_date} 朝のプライム全銘柄から判定。", ""]
    sec = sectors.set_index("業種")
    tr = sec_trend.set_index("業種") if not sec_trend.empty else pd.DataFrame()
    flow = sector_flow(stocks)
    L += _head(["業種", "騰落率", "平均値幅", "値幅/ATR", "売買代金(億)", "銘柄数", "業種の環境",
                "25日線より上", "平均20日騰落", "当日の流れ（銘柄数）"])
    for name, r in sec.sort_values("平均日中値幅%", ascending=False).iterrows():
        t = tr.loc[name] if name in tr.index else {}
        L.append(_row([str(name), _f(r["平均騰落率%"], "{:+.2f}%"), _f(r["平均日中値幅%"], "{:.2f}%"),
                       _f(r["平均値幅/ATR"], "{:.2f}"), _f(r["売買代金合計_億円"], "{:,.0f}"),
                       str(r["構成銘柄数"]), _f(t.get("業種の環境") if len(t) else None),
                       _f(t.get("25日線より上の割合%") if len(t) else None, "{:.0f}%"),
                       _f(t.get("平均20日騰落%") if len(t) else None, "{:+.2f}%"),
                       _f(flow.get(name))]))

    liquid = stocks[stocks["売買代金_億円"] >= min_turnover]
    L += ["", f"### 候補銘柄：日中値幅が大きい 上位15（売買代金{min_turnover:g}億円以上）", ""]
    L += _head(STOCK_COLS) + [_stock_line(s) for _, s in liquid.nlargest(15, "日中値幅%").iterrows()]
    L += ["", f"### 候補銘柄：いつもより動いている 上位10（値幅/ATR順、売買代金{min_turnover:g}億円以上）", ""]
    L += _head(STOCK_COLS) + [_stock_line(s) for _, s in liquid.nlargest(10, "値幅/ATR").iterrows()]
    gap = liquid.assign(_g=liquid["ギャップ%"].abs()).nlargest(10, "_g")
    L += ["", f"### 候補銘柄：ギャップ（寄り付きの窓）が大きい 上位10（|ギャップ|順、売買代金{min_turnover:g}億円以上）",
          "", "ギャップ＝始値の前日終値比。日中値幅にはギャップ分が含まれないため、別に並べる。", ""]
    L += _head(STOCK_COLS) + [_stock_line(s) for _, s in gap.iterrows()]

    L += ["", "## ウォッチリスト", ""] + _head(STOCK_COLS)
    w = stocks.drop_duplicates("コード").set_index("コード", drop=False)
    for code in watch:
        if code in w.index:
            L.append(_stock_line(w.loc[code]))
        else:
            L.append(_row([code, "取得できず"] + [""] * (len(STOCK_COLS) - 3) + ["未確認"]))
    return "\n".join(L) + "\n"


# ---------------------------------------------------------------- メイン
def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="東証プライム 33業種・銘柄スキャン")
    ap.add_argument("--mode", default="auto", choices=["auto", "morning", "yori", "zenba", "close"])
    ap.add_argument("--top", type=int, default=300, help="前場後・大引け後に使う売買代金上位の銘柄数")
    ap.add_argument("--min-turnover", type=float, default=10.0,
                    help="まとめに載せる銘柄の売買代金下限（億円）")
    ap.add_argument("--refresh-universe", action="store_true", help="JPX銘柄一覧を再取得する")
    args = ap.parse_args(argv)

    now = datetime.now(JST)
    today = now.date()
    mode = resolve_mode(args.mode, now)
    warn = timing_warning(mode, now)
    print(f"モード: {MODE_LABEL[mode]}  ({now:%Y-%m-%d %H:%M} JST)")
    if warn:
        print(f"[注意] {warn}")

    uni = load_universe(force=args.refresh_universe)
    watch = load_watchlist()

    if mode == "morning":
        target = uni
    else:
        target = uni
        if TOPN_FILE.exists():
            top = pd.read_csv(TOPN_FILE, dtype=str)
            target = uni[uni["ticker"].isin(top["ticker"])]
            print(f"売買代金上位リスト（{top['作成日'].iloc[0]}作成）の {len(target)}銘柄を対象にします")
            if top["作成日"].iloc[0] != today.isoformat():
                print("[注意] 上位リストが今日の朝のものではありません。朝モードを先に実行すると更新されます。")
        else:
            print("[注意] 売買代金上位リストがないため、プライム全銘柄を対象にします。")
        watch_t = [c + ".T" for c in watch]
        target = uni[uni["ticker"].isin(set(target["ticker"]) | set(watch_t))]

    tickers = target["ticker"].tolist()
    prices = download_prices(tickers)
    basis = {"morning": "前日の終値（確定値）",
             "yori": f"約20分遅延のため {(now - timedelta(minutes=20)):%H:%M}ごろまで",
             "zenba": "前場（〜11:30）",
             "close": "大引け（確定値）"}[mode]
    cutoff = (11, 30) if mode == "zenba" else None
    bars_map: dict[str, pd.DataFrame] = {}
    if mode != "morning":
        bars_map = fetch_intraday(tickers, today, cutoff)
        print(f"当日の5分足を取得: {len(bars_map)}/{len(tickers)}銘柄")
    if bars_map:
        print(f"5分足の最初の足を補正: {sanitize_intraday(prices, bars_map, today)}銘柄")
    # 12:50より前なら、20分遅延の日足はまだ前場の値だけなので日足をそのまま使う
    if mode == "zenba" and now.hour * 60 + now.minute >= 12 * 60 + 50:
        n = replace_today_row(prices, bars_map, today)
        print(f"前場の5分足で当日値を置き換え: {n}/{len(prices)}銘柄")
        if n < len(prices) * 0.5:
            basis = f"日足（{(now - timedelta(minutes=20)):%H:%M}ごろまで。後場が混ざっている可能性あり）"
    stocks = build_stock_table(prices, uni, mode, today, bars_map)
    if stocks.empty:
        print("データが取得できませんでした。", file=sys.stderr)
        return 1

    sectors = build_sector_table(stocks)
    stocks = sort_stocks(stocks, sectors)
    stamp = now.strftime("%Y-%m-%d %H:%M")
    for df in (sectors, stocks):
        df["取得時刻JST"] = stamp
        df["出典"] = "Yahoo Finance(約20分遅延)"

    DATA_DIR.mkdir(exist_ok=True)
    base = DATA_DIR / f"{today.isoformat()}_{mode}"
    sectors.to_csv(f"{base}_sectors.csv", index=False, encoding="utf-8-sig")
    stocks.to_csv(f"{base}_stocks.csv", index=False, encoding="utf-8-sig")
    print(f"出力: {base}_sectors.csv / {base}_stocks.csv")

    if mode == "morning":
        top = stocks.nlargest(args.top, "売買代金20日平均_億円")[["コード", "銘柄名", "業種"]].copy()
        top["ticker"] = top["コード"] + ".T"
        top["作成日"] = today.isoformat()
        top.to_csv(TOPN_FILE, index=False, encoding="utf-8-sig")
        print(f"売買代金上位 {len(top)}銘柄のリストを更新: {TOPN_FILE.name}")
        sec_trend = build_sector_trend(stocks, today)
        sec_trend.to_csv(SECTOR_TREND_FILE, index=False, encoding="utf-8-sig")
        print(f"業種の環境を更新: {SECTOR_TREND_FILE.name}")
    elif SECTOR_TREND_FILE.exists():
        sec_trend = pd.read_csv(SECTOR_TREND_FILE)
    else:
        print("[注意] 業種の環境ファイルがないため、対象銘柄だけで判定します。")
        sec_trend = build_sector_trend(stocks, today)

    market = build_market_table(mode, today, cutoff)

    report = report_markdown(mode, now, stocks, sectors, market, sec_trend, watch, basis,
                             args.min_turnover)
    (DATA_DIR / f"latest_{mode}.md").write_text(report, encoding="utf-8")
    print(f"レポート用まとめ: data/latest_{mode}.md")

    summary = markdown_summary(mode, now, sectors, stocks, warn, args.min_turnover)
    summary += "\n\n" + report.split("\n", 1)[1]
    print("\n" + summary)
    if os.environ.get("GITHUB_STEP_SUMMARY"):
        with open(os.environ["GITHUB_STEP_SUMMARY"], "a", encoding="utf-8") as f:
            f.write(summary + "\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())
