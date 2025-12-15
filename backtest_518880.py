#!/usr/bin/env python3
"""Backtest 518880.SH (黄金ETF) with simple, robust mid/long-term rules.

Design goals:
- Minimal data requirements: only ETF daily OHLCV from TuShare.
- Avoid look-ahead: signals computed on rebalance day; positions applied next day.
- Practical: monthly rebalance, transaction cost, export results.

Security:
- Reads TuShare token from env var TUSHARE_TOKEN. DO NOT hardcode credentials.

Example:
  export TUSHARE_TOKEN="..."
  python3 backtest_518880.py --start 20130101 --tc 0.001
"""

from __future__ import annotations

import argparse
import os
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd


def _today_yyyymmdd() -> str:
    return datetime.now().strftime("%Y%m%d")


def _ensure_outdir(outdir: str) -> Path:
    p = Path(outdir)
    p.mkdir(parents=True, exist_ok=True)
    return p


def _get_pro():
    token = os.getenv("TUSHARE_TOKEN", "").strip()
    if not token:
        raise SystemExit(
            "Missing TuShare token. Set env var TUSHARE_TOKEN first.\n"
            "Example: export TUSHARE_TOKEN=\"your_token\""
        )

    import tushare as ts  # imported here to keep startup clean

    return ts.pro_api(token)


def fetch_fund_daily(
    ts_code: str,
    start_date: str,
    end_date: str,
    cache_csv: str | None = None,
) -> pd.DataFrame:
    """Fetch ETF/fund daily data from TuShare.

    Returns a DataFrame indexed by datetime with at least:
      open, high, low, close, vol, amount
    """

    if cache_csv and Path(cache_csv).exists():
        df = pd.read_csv(cache_csv)
        if "trade_date" in df.columns:
            df["trade_date"] = pd.to_datetime(df["trade_date"], format="%Y%m%d")
            df = df.sort_values("trade_date")
            df = df.set_index("trade_date")
        return df

    pro = _get_pro()

    # TuShare may cap rows per request; chunk by year to be safe.
    start_y = int(start_date[:4])
    end_y = int(end_date[:4])
    frames: list[pd.DataFrame] = []

    for y in range(start_y, end_y + 1):
        s = start_date if y == start_y else f"{y}0101"
        e = end_date if y == end_y else f"{y}1231"
        try:
            part = pro.fund_daily(ts_code=ts_code, start_date=s, end_date=e)
        except Exception as exc:  # noqa: BLE001
            raise SystemExit(f"TuShare fund_daily failed for {ts_code} {s}-{e}: {exc}")

        if part is None or part.empty:
            continue
        frames.append(part)

    if not frames:
        raise SystemExit(f"No data returned for {ts_code} {start_date}-{end_date}")

    df = pd.concat(frames, ignore_index=True)
    # Standardize
    df["trade_date"] = pd.to_datetime(df["trade_date"], format="%Y%m%d")
    df = df.sort_values("trade_date").drop_duplicates(subset=["trade_date"], keep="last")
    df = df.set_index("trade_date")

    keep = [c for c in ["open", "high", "low", "close", "vol", "amount"] if c in df.columns]
    df = df[keep].copy()

    if cache_csv:
        Path(cache_csv).parent.mkdir(parents=True, exist_ok=True)
        out = df.reset_index()
        out["trade_date"] = out["trade_date"].dt.strftime("%Y%m%d")
        out.to_csv(cache_csv, index=False)

    return df


def try_fetch_usdcny(
    start_date: str,
    end_date: str,
    cache_csv: str | None = None,
) -> pd.Series | None:
    """Optional: fetch USD/CNY (if available in your TuShare plan).

    We keep this best-effort: if endpoint/data isn't available, return None.

    Output: Series indexed by trade_date with column name 'usdcny'.
    """

    if cache_csv and Path(cache_csv).exists():
        fx = pd.read_csv(cache_csv)
        if "trade_date" in fx.columns and "usdcny" in fx.columns:
            fx["trade_date"] = pd.to_datetime(fx["trade_date"], format="%Y%m%d")
            fx = fx.sort_values("trade_date").set_index("trade_date")
            return fx["usdcny"].astype(float)

    pro = _get_pro()

    # Different TuShare accounts may have different FX endpoints.
    # We attempt a couple of common ones.
    candidates = [
        ("fx_daily", {"ts_code": "USDCNY.FX", "start_date": start_date, "end_date": end_date}),
        ("fx_daily", {"ts_code": "USDCNY", "start_date": start_date, "end_date": end_date}),
    ]

    fx_df = None
    last_err = None
    for api_name, kwargs in candidates:
        try:
            api = getattr(pro, api_name)
            tmp = api(**kwargs)
            if tmp is not None and not tmp.empty:
                fx_df = tmp
                break
        except Exception as exc:  # noqa: BLE001
            last_err = exc
            continue

    if fx_df is None or fx_df.empty:
        return None

    # Heuristic column mapping
    fx_df["trade_date"] = pd.to_datetime(fx_df["trade_date"], format="%Y%m%d")
    fx_df = fx_df.sort_values("trade_date").set_index("trade_date")

    if "close" in fx_df.columns:
        s = fx_df["close"].astype(float)
    elif "bid" in fx_df.columns:
        s = fx_df["bid"].astype(float)
    else:
        return None

    s.name = "usdcny"

    if cache_csv:
        Path(cache_csv).parent.mkdir(parents=True, exist_ok=True)
        out = s.reset_index()
        out["trade_date"] = out["trade_date"].dt.strftime("%Y%m%d")
        out.to_csv(cache_csv, index=False)

    return s


def last_trading_day_of_period(dates: pd.DatetimeIndex, freq: str) -> pd.Series:
    """Return boolean Series marking last trading day in each period."""

    if freq == "M":
        key = dates.to_period("M")
    elif freq == "W":
        key = dates.to_period("W")
    else:
        raise ValueError("freq must be 'M' or 'W'")

    # last date per period
    df = pd.DataFrame({"k": key}, index=dates)
    last = df.groupby("k", sort=False).tail(1).index
    flag = pd.Series(False, index=dates)
    flag.loc[last] = True
    return flag


@dataclass(frozen=True)
class StrategyResult:
    name: str
    equity: pd.Series
    position: pd.Series
    trades: pd.Series
    stats: dict


def compute_indicators(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()

    out["ret"] = out["close"].pct_change().fillna(0.0)
    out["logret"] = np.log(out["close"]).diff().fillna(0.0)

    # 10-month MA ~ 210 trading days
    out["ma10"] = out["close"].rolling(210, min_periods=210).mean()

    # 12-1 momentum: close(t-21)/close(t-252) - 1
    out["mom12_1"] = out["close"].shift(21) / out["close"].shift(252) - 1.0

    # Realized vol (annualized) using log returns
    out["vol63"] = out["logret"].rolling(63, min_periods=63).std() * np.sqrt(252)

    # Vol regime: 3-year rolling 80% quantile
    out["vol_q80_3y"] = out["vol63"].rolling(756, min_periods=252).quantile(0.8)

    # Drawdown relative to trailing 252d max
    rolling_max = out["close"].rolling(252, min_periods=252).max()
    out["dd1y"] = out["close"] / rolling_max - 1.0

    return out


def backtest_single_asset(
    df: pd.DataFrame,
    target_w: pd.Series,
    tc_one_way: float,
    name: str,
) -> StrategyResult:
    """Backtest ETF + cash(0) with target weights (rebalanced discretely).

    Convention:
    - target_w is desired weight decided on rebalance days.
    - execution occurs next trading day (avoid look-ahead).
    """

    # Apply next-day execution
    pos = target_w.reindex(df.index).ffill().fillna(0.0).shift(1).fillna(0.0)

    # Turnover and trading cost on days position changes
    delta = pos.diff().abs().fillna(pos.abs())
    cost = delta * tc_one_way

    port_ret = pos * df["ret"] - cost
    equity = (1.0 + port_ret).cumprod()

    trades = (delta > 1e-12).astype(int)
    stats = perf_stats(equity, port_ret, pos, delta)

    return StrategyResult(name=name, equity=equity, position=pos, trades=trades, stats=stats)


def max_drawdown(equity: pd.Series) -> float:
    peak = equity.cummax()
    dd = equity / peak - 1.0
    return float(dd.min())


def perf_stats(
    equity: pd.Series,
    daily_ret: pd.Series,
    pos: pd.Series,
    turnover: pd.Series,
) -> dict:
    n = len(equity)
    if n < 3:
        return {}

    ann = 252
    cagr = float(equity.iloc[-1] ** (ann / n) - 1.0)
    vol = float(daily_ret.std(ddof=0) * np.sqrt(ann))
    sharpe = float((daily_ret.mean() / (daily_ret.std(ddof=0) + 1e-12)) * np.sqrt(ann))
    mdd = max_drawdown(equity)
    calmar = float(cagr / (abs(mdd) + 1e-12))
    avg_pos = float(pos.mean())
    time_in = float((pos > 0).mean())
    ann_turnover = float(turnover.sum() / (n / ann))

    return {
        "CAGR": cagr,
        "AnnVol": vol,
        "Sharpe(0rf)": sharpe,
        "MaxDD": mdd,
        "Calmar": calmar,
        "AvgPos": avg_pos,
        "TimeInMkt": time_in,
        "AnnTurnover": ann_turnover,
        "EndEquity": float(equity.iloc[-1]),
    }


def make_targets(
    df: pd.DataFrame,
    rebalance_flag: pd.Series,
    target_vol: float,
    use_fx: bool,
    usdcny: pd.Series | None,
) -> dict[str, pd.Series]:
    idx = df.index

    # helper: build target weights only on rebalance dates
    def _t(x: pd.Series) -> pd.Series:
        out = pd.Series(np.nan, index=idx, dtype=float)
        out.loc[rebalance_flag] = x.loc[rebalance_flag].astype(float)
        return out

    trend = (df["close"] > df["ma10"]).astype(float)

    # A: MA10
    w_a = _t(trend)

    # B: MA10 + 12-1 momentum confirmation
    w_b = _t(((trend > 0) & (df["mom12_1"] > 0)).astype(float))

    # C: Vol target on top of MA10 direction
    vol = df["vol63"].replace(0.0, np.nan)
    w_c_raw = trend * (target_vol / vol)
    w_c = _t(w_c_raw.clip(lower=0.0, upper=1.0).fillna(0.0))

    # Optional: extra risk-off filter (halve exposure in high-vol regimes)
    # Applied only on rebalance dates to keep it simple.
    hi_vol = (df["vol63"] > df["vol_q80_3y"]).fillna(False)
    reb_idx = idx[rebalance_flag.values]
    for w in (w_a, w_b, w_c):
        hv = hi_vol.reindex(reb_idx).fillna(False)
        w.loc[reb_idx] = np.where(hv.values, w.loc[reb_idx].values * 0.5, w.loc[reb_idx].values)

    targets: dict[str, pd.Series] = {
        "A_MA10M": w_a,
        "B_MA10M_MOM12_1": w_b,
        "C_MA10M_VOLTARGET": w_c,
    }

    # D: MA10 + FX filter (best-effort)
    if use_fx and usdcny is not None:
        fx = usdcny.reindex(idx).ffill()
        fxmom = fx / fx.shift(63) - 1.0
        w_d = _t(((trend > 0) & (fxmom > 0)).astype(float))
        hv = hi_vol.reindex(reb_idx).fillna(False)
        w_d.loc[reb_idx] = np.where(hv.values, w_d.loc[reb_idx].values * 0.5, w_d.loc[reb_idx].values)
        targets["D_MA10M_FXMOM"] = w_d

    return targets


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--ts-code", default="518880.SH", help="ETF ts_code, default 518880.SH")
    ap.add_argument("--start", default="20130101", help="start date YYYYMMDD")
    ap.add_argument("--end", default=_today_yyyymmdd(), help="end date YYYYMMDD")
    ap.add_argument("--rebalance", choices=["monthly", "weekly"], default="monthly")
    ap.add_argument("--tc", type=float, default=0.001, help="one-way transaction cost, e.g. 0.001=10bp")
    ap.add_argument("--target-vol", type=float, default=0.12, help="annualized target vol for strategy C")
    ap.add_argument("--outdir", default="out", help="output directory")
    ap.add_argument("--cache", action="store_true", help="cache downloaded data to ./out/cache")
    ap.add_argument("--with-fx", action="store_true", help="try add USD/CNY filter strategy (best-effort)")

    args = ap.parse_args()

    outdir = _ensure_outdir(args.outdir)
    cache_dir = outdir / "cache"
    cache_dir.mkdir(parents=True, exist_ok=True)

    cache_etf = str(cache_dir / f"{args.ts_code.replace('.', '_')}_{args.start}_{args.end}.csv") if args.cache else None
    df = fetch_fund_daily(args.ts_code, args.start, args.end, cache_csv=cache_etf)

    if df.empty:
        raise SystemExit("ETF daily data is empty")
    if "close" not in df.columns:
        raise SystemExit("ETF daily data missing 'close'")

    df = compute_indicators(df)

    # Define rebalance schedule
    freq = "M" if args.rebalance == "monthly" else "W"
    reb = last_trading_day_of_period(df.index, freq=freq)

    usdcny = None
    if args.with_fx:
        cache_fx = str(cache_dir / f"USDCNY_{args.start}_{args.end}.csv") if args.cache else None
        usdcny = try_fetch_usdcny(args.start, args.end, cache_csv=cache_fx)

    targets = make_targets(df, rebalance_flag=reb, target_vol=args.target_vol, use_fx=args.with_fx, usdcny=usdcny)

    # Baseline: buy & hold
    bh_target = pd.Series(np.nan, index=df.index, dtype=float)
    bh_target.loc[reb] = 1.0

    results: list[StrategyResult] = []
    results.append(backtest_single_asset(df, bh_target, tc_one_way=0.0, name="BUY_HOLD"))

    for name, tw in targets.items():
        results.append(backtest_single_asset(df, tw, tc_one_way=args.tc, name=name))

    # Export
    perf = pd.DataFrame([{"Strategy": r.name, **r.stats} for r in results]).set_index("Strategy")
    perf.to_csv(outdir / "perf_summary.csv")

    nav = pd.DataFrame({r.name: r.equity for r in results})
    nav.to_csv(outdir / "equity_curves.csv")

    pos = pd.DataFrame({r.name: r.position for r in results})
    pos.to_csv(outdir / "positions.csv")

    factors = df[["close", "ma10", "mom12_1", "vol63", "dd1y"]].copy()
    factors.to_csv(outdir / "factors.csv")

    # Plot (optional)
    try:
        import matplotlib.pyplot as plt

        ax = nav.plot(logy=True, figsize=(11, 6), title=f"{args.ts_code} strategies (log scale)")
        ax.set_xlabel("Date")
        ax.set_ylabel("Equity")
        plt.tight_layout()
        plt.savefig(outdir / "equity_curves.png", dpi=150)
        plt.close()
    except Exception:
        pass

    print("Saved:")
    print(f"  {outdir / 'perf_summary.csv'}")
    print(f"  {outdir / 'equity_curves.csv'}")
    print(f"  {outdir / 'positions.csv'}")
    print(f"  {outdir / 'factors.csv'}")
    if (outdir / "equity_curves.png").exists():
        print(f"  {outdir / 'equity_curves.png'}")

    print("\nPerformance summary:")
    with pd.option_context("display.max_columns", 50, "display.width", 200):
        print(perf.sort_values("Calmar", ascending=False))

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
