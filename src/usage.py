"""AI利用額の計算・記録（バロメーター用）

各API呼び出しの usage（トークン数・Web検索回数）から概算コストを計算し、
output/usage_log.jsonl に1行ずつ追記する。月の合計や期間合計を集計できる。

※あくまで概算。実際の請求は Anthropic Console の Usage / Billing で確認すること。
"""

from __future__ import annotations

import datetime as dt
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
LOG_PATH = ROOT / "output" / "usage_log.jsonl"

# 100万トークンあたりの単価（USD）。入力 / 出力。
# キャッシュ書込=入力×1.25、キャッシュ読込=入力×0.1 で概算。
PRICES = {
    "claude-opus-4-8": {"in": 5.0, "out": 25.0},
    "claude-opus-4-7": {"in": 5.0, "out": 25.0},
    "claude-opus-4-6": {"in": 5.0, "out": 25.0},
    "claude-sonnet-4-6": {"in": 3.0, "out": 15.0},
    "claude-haiku-4-5": {"in": 1.0, "out": 5.0},
}
DEFAULT_PRICE = {"in": 5.0, "out": 25.0}
WEB_SEARCH_USD_PER_1K = 10.0  # Web検索ツールの概算単価


def cost_usd(model: str, usage) -> tuple[float, dict]:
    """1回のレスポンスの usage から概算コスト(USD)と内訳を返す。"""
    p = PRICES.get(model, DEFAULT_PRICE)
    inp = getattr(usage, "input_tokens", 0) or 0
    out = getattr(usage, "output_tokens", 0) or 0
    cw = getattr(usage, "cache_creation_input_tokens", 0) or 0
    cr = getattr(usage, "cache_read_input_tokens", 0) or 0
    usd = (inp * p["in"] + out * p["out"] + cw * p["in"] * 1.25 + cr * p["in"] * 0.1) / 1_000_000

    web = 0
    stu = getattr(usage, "server_tool_use", None)
    if stu is not None:
        web = getattr(stu, "web_search_requests", 0) or 0
    usd += web * WEB_SEARCH_USD_PER_1K / 1000

    detail = {"in": inp, "out": out, "cache_write": cw, "cache_read": cr, "web_search": web}
    return usd, detail


def record(cfg: dict, model: str, usage, kind: str) -> float:
    """利用を記録し、その回の概算コスト(JPY)を返す。"""
    rate = cfg.get("billing", {}).get("usd_to_jpy", 160)
    usd, detail = cost_usd(model, usage)
    jpy = usd * rate
    now = dt.datetime.now()
    entry = {
        "ts": now.isoformat(timespec="seconds"),
        "month": now.strftime("%Y-%m"),
        "kind": kind,
        "model": model,
        "usd": round(usd, 6),
        "jpy": round(jpy, 2),
        "detail": detail,
    }
    # まずスプレッドシートへ（再起動で消えない）。設定が無ければローカルJSONLへ。
    try:
        from . import sheets
    except ImportError:
        import sheets  # type: ignore

    if not sheets.append(entry):
        LOG_PATH.parent.mkdir(exist_ok=True)
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(json.dumps(entry, ensure_ascii=False) + "\n")
    return jpy


def _load() -> list[dict]:
    # スプレッドシートが使えればそちらを正とする。
    try:
        from . import sheets
    except ImportError:
        import sheets  # type: ignore
    rows = sheets.load()
    if rows is not None:
        return rows

    if not LOG_PATH.exists():
        return []
    out = []
    for line in LOG_PATH.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if line:
            try:
                out.append(json.loads(line))
            except json.JSONDecodeError:
                continue
    return out


def month_total_jpy(month: str | None = None) -> float:
    """指定月（未指定なら今月）の概算合計(JPY)。"""
    month = month or dt.datetime.now().strftime("%Y-%m")
    return round(sum(e["jpy"] for e in _load() if e.get("month") == month), 1)


def since_total_jpy(ts_iso: str) -> float:
    """指定時刻以降の概算合計(JPY)。1回の生成分の合計に使う。"""
    return round(sum(e["jpy"] for e in _load() if e.get("ts", "") >= ts_iso), 1)


def now_iso() -> str:
    return dt.datetime.now().isoformat(timespec="seconds")
