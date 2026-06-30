"""Googleスプレッドシートを利用ログの永続ストアとして使う薄いラッパー。

Streamlit Cloud はアプリが寝ると再起動時にローカルファイルを消すため、
利用量ログ（usage_log.jsonl）が消えてバロメーターが0%に戻る。
そこで記録先をスプレッドシートにして、再起動しても残るようにする。

必要な設定:
- 認証情報: ローカルは .sa-key.json、Streamlit Cloud は st.secrets["gcp_service_account"]
- シートID: 環境変数 USAGE_SHEET_ID もしくは st.secrets["USAGE_SHEET_ID"]
どちらか欠ければ None を返し、呼び出し側はローカルJSONLにフォールバックする。
"""

from __future__ import annotations

import json
import os
from pathlib import Path

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
ROOT = Path(__file__).resolve().parent.parent
KEY_PATH = ROOT / ".sa-key.json"
HEADER = ["ts", "month", "kind", "model", "usd", "jpy", "detail"]

_ws = None  # worksheet キャッシュ（モジュール存続中は使い回す）
_tried = False  # 接続を試みたか（失敗を毎回繰り返さない）


def _secret(name: str):
    try:
        import streamlit as st

        if name in st.secrets:
            return st.secrets[name]
    except Exception:
        pass
    return None


def _credentials_info() -> dict | None:
    info = _secret("gcp_service_account")
    if info is not None:
        return dict(info)
    if KEY_PATH.exists():
        return json.loads(KEY_PATH.read_text(encoding="utf-8"))
    return None


def _sheet_id() -> str | None:
    return os.environ.get("USAGE_SHEET_ID") or _secret("USAGE_SHEET_ID")


def worksheet():
    """設定が揃っていれば worksheet を返す。揃わなければ None。"""
    global _ws, _tried
    if _ws is not None:
        return _ws
    if _tried:
        return None
    _tried = True
    info = _credentials_info()
    sid = _sheet_id()
    if not info or not sid:
        return None
    try:
        import gspread
        from google.oauth2.service_account import Credentials

        creds = Credentials.from_service_account_info(info, scopes=SCOPES)
        gc = gspread.authorize(creds)
        ws = gc.open_by_key(sid).sheet1
        # ヘッダー行が無ければ用意する
        first = ws.row_values(1)
        if first != HEADER:
            if not first:
                ws.update(values=[HEADER], range_name="A1")
        _ws = ws
        return ws
    except Exception:
        return None


def append(entry: dict) -> bool:
    """1件追記。成功で True。設定不足・失敗で False（呼び出し側がフォールバック）。"""
    ws = worksheet()
    if ws is None:
        return False
    try:
        ws.append_row(
            [
                entry.get("ts", ""),
                entry.get("month", ""),
                entry.get("kind", ""),
                entry.get("model", ""),
                entry.get("usd", 0),
                entry.get("jpy", 0),
                json.dumps(entry.get("detail", {}), ensure_ascii=False),
            ],
            value_input_option="RAW",  # tsをISO文字列のまま保持（自動日付変換を防ぐ）
        )
        return True
    except Exception:
        return False


def load() -> list[dict] | None:
    """全件読み込み。設定不足・失敗なら None（呼び出し側がフォールバック）。"""
    ws = worksheet()
    if ws is None:
        return None
    try:
        records = ws.get_all_records()  # 1行目をヘッダーとして dict 化
        out = []
        for r in records:
            out.append(
                {
                    "ts": str(r.get("ts", "")),
                    "month": str(r.get("month", "")),
                    "kind": r.get("kind", ""),
                    "model": r.get("model", ""),
                    "jpy": float(r.get("jpy", 0) or 0),
                }
            )
        return out
    except Exception:
        return None
