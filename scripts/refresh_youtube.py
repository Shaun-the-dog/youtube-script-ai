"""YouTubeの実データ（タイトル・再生回数）を取り直して参照リストを更新する。

使い方: python scripts/refresh_youtube.py
必要: 環境変数 YOUTUBE_API_KEY（.env に記載）。

チャンネルの全公開動画を取得し、
- output/youtube_stats.json … 全動画のタイトル・再生回数・投稿日（分析用の生データ）
- config/reference_videos.txt … 再生回数トップ30（企画生成が学ぶ“伸びる型”）
を上書きする。
"""

from __future__ import annotations

import json
import os
import urllib.parse
import urllib.request
from pathlib import Path

from dotenv import load_dotenv

ROOT = Path(__file__).resolve().parent.parent
CHANNEL_ID = "UCpkG44X9sOjEoEklxsB4WMQ"  # かなえ弁護士
TOP_N = 30


def api(endpoint: str, key: str, **params) -> dict:
    params["key"] = key
    url = f"https://www.googleapis.com/youtube/v3/{endpoint}?" + urllib.parse.urlencode(params)
    with urllib.request.urlopen(url) as r:
        return json.load(r)


def fetch_all() -> list[dict]:
    load_dotenv(ROOT / ".env")
    key = os.environ.get("YOUTUBE_API_KEY")
    if not key:
        raise SystemExit("YOUTUBE_API_KEY が未設定です（.env に追記してください）")

    ch = api("channels", key, part="contentDetails", id=CHANNEL_ID)
    uploads = ch["items"][0]["contentDetails"]["relatedPlaylists"]["uploads"]

    video_ids, token = [], None
    while True:
        r = api("playlistItems", key, part="contentDetails", playlistId=uploads,
                maxResults=50, **({"pageToken": token} if token else {}))
        video_ids += [it["contentDetails"]["videoId"] for it in r["items"]]
        token = r.get("nextPageToken")
        if not token:
            break

    rows = []
    for i in range(0, len(video_ids), 50):
        r = api("videos", key, part="snippet,statistics", id=",".join(video_ids[i:i + 50]))
        for it in r["items"]:
            s = it["statistics"]
            rows.append({
                "views": int(s.get("viewCount", 0)),
                "title": it["snippet"]["title"],
                "date": it["snippet"]["publishedAt"][:10],
            })
    rows.sort(key=lambda x: -x["views"])
    return rows


def write_outputs(rows: list[dict]) -> None:
    (ROOT / "output").mkdir(exist_ok=True)
    (ROOT / "output" / "youtube_stats.json").write_text(
        json.dumps(rows, ensure_ascii=False, indent=1), encoding="utf-8")

    lines = [
        "# 伸びている動画タイトル（YouTube Data APIの実データ・再生回数順トップ30）",
        "# これを分析して\"クリックされる型\"を抽出する。自動更新: python scripts/refresh_youtube.py",
        "# 末尾の（）内は実際の再生回数＝どれだけ当たったかの強さの目安。",
        "",
    ]
    for d in rows[:TOP_N]:
        v = d["views"]
        vs = f"{v / 10000:.0f}万回" if v >= 10000 else f"{v:,}回"
        lines.append(f"{d['title']}（{vs}）")
    (ROOT / "config" / "reference_videos.txt").write_text(
        "\n".join(lines) + "\n", encoding="utf-8")


if __name__ == "__main__":
    rows = fetch_all()
    write_outputs(rows)
    print(f"更新完了：{len(rows)}本を取得、トップ{TOP_N}を reference_videos.txt に反映しました。")
