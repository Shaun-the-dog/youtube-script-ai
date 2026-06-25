# YouTube 台本 自動生成ツール（杉山先生・借金問題チャンネル）

過去台本(Excel)から抽出した「型」に沿って、台本を **杉山先生の様式の Excel**
（項目｜内容｜参考記事｜文字数、1行=1ナレーション、4,500字超）で生成します。

- **毎回固定の定型文はそのまま挿入**（自己紹介・3つのLINE誘導・ED など。`config/channel.yaml` の `fixed_blocks`）。
- **テーマと本編だけ AI が生成**。完成系台本に限りなく近い下書きを出力。
- 先生の価値観を反映：**借金は必ずしも悪ではない／債務者を貶めない／制度紹介で終わらせずメッセージ性を持たせる／過去動画と丸かぶりさせない**。
- 生成後に **コンプラ点検**（NGワード機械検知 ＋ AI校閲）を自動実行。

出来上がった Excel を先生に渡し、
①弁護士として適切な表現か ②情報の誤りがないか ③先生の言い回しか
を確認・修正 → 撮影、という流れを想定しています。

## セットアップ

```bash
cd ~/youtube-script-ai
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env   # .env を開いて ANTHROPIC_API_KEY を貼る
```

## 使い方A：画面（先生が直接触る・おすすめ）

```bash
streamlit run streamlit_app.py
```

ブラウザが開き、「企画を3つ生成 → 1つ選ぶ → 台本(Excel)を生成・ダウンロード」を
ボタン操作だけで行えます。コンプラ点検結果と台本プレビューも画面に出ます。

## 使い方B：コマンドライン

```bash
# テーマを指定して台本(Excel)を生成（生成後に自動でコンプラ点検）
python src/generate.py script --theme "任意整理をすると家族にバレるのか"

# 企画案を6本出す（メッセージ性・社会的切り口つき）
python src/generate.py ideate --count 6

# 企画→番号で選択→台本(Excel)→点検 を一気通し（おすすめ）
python src/generate.py run --count 6

# 既存の台本ファイル(xlsx/txt)をコンプラ点検
python src/generate.py check --file output/2026-06-25_xxx_AI下書き.xlsx
```

生成物は `output/` に `YYYY-MM-DD_タイトル_AI下書き.xlsx` で保存されます。

## 先生に渡す前提（重要）

- ファイル名・作成者欄に「AI下書き（要・杉山先生確認）」と明記されます。**そのまま撮影用ではなく、先生の確認・修正が前提**です。
- **参考記事は空欄**です。事実の裏取りは AI に任せず、本編の各まとまりに **実在する記事を2記事以上、人手で記載**してください（冒頭とED以外）。
- 中盤のLINE誘導は本編のちょうど真ん中に自動挿入されます。区切りが不自然なら先生・担当者が移動してください。

## 運用のコツ

- **企画の質は `config/reference_videos.txt` で決まる**。YouTube Studio のアナリティクスで再生回数の多い順に自分の動画タイトルを貼る／伸びている競合のタイトルも貼ると、AIがその「クリックされる型」を分析して企画を作る。実データに差し替えるほど精度が上がる。
- 動画を出すたびに、採用テーマを `config/topics_history.txt` に1行追記 → 次回以降の重複（焼き直し）を回避。
- 先生の実際の言い回しに寄せたいときは、`config/channel.yaml` の `fixed_blocks`（定型文）や `voice` を更新。
- 言い回しの好み・NG表現が出てきたら `compliance.ng_phrases` / `philosophy.forbidden_expressions` に追記。

## 構成

```
config/channel.yaml        … 型・定型文・先生の価値観・コンプラ基準（ここを編集する）
config/topics_history.txt  … 過去テーマ（重複回避）
src/generate.py            … 本体（ideate / script / check / run）
output/                    … 生成された台本(.xlsx)の保存先
```

使用モデル: `claude-opus-4-8`（Anthropic Claude）

## 補足：台本の型（過去台本から抽出）

```
最初の引き＆自己紹介 → テーマ明言 → 視聴者の気持ちを代弁
→ 認識のギャップを突いて視聴意欲掻き立て → 視聴人物像の明確化
→（LINE誘導）→ 本編 → LINE誘導(中盤) → 本編(続き)
→ まとめ → ED
```

この並びと、各所の定型文は `config/channel.yaml` に定義されています。
