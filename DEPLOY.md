# デプロイ手順（先生がURLで使えるようにする）

Streamlit Community Cloud（無料）に載せます。**APIキーとパスワードはGitHubには入れず、Cloud側の Secrets に保存**します。

所要：15〜20分。GitHubアカウントが必要（取得済み）。

---

## 全体像

```
ローカルのコード ──push──> GitHub（非公開リポジトリ）──連携──> Streamlit Cloud（URL発行）
                                                              └ Secrets に APIキー / パスワード
```

---

## STEP 1　GitHub に置く

すでに `git` の初期コミットは作成済みです（`~/youtube-script-ai`）。あとは GitHub にアップするだけ。

いちばん簡単なのは **GitHub CLI（gh）** か **GitHub Desktop**。どちらか。

### A) GitHub CLI を使う場合（ターミナル）
```bash
cd ~/youtube-script-ai
gh auth login          # ブラウザでGitHubにログイン（初回のみ）
gh repo create youtube-script-ai --private --source=. --push
```
→ 非公開リポジトリが作られ、コードがアップされます。

### B) GitHub Desktop を使う場合
1. GitHub Desktop を開く →「Add → Add Existing Repository」→ `~/youtube-script-ai` を選択
2. 「Publish repository」→ **Keep this code private にチェック** → Publish

> ✅ `.env`（APIキー）と `output/` は `.gitignore` で除外済み。アップされません。

---

## STEP 2　Streamlit Cloud にデプロイ

1. https://share.streamlit.io を開く →「Continue with GitHub」でログイン
2. 「**Create app**」→「Deploy a public app from GitHub」… ではなく **自分のリポジトリ**を選ぶ
   - Repository: `あなたのID/youtube-script-ai`
   - Branch: `main`
   - **Main file path: `app.py`**
3. 「**Advanced settings**」を開く
   - **Python version: 3.12**（3.14はまだ非対応のため）
   - **Secrets** に次を貼り付け（ここがキーの保管場所）：
     ```toml
     ANTHROPIC_API_KEY = "sk-ant-あなたの本物のキー"
     APP_PASSWORD = "先生に伝えるパスワード"
     ```
4. 「Deploy」を押す → 数分でビルド → `https://〇〇.streamlit.app` のURLが出ます

---

## STEP 3　先生に渡す

- **URL** と **パスワード** の2つだけ伝える。
- 先生はURLを開く → パスワード入力 → すぐ使えます（ターミナル不要）。
- APIキーは先生には見えません（Cloud内に隠れています）。

---

## 運用メモ

- **非公開＆パスワード必須**にしてあるので、URLが漏れても勝手に使われにくいです。万一に備え、パスワードは時々変更を（Cloudの Settings → Secrets で `APP_PASSWORD` を変えるだけ）。
- 参照リストや設定（`config/`）を更新したら、GitHubに push すると自動で反映されます（GitHub Desktop なら Commit → Push）。
- **利用額バロメーター（％）はアプリ再起動でリセットされる**ことがあります（クラウドの一時ファイルのため）。正確な請求は Anthropic Console の Usage / Billing で確認してください。月次で確実に記録したい場合は、外部保存（スプレッドシート等）への連携を後から追加できます。
- APIキーの差し替え・パスワード変更は、いずれも Streamlit Cloud の **Settings → Secrets** で行います（コード変更・再pushは不要）。
