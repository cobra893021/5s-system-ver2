# 5S System Ver.2

5S 診断システムです。`カジコン用` と `会員用` の 2 系統で運用します。

## ローカル起動

```bash
cd /Users/y/Documents/Playground/5S-System_Ver.2
source .venv/bin/activate
streamlit run app_expert.py --server.port 8501
```

```bash
cd /Users/y/Documents/Playground/5S-System_Ver.2
source .venv/bin/activate
streamlit run app_member.py --server.port 8502
```

- カジコン用: `http://localhost:8501`
- 会員用: `http://localhost:8502`

## モードごとの役割

- `app_expert.py`
  - カジコン用
  - 診断士コメントの保存と確定ができる
  - 学習対象データを作る
- `app_member.py`
  - 会員用
  - 診断と PDF 出力のみ
  - 学習用の保存はしない

## スプレッドシート

同一スプレッドシート内の 2 タブで運用します。

- `カジコン用`
- `会員用`

学習参照は `カジコン用` タブの `ステータス = 確定` の行のみです。

## secrets / 環境変数

ローカルでは `.env` を使います。
本番では Streamlit Community Cloud の secrets を使います。

必要なキー:

- `GEMINI_API_KEY`
- `GOOGLE_SHEETS_ID`
- `GOOGLE_DRIVE_FOLDER_ID`
- `GOOGLE_EXPERT_SHEET_NAME=カジコン用`
- `GOOGLE_MEMBER_SHEET_NAME=会員用`
- `GOOGLE_SERVICE_ACCOUNT_JSON`

`credentials.json`、`.env`、`.streamlit/secrets.toml` は GitHub に上げません。

## 補足

- `record_id` は内部管理用です
- `確定` 後でも再保存で上書きできます
- Drive 保存設定は診断士環境で実施予定です
