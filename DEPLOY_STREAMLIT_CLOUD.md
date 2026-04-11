# Streamlit Community Cloud 手順

## 1. GitHub へアップロード

このフォルダを GitHub リポジトリへ push します。

秘密情報は GitHub へ上げません。

- `.env`
- `credentials.json`
- `.streamlit/secrets.toml`

## 2. Streamlit Community Cloud で 2 アプリ作成

同じ GitHub リポジトリから 2 本作成します。

### カジコン用

- Main file path: `app_expert.py`
- App URL: 任意

### 会員用

- Main file path: `app_member.py`
- App URL: 任意

## 3. secrets に設定

Cloud 側の secrets に以下を設定します。

```toml
GEMINI_API_KEY = "your_gemini_api_key"
GOOGLE_SHEETS_ID = "your_google_sheets_id"
GOOGLE_DRIVE_FOLDER_ID = "your_google_drive_folder_id"
GOOGLE_EXPERT_SHEET_NAME = "カジコン用"
GOOGLE_MEMBER_SHEET_NAME = "会員用"
GOOGLE_SERVICE_ACCOUNT_JSON = """{"type":"service_account", ... }"""
```

## 4. 動作確認

- カジコン用で診断できる
- 会員用で診断できる
- 会員用で診断士コメント確定ボタンが出ない
- カジコン用で `診断士コメントを保存して確定` が動く
- 保存先タブが `カジコン用` / `会員用` で分かれる

## 5. Drive 設定

将来的に画像列へ Drive 画像を反映します。

必要な設定:

- サービスアカウントに対象 Drive フォルダへの権限付与
- `GOOGLE_DRIVE_FOLDER_ID` の設定

共有ドライブ未設定の間は、画像列の完成形は未設定運用です。
