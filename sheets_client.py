"""Google Sheets連携（診断結果の保存・確定事例の取得）"""
from __future__ import annotations

import base64
import json
import os
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import gspread
import streamlit as st
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
import io

_env_dir = Path(__file__).resolve().parent
load_dotenv(_env_dir / ".env")
load_dotenv()

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

HEADERS = [
    "診断日時", "会社名", "作業場カテゴリ", "場所カテゴリ",
    "record_id", "画像名", "画像", "AI改善アクション", "AI総合スコア", "AI総評",
    "ステータス", "診断士コメント",
]

REQUIRED_SAVE_HEADERS = [
    "record_id",
    "画像名",
    "AI改善アクション",
    "AI総合スコア",
    "AI総評",
    "ステータス",
]


def _get_secret(name: str, default: str = "") -> str:
    try:
        if name in st.secrets:
            value = st.secrets.get(name)
            return "" if value is None else str(value)
    except Exception:
        pass
    return os.getenv(name, default)


def _get_credentials():
    creds_b64 = _get_secret("GOOGLE_SERVICE_ACCOUNT_JSON_BASE64", "")
    if creds_b64:
        info = json.loads(base64.b64decode(creds_b64).decode("utf-8"))
        return Credentials.from_service_account_info(info, scopes=SCOPES)

    creds_json = None
    try:
        creds_json = st.secrets.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    except Exception:
        creds_json = None
    if creds_json:
        if isinstance(creds_json, str):
            info = json.loads(creds_json)
        else:
            info = dict(creds_json)
        return Credentials.from_service_account_info(info, scopes=SCOPES)

    creds_path = _get_secret("GOOGLE_CREDENTIALS_PATH", "credentials.json")
    return Credentials.from_service_account_file(
        str(Path(__file__).parent / creds_path),
        scopes=SCOPES
    )


def _upload_to_drive(image_bytes: bytes, filename: str) -> str:
    """画像をGoogle Driveにアップロードして IMAGE関数用URLを返す"""
    try:
        creds = _get_credentials()
        service = build("drive", "v3", credentials=creds)

        folder_id = _get_secret("GOOGLE_DRIVE_FOLDER_ID", "")
        if not folder_id:
            print("[upload_to_drive] GOOGLE_DRIVE_FOLDER_ID が設定されていません", flush=True)
            return ""

        file_metadata = {
            "name": filename,
            "parents": [folder_id]
        }
        media = MediaIoBaseUpload(
            io.BytesIO(image_bytes),
            mimetype="image/jpeg",
            resumable=False
        )
        file = service.files().create(
            body=file_metadata,
            media_body=media,
            fields="id"
        ).execute()
        file_id = file["id"]

        service.permissions().create(
            fileId=file_id,
            body={"type": "anyone", "role": "reader"}
        ).execute()

        image_url = f"https://drive.google.com/uc?id={file_id}"
        print(f"[upload_to_drive] OK url={image_url}", flush=True)
        return image_url

    except Exception as e:
        print(f"[upload_to_drive] error: {e}", flush=True)
        return ""


def _sheet_name_for_mode(mode: str) -> str:
    if mode == "member":
        return _get_secret("GOOGLE_MEMBER_SHEET_NAME", "会員用")
    return _get_secret("GOOGLE_EXPERT_SHEET_NAME", "カジコン用")


def _get_sheet(mode: str = "expert"):
    sheet_id = _get_secret("GOOGLE_SHEETS_ID", "")
    if not sheet_id:
        raise ValueError("GOOGLE_SHEETS_ID が設定されていません。")
    client = gspread.authorize(_get_credentials())
    return client.open_by_key(sheet_id).worksheet(_sheet_name_for_mode(mode))


def _row_from_headers(headers: list[str], values: dict[str, Any]) -> list[Any]:
    """シートのヘッダー順に合わせて1行データを組み立てる。"""
    return [values.get(header, "") for header in headers]


def _ensure_required_headers(headers: list[str], required_headers: list[str]) -> None:
    missing = [header for header in required_headers if header not in headers]
    if missing:
        raise ValueError(
            "必要なヘッダーが不足しています。"
            f" 不足列: {', '.join(missing)}"
        )


def save_to_sheets(
    result: dict[str, Any],
    location: str,
    filename: str,
    record_id: str,
    mode: str = "expert",
    company: str = "",
    image_bytes: bytes = b""
) -> dict[str, str]:
    """診断結果をGoogle Sheetsに1行追加する"""
    print(f"[save_to_sheets] start filename={filename!r}", flush=True)
    try:
        sheet = _get_sheet(mode)
        headers = sheet.row_values(1) or HEADERS
        _ensure_required_headers(headers, REQUIRED_SAVE_HEADERS)
        scene = str(result.get("scene_category") or "other")
        summary = str(result.get("summary") or "")
        score = result.get("overall_score", 0)
        actions = result.get("action_items") or []
        actions_json = json.dumps(actions, ensure_ascii=False)
        if len(actions_json) > 1900:
            actions_json = actions_json[:1900] + "…"

        now = datetime.now(timezone.utc).strftime("%Y/%m/%d")

        image_formula = ""
        if image_bytes:
            image_url = _upload_to_drive(image_bytes, filename)
            if image_url:
                image_formula = f'=IMAGE("{image_url}")'

        row = _row_from_headers(
            headers,
            {
                "診断日時": now,
                "会社名": company.strip() or "未入力",
                "作業場カテゴリ": location.strip() or "",
                "場所カテゴリ": scene,
                "record_id": record_id,
                "画像名": filename,
                "画像": image_formula,
                "AI改善アクション": actions_json,
                "AI総合スコア": score,
                "AI総評": summary,
                "ステータス": "AI診断済み",
                "診断士コメント": "",
            },
        )
        sheet.append_row(row, value_input_option="USER_ENTERED")
        saved_record_id = ""
        saved_row_number = ""
        rows = sheet.get_all_values()
        if rows:
            headers = rows[0]
            try:
                record_id_col = headers.index("record_id")
                for row_idx in range(len(rows), 1, -1):
                    saved_row = rows[row_idx - 1]
                    if len(saved_row) > record_id_col and saved_row[record_id_col] == record_id:
                        saved_record_id = record_id
                        saved_row_number = str(row_idx)
                        break
            except ValueError:
                saved_record_id = ""
        print(f"[save_to_sheets] end OK filename={filename!r}", flush=True)
        return {
            "sheet_name": sheet.title,
            "record_id": saved_record_id or record_id,
            "row_number": saved_row_number,
        }
    except Exception as e:
        print(f"[save_to_sheets] error: {e}", flush=True)
        raise RuntimeError(
            f"Google Sheets 保存に失敗しました。"
            f" mode={mode}, sheet={_sheet_name_for_mode(mode)}, filename={filename}, error={e}"
        ) from e


def get_confirmed_cases() -> list[dict[str, Any]]:
    """ステータスが「確定」の行を取得する"""
    try:
        sheet = _get_sheet("expert")
        rows = sheet.get_all_records()
        out = []
        for row in rows:
            if row.get("ステータス") == "確定":
                out.append({
                    "場所カテゴリ": str(row.get("場所カテゴリ") or ""),
                    "AI総合スコア": int(row.get("AI総合スコア") or 0),
                    "診断士コメント": str(row.get("診断士コメント") or ""),
                })
        return out
    except Exception as e:
        print(f"[get_confirmed_cases] error: {e}", flush=True)
        return []


def update_expert_review(record_id: str, expert_comment: str, status: str | None = None) -> None:
    """record_id をキーに診断士コメントと任意でステータスを更新する。"""
    sheet = _get_sheet("expert")
    rows = sheet.get_all_values()
    if not rows:
        raise ValueError("スプレッドシートにヘッダー行がありません。")

    headers = rows[0]
    try:
        record_id_col = headers.index("record_id") + 1
        comment_col = headers.index("診断士コメント") + 1
    except ValueError as err:
        raise ValueError("必要な列（record_id / 診断士コメント）が見つかりません。") from err

    status_col: int | None = None
    if status is not None:
        try:
            status_col = headers.index("ステータス") + 1
        except ValueError as err:
            raise ValueError("必要な列（ステータス）が見つかりません。") from err

    for row_idx, row in enumerate(rows[1:], start=2):
        if len(row) >= record_id_col and row[record_id_col - 1] == record_id:
            sheet.update_cell(row_idx, comment_col, expert_comment)
            if status_col is not None:
                sheet.update_cell(row_idx, status_col, status)
            return

    raise ValueError("対象の record_id に対応する行が見つかりません。")
