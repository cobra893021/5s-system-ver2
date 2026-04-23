from __future__ import annotations
import base64
import hashlib
import html
import io
import os
import uuid
from datetime import datetime
from typing import Any, Optional

import google.generativeai as genai
import streamlit as st
from PIL import Image
from dotenv import load_dotenv
from knowledge import get_knowledge_context
from pdf_report import generate_pdf, generate_zip


def pil_image_to_b64_jpeg(
    pil_img: Image.Image,
    size: tuple[int, int] = (400, 300),
    quality: int = 82,
) -> str:
    buf = io.BytesIO()
    t = pil_img.copy()
    t.thumbnail(size, Image.LANCZOS)
    t.save(buf, format="JPEG", quality=quality)
    return base64.b64encode(buf.getvalue()).decode()


def pil_image_to_jpeg_bytes(pil_img: Image.Image, quality: int = 88) -> bytes:
    """PDF・Sheets・Drive保存で共通利用するJPEGバイト列を作る。"""
    buf = io.BytesIO()
    pil_img.convert("RGB").save(buf, format="JPEG", quality=quality)
    return buf.getvalue()


def _gallery_item_key(item: dict[str, Any]) -> tuple[str, str]:
    """同名・同サイズの別ファイルを区別するため (名前, digest) でキー化する。"""
    if item.get("digest"):
        return (item["name"], item["digest"])
    return (item["name"], hashlib.md5(item["data"]).hexdigest())


def _upload_digest_for_file(f) -> tuple[str, bytes]:
    """Streamlit の file_id があればそれを、なければ内容の MD5 を使う。"""
    data = f.getvalue()
    fid = getattr(f, "file_id", None)
    if fid is not None:
        return (f"fid:{fid}", data)
    return (hashlib.md5(data).hexdigest(), data)


def _clear_gallery_and_diagnosis_state() -> None:
    """ギャラリー・アップローダー・診断結果をまとめて診断前の状態に戻す。"""
    st.session_state["gallery_images"] = []
    st.session_state.pop("_gallery_upload_sig", None)
    st.session_state["gallery_uploader_key"] = int(st.session_state.get("gallery_uploader_key", 0)) + 1
    st.session_state.pop("results", None)
    st.session_state.pop("selected_idx", None)
    st.query_params.pop("diag_sel", None)


def _logout_member_user() -> None:
    st.session_state.pop("member_auth", None)
    st.session_state.pop("main_company", None)
    st.session_state.pop("main_location", None)
    _clear_gallery_and_diagnosis_state()


def is_member_login_disabled() -> bool:
    """確認作業中だけ会員ログイン画面を迂回するための切り替え。"""
    value = get_runtime_secret("DISABLE_MEMBER_LOGIN", "")
    return str(value).strip().lower() in {"1", "true", "yes", "on"}


def render_member_login() -> None:
    st.markdown(
        """
        <div style="max-width:520px;margin:2.5rem auto 1.5rem;padding:2rem;background:#ffffff;
                    border:1px solid #e2e8f0;border-radius:18px;box-shadow:0 8px 32px rgba(52,109,153,0.08);">
          <div style="font-size:1.35rem;font-weight:700;color:#346D99;margin-bottom:0.4rem;">会員ログイン</div>
          <div style="color:#475569;font-size:0.95rem;line-height:1.7;">
            発行されたログインIDとパスワードを入力してください。
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    with st.form("member_login_form", clear_on_submit=False):
        login_id = st.text_input("ログインID", key="member_login_id")
        password = st.text_input("パスワード", type="password", key="member_login_password")
        submitted = st.form_submit_button("ログイン", use_container_width=True)

    if submitted:
        try:
            from sheets_client import authenticate_member_user

            user = authenticate_member_user(login_id, password)
            if not user:
                st.error("ログインIDまたはパスワードが正しくありません。")
                return

            st.session_state["member_auth"] = user
            st.session_state["main_company"] = user.get("company_name", "")
            st.session_state["main_location"] = user.get("default_location", "")
            st.rerun()
        except Exception as e:
            st.error(f"ログインに失敗しました: {e}")


# ─── カスタム CSS ────────────────────────────────────────────────────────────────
GLOBAL_CSS = """
<style>
  @import url('https://fonts.googleapis.com/css2?family=Shippori+Mincho+B1:wght@400;600;700&family=Noto+Sans+JP:wght@300;400;500;700&family=Inter:wght@300;400;600;700&display=swap');

  :root {
    --primary: #346D99;
    --primary-dark: #245E86;
    --primary-hover: #265680;
    --primary-active: #1e4a6e;
    --primary-light: #EEF5FB;
    --primary-light-hover: #DCEAF6;
    --dark-btn: #161922;
    --dark-btn-hover: #242938;
    --dark-btn-active: #0f1218;
    --bg: #f1f5f9;
    --card: #ffffff;
    --text: #1e293b;
    --text-sub: #475569;
    --text-muted: #94a3b8;
    --border: #e2e8f0;
    --shadow: 0 2px 16px rgba(52,109,153,0.08);
    --shadow-md: 0 4px 24px rgba(52,109,153,0.12);
    --radius: 14px;
    --radius-lg: 20px;
  }

  html, body, [class*="css"] {
    font-family: 'Noto Sans JP', 'Inter', sans-serif;
    color: var(--text);
  }

  /* ─── 背景 ─── */
  .stApp {
    background-color: var(--bg);
    background-image:
      linear-gradient(rgba(52,109,153,0.045) 1px, transparent 1px),
      linear-gradient(90deg, rgba(52,109,153,0.045) 1px, transparent 1px);
    background-size: 32px 32px;
    min-height: 100vh;
  }

  /* ─── サイドバー ─── */
  section[data-testid="stSidebar"] {
    background: #ffffff;
    border-right: 1px solid var(--border);
  }
  section[data-testid="stSidebar"] * { color: var(--text) !important; }
  section[data-testid="stSidebar"] label { color: var(--text-sub) !important; }

  /* ─── ヘッダー ─── */
  .hero-title {
    font-family: 'Shippori Mincho B1', 'Noto Serif JP', serif;
    font-size: 2.4rem;
    font-weight: 700;
    color: var(--primary);
    text-align: center;
    margin-bottom: 0.4rem;
    letter-spacing: 0.01em;
  }

  .report-download-heading,
  .score-detail-heading {
    color: #346D99;
    font-weight: 800;
    margin-bottom: 1rem;
  }

  .score-detail-heading {
    font-size: 1.3rem;
    margin-top: 2rem;
  }

  .report-download-heading {
    font-size: 1.3rem;
    margin-top: 1rem;
  }

  .anchor-link,
  a.anchor-link,
  a[class*="anchor"],
  [data-testid="stMarkdownContainer"] a[href^="#"],
  [data-testid="stMarkdownContainer"] h1 a,
  [data-testid="stMarkdownContainer"] h2 a,
  [data-testid="stMarkdownContainer"] h3 a,
  [data-testid="stMarkdownContainer"] h4 a {
    display: none !important;
    visibility: hidden !important;
    width: 0 !important;
    height: 0 !important;
    margin: 0 !important;
    padding: 0 !important;
    pointer-events: none !important;
  }
  .hero-subtitle {
    text-align: center;
    color: var(--text-sub);
    font-size: 0.98rem;
    font-weight: 400;
    margin-bottom: 2rem;
    letter-spacing: 0.01em;
  }
  .mode-badge-wrap {
    display: flex;
    justify-content: flex-end;
    margin-bottom: 0.4rem;
  }
  .mode-badge {
    display: inline-flex;
    align-items: center;
    gap: 0.35rem;
    padding: 0.24rem 0.7rem;
    border-radius: 999px;
    font-size: 0.74rem;
    font-weight: 700;
    letter-spacing: 0.04em;
    border: 1px solid rgba(52,109,153,0.16);
    background: rgba(255,255,255,0.88);
    color: var(--primary-dark);
    box-shadow: 0 8px 18px rgba(52,109,153,0.08);
  }

  /* ─── サムネイルグリッド ─── */
  .thumb-grid { margin-bottom: 1rem; }
  .thumb-row {
    display: grid;
    grid-template-columns: repeat(5, 1fr);
    gap: 8px;
    margin-bottom: 8px;
  }
  .thumb-cell {
    border: 2px dashed #cbd5e1;
    border-radius: 10px;
    overflow: hidden;
    aspect-ratio: 4/3;
    position: relative;
    background: #f8fafc;
    transition: border-color 0.2s, box-shadow 0.2s;
  }
  .thumb-cell.filled {
    border: 2px solid var(--primary);
    box-shadow: 0 2px 8px rgba(52,109,153,0.12);
  }
  .thumb-cell.filled img {
    width: 100%; height: 100%;
    object-fit: cover; display: block;
  }
  .thumb-cell.empty {
    display: flex; align-items: center; justify-content: center;
    color: #cbd5e1; font-size: 1.6rem;
  }
  .thumb-label {
    position: absolute; bottom: 0; left: 0; right: 0;
    background: rgba(30,41,59,0.55);
    color: #fff; font-size: 0.68rem;
    padding: 3px 6px;
    white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
  }

  /* ─── カード ─── */
  .card {
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: var(--radius-lg);
    padding: 1.6rem;
    box-shadow: var(--shadow);
    margin-bottom: 1rem;
    transition: box-shadow 0.2s ease;
  }
  .card:hover { box-shadow: var(--shadow-md); }

  /* ─── バッジ ─── */
  .badge-row {
    display: flex; gap: 0.5rem; flex-wrap: wrap;
    margin-bottom: 1.8rem; justify-content: center;
  }
  .s-badge {
    padding: 0.3rem 1rem;
    border-radius: 999px;
    font-size: 0.78rem;
    font-weight: 600;
    letter-spacing: 0.04em;
    background: var(--primary-light);
    color: var(--primary-dark);
    border: 1px solid rgba(52,109,153,0.2);
  }

  /* ─── スコア ─── */
  .score-value {
    font-size: 3.6rem;
    font-weight: 700;
    color: var(--primary);
    line-height: 1;
  }

  /* ─── セクション見出し ─── */
  .section-heading {
    color: var(--text);
    font-size: 1rem;
    font-weight: 600;
    margin-bottom: 0.6rem;
  }

  /* ─── 診断結果テキスト ─── */
  .result-box {
    background: var(--primary-light);
    border-left: 3px solid var(--primary);
    border-radius: 0 var(--radius) var(--radius) 0;
    padding: 1rem 1.2rem;
    color: var(--text);
    font-size: 0.92rem;
    line-height: 1.8;
    margin-bottom: 0.6rem;
    white-space: pre-wrap;
  }

  /* ─── プログレスバー ─── */
  .stProgress > div > div > div > div {
    background: var(--primary) !important;
    border-radius: 999px !important;
  }
  .stProgress > div > div > div {
    background: var(--border) !important;
    border-radius: 999px !important;
  }

  /* 一括診断の進捗メッセージ（st.progress の text ではなく別要素で表示） */
  p.diagnose-progress-msg {
    color: #000000 !important;
    font-size: 0.92rem;
    font-weight: 500;
    margin: 0 0 0.45rem 0;
    line-height: 1.45;
  }

  /* st.image 周りのフルスクリーン／拡大 UI を非表示（診断写真は HTML img に差し替え済みの想定） */
  [data-testid="stImage"] button,
  [data-testid="stImage"] [role="button"] {
    display: none !important;
  }

  /* ─── ボタン ─── */
  .stButton > button {
    background: var(--primary) !important;
    color: white !important;
    border: none !important;
    border-radius: var(--radius);
    padding: 0.65rem 1.8rem;
    font-size: 0.95rem;
    font-weight: 600;
    width: 100%;
    letter-spacing: 0.04em;
    box-shadow: 0 4px 14px rgba(52,109,153,0.25);
    transition: background 0.18s, transform 0.15s, box-shadow 0.18s;
  }

  .stButton > button[kind="primary"],
  .stButton button[kind="primary"] {
    background: var(--primary) !important;
    color: white !important;
    border: none !important;
    box-shadow: 0 4px 14px rgba(52,109,153,0.25) !important;
  }

  .stButton > button:active {
    filter: brightness(0.70) !important;
    opacity: 1 !important;
  }

  .stButton > button:focus {
    filter: brightness(0.80) !important;
    opacity: 1 !important;
    box-shadow: none !important;
  }

  .stDownloadButton > button:active {
    filter: brightness(0.70) !important;
    opacity: 1 !important;
  }

  /* ─── ファイルアップロード ─── */
  .stFileUploader > div {
    border: 2px dashed rgba(52,109,153,0.3) !important;
    border-radius: var(--radius-lg) !important;
    background: white !important;
    transition: border-color 0.2s;
  }
  .stFileUploader > div:hover {
    border-color: var(--primary) !important;
  }

  /* ─── テキスト入力 ─── */
  .stTextInput > div > div > input {
    background: white !important;
    border: 1px solid var(--border) !important;
    border-radius: 10px !important;
    color: var(--text) !important;
  }
  .stTextInput > div > div > input:focus {
    border-color: var(--primary) !important;
    box-shadow: 0 0 0 3px rgba(52,109,153,0.12) !important;
  }
  .stTextInput > div > div > input::placeholder {
    color: #94a3b8 !important;
    opacity: 1 !important;
  }
  /* サイドバー全体の * { color } より優先してプレースホルダーをグレーに */
  section[data-testid="stSidebar"] .stTextInput input::placeholder {
    color: #94a3b8 !important;
    opacity: 1 !important;
  }

  /* ─── セレクトボックス ─── */
  .stSelectbox > div > div {
    background: white !important;
    border: 1px solid var(--border) !important;
    border-radius: 10px !important;
    color: var(--text) !important;
  }

  /* ─── ラジオ ─── */
  .stRadio label { color: var(--text-sub) !important; }

  /* ─── 区切り線 ─── */
  hr { border-color: var(--border) !important; }

  /* ─── タブ ─── */
  button[data-baseweb="tab"] {
    color: var(--text-sub) !important;
    font-weight: 500;
  }
  button[data-baseweb="tab"][aria-selected="true"] {
    color: var(--primary) !important;
    border-bottom-color: var(--primary) !important;
  }

  /* ─── エクスパンダー（5S詳細項目用） ─── */
  div[data-testid="stExpander"],
  div[data-testid="stExpander"] > div,
  div[data-testid="stExpander"] details {
    border: none !important;
    box-shadow: none !important;
    background: transparent !important;
    outline: none !important;
  }

  details {
    border: none !important;
    box-shadow: none !important;
    background: transparent !important;
    outline: none !important;
  }

  details > summary {
    background-color: white !important;
    border: 1px solid var(--border) !important;
    border-radius: var(--radius) !important;
    padding: 0.8rem 1.2rem !important;
    color: var(--text) !important;
    font-weight: 600 !important;
    margin-bottom: 0.5rem !important;
    transition: all 0.2s ease !important;
    box-shadow: 0 1px 3px rgba(0,0,0,0.02);
  }
  details > summary:hover {
    border-color: var(--primary) !important;
    background-color: var(--primary-light-hover) !important;
  }
  details[open] > summary {
    border-bottom: none !important;
    border-radius: var(--radius) var(--radius) 0 0 !important;
    margin-bottom: 0 !important;
  }
  /* エクスパンダーの中身のパディング調整 */
  details > div {
    border: 1px solid var(--border) !important;
    border-top: none !important;
    border-radius: 0 0 var(--radius) var(--radius) !important;
    padding: 1rem !important;
    background-color: white !important;
    margin-bottom: 1rem !important;
  }

  /* アクションカード */
  .action-card {
    background: white;
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 0.9rem 1.2rem;
    margin-bottom: 0.5rem;
    display: flex;
    align-items: flex-start;
    gap: 0.8rem;
    box-shadow: var(--shadow);
  }
  .action-num {
    min-width: 24px; height: 24px;
    background: var(--primary); color: white;
    border-radius: 50%;
    font-size: 0.75rem; font-weight: 700;
    display: flex; align-items: center; justify-content: center;
  }
  .action-num.mid { background: #64748b; }
  .action-num.low { background: #94a3b8; }

  /* ─── サムネイル番号 ─── */
  .thumb-num {
    position: absolute; top: 4px; left: 6px;
    background: rgba(52,109,153,0.82);
    color: white; font-size: 0.6rem; font-weight: 700;
    padding: 1px 5px; border-radius: 999px; line-height: 1.6;
  }

  /* ─── 左パネル ─── */
  .list-panel-title {
    font-size: 0.75rem; font-weight: 700; color: var(--text-sub);
    text-transform: uppercase; letter-spacing: 0.09em;
    margin-bottom: 10px;
  }

  /* ─── ファイルアップロード 強調 ─── */
  [data-testid="stFileUploaderDropzone"] {
    border: 2px dashed rgba(52,109,153,0.4) !important;
    border-radius: var(--radius-lg) !important;
    background: #f8fafc !important; /* 薄い背景にして領域をわかりやすく */
    padding: 2.5rem 1.5rem !important; /* 広めの余白 */
    transition: all 0.25s !important;
  }
  /* 中の文字をはっきり見せる */
  [data-testid="stFileUploaderDropzone"] * {
    color: var(--text) !important;
  }
  [data-testid="stFileUploaderDropzone"] div[data-testid="stMarkdownContainer"] p {
    font-size: 1.1rem !important;
    font-weight: 500 !important;
    color: var(--text) !important;
  }
  /* サイズ制限の小文字など */
  [data-testid="stFileUploaderDropzone"] small {
    color: var(--text-sub) !important;
  }
  /* アップロードアイコンの色 */
  [data-testid="stFileUploaderDropzone"] svg {
    color: var(--primary) !important;
    fill: var(--primary) !important;
    width: 2.5rem !important;
    height: 2.5rem !important;
  }
  /* Browse files（ファイル選択ボタン）をわかりやすく */
  [data-testid="stFileUploaderDropzone"] button {
    background: var(--primary) !important;
    color: white !important;
    border: none !important;
    font-weight: 600 !important;
    padding: 0.5rem 1.2rem !important;
    border-radius: 8px !important;
  }
  [data-testid="stFileUploaderDropzone"]:hover,
  [data-testid="stFileUploaderDropzone"]:focus-within {
    border-color: var(--primary) !important;
    background: var(--primary-light-hover) !important;
    box-shadow: 0 0 0 4px rgba(52,109,153,0.1) !important;
  }
  /* 親要素の余計な枠・背景を消して隙間を無くす */
  .stFileUploader > div {
    border: none !important;
    background: transparent !important;
    padding: 0 !important;
    min-height: auto !important;
  }

  /* ─── アップロード済みファイルリスト欄を非表示（サムネイルグリッドで代替）─── */
  [data-testid="stFileUploaderFileList"],
  .stFileUploader small,
  .stFileUploader [data-testid="stFileUploaderFile"] {
    display: none !important;
  }
  /* ファイルアップローダー全体の下マージンを詰める */
  .stFileUploader {
    margin-bottom: 0 !important;
  }

  /* ─── セカンダリボタン ─── */
  .stButton > button[kind="secondary"] {
    background: white !important;
    color: var(--primary) !important;
    border: 1.5px solid var(--primary) !important;
    box-shadow: none !important;
  }

  /* ─── ページリンクボタン（5Sを学ぶ）のスタイル ─── */
  [data-testid="stPageLink-NavLink"] {
    background-color: white !important;
    border: 1.5px solid var(--primary) !important;
    border-radius: 8px !important;
    text-align: center !important;
    display: flex !important;
    justify-content: center !important;
    transition: all 0.2s ease !important;
  }
  [data-testid="stPageLink-NavLink"] p {
    color: var(--primary) !important;
    font-weight: 700 !important;
    margin: 0 !important;
  }
  [data-testid="stPageLink-NavLink"]:hover {
    background-color: var(--primary-light-hover) !important;
    box-shadow: 0 4px 6px rgba(0,0,0,0.05) !important;
  }

  /* ─── メイン診断写真のサイズ制限 ─── */
  [data-testid="stImage"] img {
    max-height: 480px !important;
    width: auto !important;
    object-fit: contain !important;
    margin: 0 auto !important;
    display: block !important;
    border-radius: 8px !important;
  }

  /* 診断結果「ビフォー」— 表示枠サイズを写真ごとに統一（枠内に収めてトリミングしない） */
  .before-diagnosis-photo-frame {
    width: 100%;
    height: 400px;
    background: #f1f5f9;
    border: 1px solid var(--border);
    border-radius: 10px;
    display: flex;
    align-items: center;
    justify-content: center;
    overflow: hidden;
    box-sizing: border-box;
  }
  .before-diagnosis-photo-img {
    max-width: 100%;
    max-height: 100%;
    width: auto;
    height: auto;
    object-fit: contain;
    display: block;
    border-radius: 8px;
  }

  /* ZIPダウンロードボタン */
  .stDownloadButton > button,
  .stDownloadButton button {
    background: var(--dark-btn) !important;
    color: white !important;
    border: 1px solid rgba(255,255,255,0.12) !important;
    box-shadow:
      0 10px 24px rgba(15, 23, 42, 0.14),
      inset 0 1px 0 rgba(255,255,255,0.05) !important;
    transition: filter 0.2s, box-shadow 0.2s, transform 0.15s !important;
  }

  .stButton > button[kind="primary"]:hover,
  .stButton > button[kind="primary"]:focus,
  .stButton button[kind="primary"]:hover,
  .stButton button[kind="primary"]:focus {
    background: var(--primary-hover) !important;
    color: white !important;
    box-shadow:
      0 10px 26px rgba(36, 94, 134, 0.24),
      0 0 0 3px rgba(255,255,255,0.16),
      inset 0 1px 0 rgba(255,255,255,0.08) !important;
    transform: translateY(-1px) !important;
  }

  .stButton > button[kind="primary"]:active,
  .stButton button[kind="primary"]:active {
    background: var(--primary-active) !important;
    color: white !important;
    box-shadow: 0 6px 16px rgba(36, 94, 134, 0.22) !important;
    transform: translateY(0) !important;
  }

  /* 指定ボタンだけを黒系デザインで固定 */
  div[data-testid="stVerticalBlock"]:has(img[alt^="score-btn-anchor"])
    > div.element-container:has(img[alt^="score-btn-anchor"])
    + div.element-container button,
  div[data-testid="stVerticalBlock"]:has(img[alt^="clear-btn-anchor"])
    > div.element-container:has(img[alt^="clear-btn-anchor"])
    + div.element-container button {
    background: var(--dark-btn) !important;
    color: white !important;
    border: 1px solid rgba(255,255,255,0.12) !important;
    box-shadow:
      0 10px 24px rgba(15, 23, 42, 0.14),
      inset 0 1px 0 rgba(255,255,255,0.05) !important;
  }

  div[data-testid="stVerticalBlock"]:has(img[alt^="score-btn-anchor"])
    > div.element-container:has(img[alt^="score-btn-anchor"])
    + div.element-container button:hover,
  div[data-testid="stVerticalBlock"]:has(img[alt^="score-btn-anchor"])
    > div.element-container:has(img[alt^="score-btn-anchor"])
    + div.element-container button:focus,
  div[data-testid="stVerticalBlock"]:has(img[alt^="clear-btn-anchor"])
    > div.element-container:has(img[alt^="clear-btn-anchor"])
    + div.element-container button:hover,
  div[data-testid="stVerticalBlock"]:has(img[alt^="clear-btn-anchor"])
    > div.element-container:has(img[alt^="clear-btn-anchor"])
    + div.element-container button:focus {
    background: var(--dark-btn-hover) !important;
    color: white !important;
    border-color: rgba(255,255,255,0.22) !important;
    box-shadow:
      0 12px 28px rgba(15, 23, 42, 0.18),
      0 0 0 3px rgba(255,255,255,0.18),
      inset 0 1px 0 rgba(255,255,255,0.10) !important;
    transform: translateY(-1px) !important;
  }

  div[data-testid="stVerticalBlock"]:has(img[alt^="score-btn-anchor"])
    > div.element-container:has(img[alt^="score-btn-anchor"])
    + div.element-container button:active,
  div[data-testid="stVerticalBlock"]:has(img[alt^="clear-btn-anchor"])
    > div.element-container:has(img[alt^="clear-btn-anchor"])
    + div.element-container button:active {
    background: var(--dark-btn-active) !important;
    color: white !important;
    transform: translateY(0) !important;
  }

  [data-testid="stPageLink-NavLink"]:focus {
    background: var(--primary-light-hover) !important;
    color: var(--primary-dark) !important;
    border-color: var(--primary-dark) !important;
    box-shadow: 0 4px 10px rgba(52,109,153,0.10) !important;
  }

  div[data-testid="stDownloadButton"] button:hover,
  div[data-testid="stDownloadButton"] button:focus,
  .stDownloadButton button:hover,
  .stDownloadButton button:focus {
    background: var(--dark-btn-hover) !important;
    color: white !important;
    opacity: 1 !important;
    border-color: rgba(255,255,255,0.22) !important;
    box-shadow:
      0 12px 28px rgba(15, 23, 42, 0.16),
      0 0 0 3px rgba(255,255,255,0.16),
      inset 0 1px 0 rgba(255,255,255,0.08) !important;
    transform: translateY(-1px) !important;
  }

  .mobile-file-pill {
    display: none;
  }

  @media (max-width: 768px) {
    .block-container {
      padding-left: 1rem !important;
      padding-right: 1rem !important;
      padding-top: 1.2rem !important;
    }

    .hero-title {
      font-size: clamp(1.34rem, 7vw, 1.72rem);
      line-height: 1.2;
      margin-top: 0.4rem;
      letter-spacing: -0.035em;
      white-space: nowrap;
    }

    .report-download-heading,
    .score-detail-heading {
      white-space: nowrap;
      line-height: 1.25;
      letter-spacing: -0.02em;
    }

    .report-download-heading {
      font-size: clamp(1.05rem, 5.1vw, 1.24rem);
    }

    .score-detail-heading {
      font-size: clamp(1.05rem, 5.2vw, 1.28rem);
    }

    .hero-subtitle {
      font-size: 0.88rem;
      line-height: 1.65;
      margin-bottom: 1.2rem;
    }

    .mode-badge-wrap {
      justify-content: center;
      margin-bottom: 0.7rem;
    }

    .badge-row {
      gap: 0.35rem;
      margin-bottom: 1rem;
    }

    .s-badge {
      padding: 0.24rem 0.7rem;
      font-size: 0.72rem;
    }

    [data-testid="stFileUploaderDropzone"] {
      padding: 1.4rem 1rem !important;
    }

    div[data-testid="stHorizontalBlock"] {
      flex-direction: column !important;
      gap: 0.65rem !important;
    }

    div[data-testid="stHorizontalBlock"] > div {
      width: 100% !important;
      min-width: 100% !important;
    }

    div[data-testid="stExpander"],
    div[data-testid="stExpander"] > div,
    div[data-testid="stExpander"] details,
    details {
      border: none !important;
      box-shadow: none !important;
      background: transparent !important;
      outline: none !important;
    }

    details > summary {
      padding: 0.72rem 0.9rem !important;
      font-size: clamp(0.92rem, 4.1vw, 1rem) !important;
    }

    .gallery-thumb-card,
    .thumb-cell.empty {
      display: none !important;
    }

    .mobile-file-pill {
      display: flex;
      align-items: center;
      gap: 0.55rem;
      width: 100%;
      min-height: 44px;
      padding: 0.72rem 0.9rem;
      margin-bottom: 0.45rem;
      border: 1.5px solid rgba(52,109,153,0.32);
      border-radius: 12px;
      background: rgba(255,255,255,0.92);
      color: var(--primary-dark);
      font-size: 0.86rem;
      font-weight: 700;
      box-shadow: 0 4px 14px rgba(52,109,153,0.08);
      box-sizing: border-box;
    }

    .mobile-file-pill .file-index {
      flex: 0 0 auto;
      min-width: 24px;
      height: 24px;
      border-radius: 999px;
      background: var(--primary);
      color: #fff;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      font-size: 0.72rem;
      font-weight: 800;
    }

    .mobile-file-pill .file-name {
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }

    .before-diagnosis-photo-frame {
      height: 260px;
    }

    .card {
      padding: 1rem;
      border-radius: 16px;
    }

    .score-value {
      font-size: 2.6rem;
    }

    .stButton > button,
    .stDownloadButton > button,
    .stDownloadButton button {
      min-height: 44px;
      font-size: 0.9rem;
      padding: 0.62rem 1rem;
    }
  }
</style>
"""



def initialize_app() -> None:
    """各再実行でページ設定とグローバルCSSを確実に適用する。"""
    load_dotenv()
    st.set_page_config(
        page_title="5S アドバイスシステム",
        page_icon="",
        layout="wide",
        initial_sidebar_state="collapsed" if should_hide_settings_panel() else "expanded",
    )
    st.markdown(GLOBAL_CSS, unsafe_allow_html=True)
    if should_hide_settings_panel():
        st.markdown(
            """
            <style>
            section[data-testid="stSidebar"] {
              display: none !important;
            }
            button[kind="header"][aria-label*="sidebar"],
            button[kind="header"][aria-label*="Sidebar"] {
              display: none !important;
            }
            </style>
            """,
            unsafe_allow_html=True,
        )


# ─── Gemini 初期化 ────────────────────────────────────────────────────────────────
@st.cache_resource
def init_gemini(api_key: str, model_name: str = "gemini-2.5-flash-lite"):
    """Gemini API を初期化する"""
    genai.configure(api_key=api_key)
    return genai.GenerativeModel(model_name)


# ─── 5S 診断プロンプト ────────────────────────────────────────────────────────────
def build_5s_prompt(location: str) -> str:
    detail_instruction = "各項目3〜5文で具体的に"
    knowledge = get_knowledge_context()

    return f"""
あなたは製造業・職場環境改善の専門家です。
以下の「5S・2S ナレッジ」を判断基準として、アップロードされた写真を診断してください。

=== 5S・2S ナレッジ ===
{knowledge}
=== ナレッジここまで ===

【対象場所】: {location if location else "不明（写真から判断）"}
【詳細レベル】: {detail_instruction}

以下のJSON形式で必ず回答してください（日本語で）：

{{
  "overall_score": <0〜100の整数。総合5Sスコア>,
  "summary": "<写真全体の印象と総評（2〜3文）>",
  "scene_category": "<以下のいずれかから最も近いものを選択: desk, drawer, cabinet, tool_board, factory_floor, other>",
  "seiri": {{
    "score": <0〜100>,
    "title": "整理（Seiri）",
    "comment": "<不要なものを除去できているかの評価と改善点>",
    "priority": "<高／中／低>"
  }},
  "seiton": {{
    "score": <0〜100>,
    "title": "整頓（Seiton）",
    "comment": "<必要なものが定位置に配置されているかの評価と改善点>",
    "priority": "<高／中／低>"
  }},
  "seiso": {{
    "score": <0〜100>,
    "title": "清掃（Seiso）",
    "comment": "<清潔に保たれているかの評価と改善点>",
    "priority": "<高／中／低>"
  }},
  "seiketsu": {{
    "score": <0〜100>,
    "title": "清潔（Seiketsu）",
    "comment": "<清潔な状態が維持される仕組みがあるかの評価と改善点>",
    "priority": "<高／中／低>"
  }},
  "shitsuke": {{
    "score": <0〜100>,
    "title": "しつけ（Shitsuke）",
    "comment": "<ルールが習慣化されているかの評価と改善点>",
    "priority": "<高／中／低>"
  }},
  "action_items": [
    "<即実施すべき改善アクション1>",
    "<即実施すべき改善アクション2>",
    "<中長期的な改善アクション>"
  ]
}}

JSONのみを返してください。余分な説明文は不要です。
"""


# ─── Gemini で診断実行 ───────────────────────────────────────────────────────────
def analyze_image(model, image: Image.Image, location: str) -> dict:
    import json, re
    prompt = build_5s_prompt(location)

    # PIL → bytes
    buf = io.BytesIO()
    image.save(buf, format="JPEG", quality=90)
    image_bytes = buf.getvalue()

    response = model.generate_content([
        prompt,
        {"mime_type": "image/jpeg", "data": image_bytes},
    ])

    raw = response.text.strip()
    match = re.search(r"\{.*\}", raw, re.DOTALL)
    if match:
        return json.loads(match.group())
    raise ValueError("有効なJSONが取得できませんでした:\n" + raw)





# ─── スコア→グレード ─────────────────────────────────────────────────────────────
def score_to_grade(score: int) -> tuple[str, str]:
    if score >= 90: return "S", "#f093fb"
    if score >= 75: return "A", "#667eea"
    if score >= 60: return "B", "#22c55e"
    if score >= 40: return "C", "#eab308"
    return "D", "#ef4444"


# ─── 優先度→色 ───────────────────────────────────────────────────────────────────
def priority_color(p: str) -> str:
    return {"高": "#ef4444", "中": "#eab308", "低": "#22c55e"}.get(p, "#888")


def build_expert_comment(
    edited_summary: str,
    edited_seiri_comment: str,
    edited_seiton_comment: str,
    edited_actions: list[str],
) -> str:
    """診断士コメント列へ保存するため、編集内容を読みやすい1テキストに整形する。"""
    lines = [
        "【総評】",
        (edited_summary or "").strip() or "記載なし",
        "",
        "【整理コメント】",
        (edited_seiri_comment or "").strip() or "記載なし",
        "",
        "【整頓コメント】",
        (edited_seiton_comment or "").strip() or "記載なし",
        "",
        "【改善アクション】",
    ]
    cleaned_actions = [(a or "").strip() for a in edited_actions if (a or "").strip()]
    if cleaned_actions:
        lines.extend(f"{i + 1}. {action}" for i, action in enumerate(cleaned_actions))
    else:
        lines.append("記載なし")
    return "\n".join(lines)


def generate_record_id() -> str:
    """シート更新用の短い一意IDを発行する。"""
    return f"rec-{datetime.now().strftime('%Y%m%d%H%M%S')}-{uuid.uuid4().hex[:8]}"


def get_app_mode(mode: str | None = None) -> str:
    raw = (mode or os.getenv("APP_MODE", "expert")).strip().lower()
    return "member" if raw == "member" else "expert"


def get_runtime_secret(name: str, default: str = "") -> str:
    try:
        if name in st.secrets:
            value = st.secrets.get(name)
            return "" if value is None else str(value)
    except Exception:
        pass
    return os.getenv(name, default)


def should_hide_settings_panel() -> bool:
    raw = get_runtime_secret("HIDE_SETTINGS_PANEL", "")
    return raw.strip().lower() in {"1", "true", "yes", "on"}


# ─── UI 描画：診断結果 ────────────────────────────────────────────────────────────
def render_results(result: dict, img, mode: str = "expert"):
    st.markdown("---")

    st.markdown("<div style='color:#346D99; font-weight:700; margin-bottom:0.6rem;'>診断写真</div>", unsafe_allow_html=True)

    col_img, col_info = st.columns([1, 1], gap="large")

    with col_img:
        b64_before = pil_image_to_b64_jpeg(img, size=(1200, 900), quality=88)
        st.markdown(
            f"""
<div class="before-diagnosis-photo-frame">
  <img src="data:image/jpeg;base64,{b64_before}" alt="診断写真"
       class="before-diagnosis-photo-img" />
</div>
""",
            unsafe_allow_html=True,
        )

    with col_info:
        # 総合スコア
        overall = result.get("overall_score", 0)
        grade, grade_color = score_to_grade(overall)

        st.markdown(f"""
<div style="display: flex; flex-direction: column; justify-content: space-between; height: 100%;">
  <div class="card" style="text-align:center; margin-bottom: 1rem;">
    <div style="color:#94a3b8;font-size:0.78rem;margin-bottom:0.4rem;letter-spacing:0.06em;">TOTAL SCORE</div>
    <div class="score-value">{overall}</div>
    <div style="color:{grade_color};font-size:1.5rem;font-weight:700;margin:0.3rem 0;">Grade {grade}</div>
    <div style="color:#94a3b8;font-size:0.75rem;">/ 100 点</div>
  </div>
  
  <div class="card" style="flex-grow: 1; display: flex; flex-direction: column; justify-content: center;">
    <div class="section-heading">総評</div>
    <div style="color:#475569;line-height:1.7;font-size:0.94rem;">
      {result.get("summary", "")}
    </div>
  </div>
</div>
""", unsafe_allow_html=True)

    _current_img_bytes = result.get("_pdf_image_bytes") or pil_image_to_jpeg_bytes(img, quality=88)
    img_fname = st.session_state.get("current_report_fname", "")
    current_record_id = str(result.get("_record_id") or "")

    # 総評の右下に個別DLボタン
    spacer_col, button_col = st.columns([3, 2], gap="small")
    with button_col:
        try:
            pdf_bytes_quick = generate_pdf(
                result=result,
                image_bytes=_current_img_bytes if '_current_img_bytes' in dir() else b"",
                filename=img_fname if 'img_fname' in dir() else "",
                company=st.session_state.get("main_company", ""),
                location=st.session_state.get("main_location", ""),
                edited_summary="",
                edited_actions=[],
            )
            st.download_button(
                label="個別で診断結果をダウンロード",
                data=pdf_bytes_quick,
                file_name=f"5S診断レポート_{datetime.now().strftime('%Y%m%d')}.pdf",
                mime="application/pdf",
                key=f"dl_pdf_quick_{id(result)}",
                use_container_width=True,
            )
        except Exception as e:
            st.error(f"PDF生成エラー: {e}")

    st.markdown("---")

    # 2S 各項目のみ表示
    s_keys = ["seiri", "seiton"]

    st.markdown("<h3 class='score-detail-heading'>2S 診断スコア詳細</h3>", unsafe_allow_html=True)
    for key in s_keys:
        item = result.get(key, {})
        score = item.get("score", 0)
        title = item.get("title", key)
        comment = item.get("comment", "")
        priority = item.get("priority", "中")

        with st.expander(f"{title}　{score}点　/ 優先度：{priority}", expanded=False):
            st.progress(score / 100)
            st.markdown(f"""
            <div class="result-box">{comment}</div>
            <span style="font-size:0.8rem;color:{priority_color(priority)};font-weight:600;">
              改善優先度：{priority}
            </span>
            """, unsafe_allow_html=True)

    # アクションアイテム
    actions = result.get("action_items", [])
    if actions:
        st.markdown("<h3 style='color:#346D99; margin-top:1.5rem; margin-bottom:0.5rem;'>すぐに実行できる改善アクション</h3>", unsafe_allow_html=True)
        st.markdown("<p style='color:var(--text-sub); font-size:0.9rem; margin-bottom:1rem;'>AIの診断結果に基づき、優先して取り組むべき具体的な改善策です。</p>", unsafe_allow_html=True)
        num_cls = ["action-num", "action-num mid", "action-num low"]
        for i, action in enumerate(actions):
            cls = num_cls[i] if i < len(num_cls) else "action-num low"
            st.markdown(f"""
            <div class="action-card">
              <div class="{cls}">{i+1}</div>
              <div style="color:#1e293b;font-size:0.94rem;line-height:1.6;">{action}</div>
            </div>
            """, unsafe_allow_html=True)

    # ── PDF編集・DLセクション ──
    st.markdown("---")
    st.markdown(
        "<h3 class='report-download-heading'>"
        "レポート作成・ダウンロード</h3>",
        unsafe_allow_html=True
    )

    expander_title = "内容を編集してからダウンロード" if mode == "expert" else "内容を調整してからダウンロード"

    # ── 編集フォーム ──
    with st.expander(expander_title, expanded=False):

        # 総評
        st.markdown(
            "<p style='color:#1e293b; font-weight:bold; margin-bottom:4px;'>総評</p>",
            unsafe_allow_html=True
        )
        edited_summary = st.text_area(
            label="総評",
            value=result.get("summary", ""),
            key=f"edit_summary_{id(result)}",
            label_visibility="collapsed"
        )

        # 診断スコア詳細
        st.markdown(
            "<p style='color:#1e293b; font-weight:bold; margin-bottom:4px;'>診断スコア詳細（整理）</p>",
            unsafe_allow_html=True
        )
        seiri = result.get("seiri", {})
        edited_seiri_comment = st.text_area(
            label="整理コメント",
            value=seiri.get("comment", ""),
            key=f"edit_seiri_comment_{id(result)}",
            label_visibility="collapsed"
        )

        st.markdown(
            "<p style='color:#1e293b; font-weight:bold; margin-bottom:4px;'>診断スコア詳細（整頓）</p>",
            unsafe_allow_html=True
        )
        seiton = result.get("seiton", {})
        edited_seiton_comment = st.text_area(
            label="整頓コメント",
            value=seiton.get("comment", ""),
            key=f"edit_seiton_comment_{id(result)}",
            label_visibility="collapsed"
        )

        # 改善アクション
        actions = result.get("action_items") or []
        edited_actions = []
        for i, action in enumerate(actions):
            st.markdown(
                f"<p style='color:#1e293b; font-weight:bold; margin-bottom:4px;'>改善アクション {i+1}</p>",
                unsafe_allow_html=True
            )
            edited = st.text_area(
                label=f"改善アクション{i+1}",
                value=action,
                key=f"edit_action_{id(result)}_{i}",
                label_visibility="collapsed"
            )
            edited_actions.append(edited)

        if mode == "expert":
            if st.button("診断士コメントを保存して確定", key=f"save_confirm_{id(result)}", use_container_width=True):
                try:
                    from sheets_client import update_expert_review
                    if not current_record_id:
                        raise ValueError("record_id が見つかりません。再診断後にお試しください。")
                    expert_comment = build_expert_comment(
                        edited_summary,
                        edited_seiri_comment,
                        edited_seiton_comment,
                        edited_actions,
                    )
                    update_expert_review(current_record_id, expert_comment, status="確定")
                    st.success("診断士コメントを保存し、ステータスを「確定」にしました")
                except Exception as e:
                    st.error(f"保存エラー: {e}")

    # 個別PDFダウンロード
    spacer_col, button_col = st.columns([3, 2], gap="small")
    with button_col:
        try:
            pdf_bytes = generate_pdf(
                result=result,
                image_bytes=_current_img_bytes if '_current_img_bytes' in dir() else b"",
                filename=img_fname if 'img_fname' in dir() else "",
                company=st.session_state.get("main_company", ""),
                location=st.session_state.get("main_location", ""),
                edited_summary=edited_summary if 'edited_summary' in dir() else "",
                edited_actions=edited_actions if 'edited_actions' in dir() else [],
            )
            st.download_button(
                label="個別で診断結果をダウンロード",
                data=pdf_bytes,
                file_name=f"5S診断レポート_{datetime.now().strftime('%Y%m%d')}.pdf",
                mime="application/pdf",
                key=f"dl_pdf_{id(result)}",
                use_container_width=True,
            )
        except Exception as e:
            st.error(f"PDF生成エラー: {e}")

    return result


def _set_diagnosis_index(idx: int) -> None:
    st.session_state["selected_idx"] = int(idx)


def _nav_diagnosis(delta: int) -> None:
    res = st.session_state.get("results") or []
    if not res:
        return
    cur = int(st.session_state.get("selected_idx", 0))
    n = len(res)
    st.session_state["selected_idx"] = max(0, min(n - 1, cur + int(delta)))


@st.fragment
def render_diagnosis_results_fragment() -> None:
    """左リストの切り替え・右詳細のみ再描画し、ページ全体のリロードを避ける。"""
    app_mode = st.session_state.get("app_mode", "expert")
    results = st.session_state.get("results") or []
    if not results:
        return
    if "selected_idx" not in st.session_state:
        st.session_state["selected_idx"] = 0
    sel = min(int(st.session_state["selected_idx"]), len(results) - 1)
    st.session_state["selected_idx"] = sel

    st.markdown("<hr style='margin:1.5rem 0'>", unsafe_allow_html=True)
    left_col, right_col = st.columns([1, 3], gap="large")

    with left_col:
        st.markdown('<div class="list-panel-title">診断済みリスト</div>', unsafe_allow_html=True)
        def _bulk_download():
            reports = []
            gallery_items = st.session_state.get("gallery_images") or []
            for idx, (fname, img, res) in enumerate(results):
                if res is None:
                    continue
                try:
                    image_bytes = res.get("_pdf_image_bytes") or b""
                    if not image_bytes and idx < len(gallery_items):
                        image_bytes = gallery_items[idx].get("data", b"") or b""
                    if not image_bytes:
                        image_bytes = pil_image_to_jpeg_bytes(img, quality=88)
                    pdf = generate_pdf(
                        result=res,
                        image_bytes=image_bytes,
                        filename=fname,
                        company=st.session_state.get("main_company", ""),
                        location=st.session_state.get("main_location", ""),
                    )
                    reports.append((fname, pdf))
                except Exception:
                    pass
            return generate_zip(reports)

        zip_bytes_top = _bulk_download()
        st.download_button(
            label="全件ZIPダウンロード",
            data=zip_bytes_top,
            file_name=f"5S診断レポート一括_{datetime.now().strftime('%Y%m%d')}.zip",
            mime="application/zip",
            key="bulk_dl_zip_top",
            use_container_width=True
        )

        for i, (fname, img, result) in enumerate(results):
            is_sel = i == sel
            overall = result.get("overall_score", None) if result else None
            score_txt = f"{overall}点" if overall is not None else "診断失敗"
            score_col = "var(--primary)" if overall is not None else "#ef4444"
            short = (fname[:15] + "…") if len(fname) > 15 else fname
            safe_short = html.escape(short)
            safe_score = html.escape(score_txt)

            with st.container(border=is_sel):
                thumb = img.copy()
                thumb.thumbnail((280, 160), Image.LANCZOS)
                tb64 = pil_image_to_b64_jpeg(thumb, size=(320, 200), quality=72)
                st.markdown(
                    f"""
<div style="line-height:0;border-radius:6px;overflow:hidden;">
  <img src="data:image/jpeg;base64,{tb64}" alt=""
       style="width:100%;height:62px;object-fit:cover;display:block;" />
</div>
""",
                    unsafe_allow_html=True,
                )
                st.markdown(
                    f"<div style='font-size:0.7rem;color:#475569;overflow:hidden;"
                    f"text-overflow:ellipsis;white-space:nowrap;'>{i + 1}. {safe_short}</div>",
                    unsafe_allow_html=True,
                )
                st.markdown(
                    f"<div style='font-size:0.77rem;font-weight:700;color:{score_col};'>"
                    f"{safe_score}</div>",
                    unsafe_allow_html=True,
                )
                st.markdown(
                    f'<img alt="score-btn-anchor-{i}" src="data:," style="display:none;" />',
                    unsafe_allow_html=True,
                )
                st.button(
                    "✓ 表示中" if is_sel else "スコア・総評を見る",
                    key=f"diag_pick_{i}",
                    use_container_width=True,
                    type="primary" if is_sel else "secondary",
                    disabled=is_sel,
                    on_click=_set_diagnosis_index,
                    args=(i,),
                )

        zip_bytes_bottom = _bulk_download()
        st.download_button(
            label="全件ZIPダウンロード",
            data=zip_bytes_bottom,
            file_name=f"5S診断レポート一括_{datetime.now().strftime('%Y%m%d')}.zip",
            mime="application/zip",
            key="bulk_dl_zip_bottom",
            use_container_width=True
        )

    with right_col:
        fname, img, result = results[sel]
        safe_fname = html.escape(fname)

        nv_l, nv_c, nv_r = st.columns([1, 7, 1])
        with nv_l:
            st.button(
                "←",
                use_container_width=True,
                key="nav_prev",
                disabled=sel <= 0,
                on_click=_nav_diagnosis,
                args=(-1,),
            )
        with nv_c:
            st.markdown(
                f"<div style='text-align:center;padding:0.35rem 0;"
                f"font-size:0.85rem;color:#475569;'>"
                f"{sel + 1} / {len(results)}&ensp;｜&ensp;{safe_fname}</div>",
                unsafe_allow_html=True,
            )
        with nv_r:
            st.button(
                "→",
                use_container_width=True,
                key="nav_next",
                disabled=sel >= len(results) - 1,
                on_click=_nav_diagnosis,
                args=(1,),
            )

        if result is None:
            st.error(f"{fname} の診断に失敗しました")
        else:
            st.session_state["current_report_fname"] = fname
            render_results(result, img, app_mode)


# ─── サイドバー ──────────────────────────────────────────────────────────────────
def render_sidebar() -> tuple[str, str, str]:
    if should_hide_settings_panel():
        api_key = get_runtime_secret("GEMINI_API_KEY", "")
        model_name = get_runtime_secret("GEMINI_MODEL", "gemini-2.5-flash-lite")
        detail_level = "標準"
        return api_key, detail_level, model_name

    with st.sidebar:
        st.markdown("""
        <div style="padding:1.4rem 0 1rem;">
          <div style="font-size:1.15rem;font-weight:700;color:#346D99;letter-spacing:-0.01em;">5S アドバイスシステム</div>
          <div style="color:#94a3b8;font-size:0.75rem;margin-top:0.2rem;">powered by Gemini AI</div>
        </div>
        <hr style="margin-bottom:1.2rem;">
        """, unsafe_allow_html=True)

        st.markdown("#### 設定")

        api_key = get_runtime_secret("GEMINI_API_KEY", "")
        if api_key:
            st.markdown(
                "<div style='color:#475569;font-size:0.92rem;margin-bottom:0.75rem;'>"
                "Gemini API キー: 設定済み</div>",
                unsafe_allow_html=True,
            )
        else:
            api_key = st.text_input(
                "Gemini API キー",
                value="",
                type="password",
                placeholder="AIza...",
                help="Google AI Studio から取得したAPIキーを入力してください",
                key="sidebar_api_key"
            )

        model_name = st.selectbox(
            "使用モデル",
            [
                "gemini-2.5-flash-lite",   # 無料枠 ✅ 推奨
                "gemini-2.0-flash-lite",   # 無料枠 ✅
                "gemini-2.5-flash",        # 有料
                "gemini-2.0-flash",        # 有料
                "gemini-1.5-pro",          # 有料
            ],
            help="無料枠で使うなら gemini-2.5-flash-lite または gemini-2.0-flash-lite を選択",
        )
        st.markdown("<hr>", unsafe_allow_html=True)
        st.markdown(
            "<div style='color:#94a3b8;font-size:0.75rem;'>"
            "※ 本番環境では secrets、ローカルでは .env で管理</div>",
            unsafe_allow_html=True
        )

    detail_level = "標準"
    return api_key, detail_level, model_name


@st.fragment
def render_photo_gallery() -> None:
    """サムネイルグリッドのみを描画。削除時はフルリロードせずこの断片だけ再実行する。"""
    st.markdown(
        """
<style>
/* サムネイル画像（alt が gallery-del-anchor）の直後の削除ボタンを画像枠の右上に重ねる */
div[data-testid="stVerticalBlock"]:has(img[alt^="gallery-del-anchor"]) {
  position: relative !important;
}
div[data-testid="stVerticalBlock"]:has(img[alt^="gallery-del-anchor"])
  > div.element-container:has(img[alt^="gallery-del-anchor"])
  + div.element-container {
  position: absolute !important;
  top: 6px !important;
  right: 6px !important;
  z-index: 10 !important;
  width: auto !important;
}
div[data-testid="stVerticalBlock"]:has(img[alt^="gallery-del-anchor"])
  > div.element-container:has(img[alt^="gallery-del-anchor"])
  + div.element-container
  .stButton
  > button {
  padding: 0.1rem 0.4rem !important;
  min-height: 1.7rem !important;
  min-width: 1.7rem !important;
  width: auto !important;
  background: rgba(255, 255, 255, 0.95) !important;
  color: #64748b !important;
  border: 1px solid #cbd5e1 !important;
  box-shadow: 0 1px 4px rgba(0, 0, 0, 0.12) !important;
}
</style>
""",
        unsafe_allow_html=True,
    )

    gallery = st.session_state.get("gallery_images", [])
    n = len(gallery)
    if n:
        images = [Image.open(io.BytesIO(x["data"])).convert("RGB") for x in gallery]
    else:
        images = []

    for row in range(2):
        cols = st.columns(5, gap="small")
        for col_i in range(5):
            idx = row * 5 + col_i
            with cols[col_i]:
                if idx < n:
                    item = gallery[idx]
                    b64 = pil_image_to_b64_jpeg(images[idx])
                    fn = item["name"]
                    lbl = (fn[:15] + "…") if len(fn) > 15 else fn
                    safe_lbl = html.escape(lbl)
                    del_anchor = html.escape(f"gallery-del-anchor-{item['id']}", quote=True)
                    st.markdown(
                        f"""
<div class="gallery-thumb-card" style="position:relative;width:100%;border-radius:10px;overflow:hidden;border:2px solid #346D99;
            box-shadow:0 2px 8px rgba(52,109,153,0.12);margin-bottom:4px;">
  <img src="data:image/jpeg;base64,{b64}" alt="{del_anchor}"
       style="width:100%;vertical-align:middle;display:block;aspect-ratio:4/3;object-fit:cover;background:#f8fafc;" />
  <div style="position:absolute;top:6px;left:6px;z-index:2;background:rgba(52,109,153,0.82);color:white;
              font-size:0.6rem;font-weight:700;padding:1px 6px;border-radius:999px;line-height:1.6;">{idx + 1}</div>
  <div style="position:absolute;bottom:0;left:0;right:0;background:rgba(30,41,59,0.55);color:#fff;
              font-size:0.68rem;padding:3px 6px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">{safe_lbl}</div>
</div>
<div class="mobile-file-pill">
  <span class="file-index">{idx + 1}</span>
  <span class="file-name">{safe_lbl}</span>
</div>
""",
                        unsafe_allow_html=True,
                    )
                    if st.button(
                        "✕",
                        key=f"del_img_{item['id']}",
                        help="この写真をリストから削除",
                        use_container_width=False,
                    ):
                        st.session_state.gallery_images = [
                            x
                            for x in st.session_state.gallery_images
                            if x.get("id") != item["id"]
                        ]
                        if not st.session_state.gallery_images:
                            _clear_gallery_and_diagnosis_state()
                            st.rerun()
                        else:
                            st.rerun(scope="fragment")
                else:
                    st.markdown(
                        """
                        <div class="thumb-cell empty"
                             style="min-height:120px;margin-top:1.5rem;border-radius:10px;
                                    display:flex;align-items:center;justify-content:center;">
                            +
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )

    # 2段グリッドの直下・右下相当（5列のうち右端）に一括クリア
    r_clear = st.columns(5, gap="small")
    for i in range(4):
        with r_clear[i]:
            st.empty()
    with r_clear[4]:
        if n > 0:
            st.markdown(
                '<img alt="clear-btn-anchor" src="data:," style="display:none;" />',
                unsafe_allow_html=True,
            )
            if st.button(
                "一括クリア",
                key="gallery_clear_all",
                use_container_width=True,
            ):
                _clear_gallery_and_diagnosis_state()
                st.rerun()


# ─── メイン ──────────────────────────────────────────────────────────────────────
def main(mode: str | None = None):
    initialize_app()
    app_mode = get_app_mode(mode)
    st.session_state["app_mode"] = app_mode
    st.session_state.setdefault("gallery_images", [])
    st.session_state.setdefault("gallery_uploader_key", 0)

    api_key, detail_level, model_name = render_sidebar()

    # ヘッダー
    hero_title = "5S 現場改善エンジン" if app_mode == "expert" else "5S 現場改善エンジン"
    hero_subtitle = "写真1枚で、現場の課題と改善策が見える"
    mode_label = "カジコン用" if app_mode == "expert" else "会員用"
    st.markdown(
        f"""
    <div class="mode-badge-wrap">
      <div class="mode-badge">{mode_label}</div>
    </div>
    <div class="hero-title">{hero_title}</div>
    <div class="hero-subtitle">{hero_subtitle}</div>
    """,
        unsafe_allow_html=True,
    )

    # バッジ行
    st.markdown("""
    <div class="badge-row">
      <span class="s-badge">整理</span>
      <span class="s-badge">整頓</span>
      <span class="s-badge">清掃</span>
      <span class="s-badge">清潔</span>
      <span class="s-badge">しつけ</span>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("<div style='margin-bottom:1.5rem;'></div>", unsafe_allow_html=True)

    if app_mode == "member" and not is_member_login_disabled():
        member_auth = st.session_state.get("member_auth")
        if not member_auth:
            render_member_login()
            return

        action_col_l, action_col_r = st.columns([6, 1])
        with action_col_r:
            if st.button("ログアウト", key="member_logout_btn", use_container_width=True):
                _logout_member_user()
                st.rerun()

    # APIキー確認
    if not api_key or api_key == "your_api_key_here":
        st.warning("サイドバーに Gemini API キーを入力してください。\n\n"
                   "[Google AI Studio](https://aistudio.google.com/app/apikey) から無料で取得できます。")
        st.stop()

    _, center_inputs, _ = st.columns([1, 3, 1])
    with center_inputs:
        col_company, col_location = st.columns(2)
        with col_company:
            company = st.text_input(
                "会社名（必須）",
                placeholder="例：株式会社〇〇",
                key="main_company",
                value=st.session_state.get("main_company", ""),
            )
        with col_location:
            location = st.text_input(
                "診断場所（必須）",
                placeholder="例：製造ライン、倉庫",
                key="main_location",
                value=st.session_state.get("main_location", ""),
            )

    # ─── 画像アップロード案内 ──────────────────────────────────────
    st.markdown("""
    <div style="background-color:#EEF5FB; border-left:4px solid #346D99; padding:1rem 1.4rem; border-radius:8px; margin-bottom:1rem;">
        <div style="color:#346D99; font-weight:700; font-size:1.05rem; margin-bottom:0.4rem;">診断する写真をアップロード</div>
        <ul style="color:#475569; font-size:0.9rem; margin:0; padding-left:1.2rem; line-height:1.6;">
            <li>下の点線枠内にファイルを<b>ドラッグ＆ドロップ</b>するか、<b>Browse files</b>ボタンから選択してください。</li>
            <li>アップロード可能な形式：<b>JPG, JPEG, PNG, WEBP</b></li>
            <li>一度にアップロードできる枚数上限：<b>最大 10 枚</b></li>
        </ul>
    </div>
    """, unsafe_allow_html=True)

    # ─── 画像アップロード本体 ──────────────────────────────────────────
    uploaded_files = st.file_uploader(
        label="写真アップロード領域",
        label_visibility="collapsed",
        type=["jpg", "jpeg", "png", "webp"],
        accept_multiple_files=True,
        key=f"gallery_uploader_{st.session_state['gallery_uploader_key']}",
    )

    MAX_IMAGES = 10
    if uploaded_files and len(uploaded_files) > MAX_IMAGES:
        st.error(
            f"枚数制限\n\n"
            f"最大{MAX_IMAGES}枚まで対応しています。"
            f"アップローダーで選ばれたファイルのうち、先頭の{MAX_IMAGES}枚のみ取り込みます。"
        )
        uploaded_files = uploaded_files[:MAX_IMAGES]

    for _it in st.session_state.gallery_images:
        if "id" not in _it:
            _it["id"] = str(uuid.uuid4())
        if "digest" not in _it:
            _it["digest"] = hashlib.md5(_it["data"]).hexdigest()

    # ギャラリーが空のときは署名を捨てる（消したあと同じ選択でも再マージできる）
    if len(st.session_state.gallery_images) == 0:
        st.session_state.pop("_gallery_upload_sig", None)

    # アップロード選択が変わったときだけマージ（署名は file_id または内容 MD5 で区別）
    upload_entries: list[tuple[str, str, bytes]] = []
    if uploaded_files:
        for f in uploaded_files:
            dig, data = _upload_digest_for_file(f)
            upload_entries.append((f.name, dig, data))
    upload_sig = tuple((n, d) for n, d, _ in upload_entries)

    if upload_sig != st.session_state.get("_gallery_upload_sig"):
        st.session_state["_gallery_upload_sig"] = upload_sig
        if upload_entries:
            existing_keys = {_gallery_item_key(x) for x in st.session_state.gallery_images}
            for name, dig, data in upload_entries:
                key = (name, dig)
                if key not in existing_keys and len(st.session_state.gallery_images) < MAX_IMAGES:
                    st.session_state.gallery_images.append(
                        {
                            "id": str(uuid.uuid4()),
                            "name": name,
                            "data": data,
                            "digest": dig,
                        }
                    )
                    existing_keys.add(key)

    gallery = st.session_state.gallery_images
    n = len(gallery)
    if n:
        images = [Image.open(io.BytesIO(x["data"])).convert("RGB") for x in gallery]
    else:
        images = []

    render_photo_gallery()

    if images:
        diagnose = st.button(
            f"{n} 枚を一括診断する  ›",
            type="primary",
            use_container_width=True,
        )

        if diagnose:
            if not company.strip() or not location.strip():
                st.error("会社名と診断場所は必須です。入力してから診断してください。")
                st.stop()
            if uploaded_files and len(uploaded_files) > MAX_IMAGES:
                uploaded_files = uploaded_files[:MAX_IMAGES]
            st.session_state.pop("show_limit_warning", None)
            if "results" in st.session_state:
                del st.session_state["results"]
            results = []
            saved_count = 0
            diagnose_status = st.empty()
            diagnose_status.markdown(
                '<p class="diagnose-progress-msg">分析を開始します...</p>',
                unsafe_allow_html=True,
            )
            progress = st.progress(0)
            model = init_gemini(api_key, model_name)
            for idx, (img, item) in enumerate(zip(images, gallery)):
                fname_i = item["name"]
                diagnose_status.markdown(
                    f'<p class="diagnose-progress-msg">'
                    f"{html.escape(fname_i)} を分析中... ({idx + 1}/{n})</p>",
                    unsafe_allow_html=True,
                )
                # 1枚目の API 待ち中も 0% のままに見えないよう、着手時に少し進める
                progress.progress(min((idx + 0.12) / n, 1.0) if n else 0)
                last_error = None
                for attempt in range(3):
                    try:
                        res = analyze_image(model, img, location)
                        record_id = generate_record_id()
                        image_bytes = pil_image_to_jpeg_bytes(img, quality=88)
                        res["_record_id"] = record_id
                        res["_pdf_image_bytes"] = image_bytes
                        results.append((fname_i, img, res))
                        try:
                            from sheets_client import save_to_sheets

                            save_to_sheets(res, location, fname_i, record_id, app_mode, company, image_bytes)
                            saved_count += 1
                        except Exception as e:
                            print(f"[Google Sheets保存エラー] {e}", flush=True)
                        break
                    except Exception as e:
                        last_error = e
                        err_str = str(e)
                        if ("504" in err_str or "timed out" in err_str.lower()) and attempt < 2:
                            continue
                        st.error(f"{fname_i} の診断でエラー: {last_error}")
                        results.append((fname_i, img, None))
                        break
                progress.progress((idx + 1) / n if n else 0)
            diagnose_status.markdown(
                '<p class="diagnose-progress-msg">診断完了</p>',
                unsafe_allow_html=True,
            )
            progress.progress(1.0)
            st.session_state["results"] = results
            st.session_state["selected_idx"] = 0
            st.rerun()

    # ─── 結果表示: 左サムネイルリスト ＋ 右詳細（@st.fragment で部分更新） ───
    if "results" in st.session_state and st.session_state["results"]:
        render_diagnosis_results_fragment()


if __name__ == "__main__":
    main()
