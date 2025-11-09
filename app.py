# app.py
# ------------------------------------------------------------
# 故障メール → 正規表現抽出 → 既存テンプレ(.xlsm)へ書込み → ダウンロード
# 3ステップUI / パスコード認証 / 編集不可 / 折りたたみ表示（時系列）
# 仕様反映：
#   - 曜日：日本語（例：月）
#   - 複数行：最大5行。超過は「…」付与
#   - 通報者：原文そのまま（様/電話番号含む）
#   - ファイル名：管理番号_物件名_日付（yyyymmdd）
#   - マクロ保持対応（keep_vba=True）
# ------------------------------------------------------------
import io
import re
import unicodedata
from datetime import datetime, timedelta, timezone
from typing import Dict, Optional, Tuple, List
import os
import sys
import traceback
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage  # 画像機能は将来用
import streamlit as st

# ---- 基本設定 ------------------------------------------------
JST = timezone(timedelta(hours=9))
APP_TITLE = "故障報告書自動生成"

def _get_passcode() -> str:
    """
    PASSCODEの安全取得。
    優先度: st.secrets -> 環境変数 -> 開発用デフォルト("")
    """
    try:
        val = st.secrets.get("APP_PASSCODE")
        if val:
            return str(val)
    except Exception:
        # st.secrets 未設定でも落ちないようにする
        pass
    env_val = os.getenv("APP_PASSCODE")
    if env_val:
        return str(env_val)
    # 開発用の空デフォルト（本番は必ずSecrets/環境変数で上書きする想定）
    return ""

SHEET_NAME = "緊急出動報告書（リンク付き）"
WEEKDAYS_JA = ["月", "火", "水", "木", "金", "土", "日"]

# -------------------------------------------------------------
# ✏️ 編集フィールド共通関数（どのStepでも利用可能）
# -------------------------------------------------------------
def editable_field(label, key, max_lines=1):
    """共通：左アイコン付きの編集UI（セッション未初期化でも安全にアクセス）"""
    if "extracted" not in st.session_state or st.session_state.extracted is None:
        st.session_state.extracted = {}
    data = st.session_state.extracted

    edit_key = f"edit_{key}"
    if edit_key not in st.session_state:
        st.session_state[edit_key] = False

    # 通常表示モード
    if not st.session_state[edit_key]:
        value = data.get(key) or ""
        lines = value.split("\n") if max_lines > 1 else [value]
        display_text = "<br>".join(lines)
        cols = st.columns([0.07, 0.93])
        with cols[0]:
            if st.button("✏️", key=f"btn_{key}", help=f"{label}を編集"):
                st.session_state[edit_key] = True
                st.rerun()
        with cols[1]:
            st.markdown(f"**{label}：**<br>{display_text}", unsafe_allow_html=True)

    # 編集モード
    else:
        st.markdown(f"✏️ **{label} 編集中**")
        value = data.get(key) or ""
        if max_lines == 1:
            new_val = st.text_input(f"{label}を入力", value=value, key=f"in_{key}")
        else:
            new_val = st.text_area(f"{label}を入力", value=value, height=max_lines * 25, key=f"ta_{key}")
        c1, c2 = st.columns([0.3, 0.7])
        with c1:
            if st.button("💾 保存", key=f"save_{key}"):
                st.session_state.extracted[key] = new_val
                st.session_state[edit_key] = False
                st.rerun()
        with c2:
            if st.button("❌ キャンセル", key=f"cancel_{key}"):
                st.session_state[edit_key] = False
                st.rerun()

# ====== テキスト整形・抽出ユーティリティ ======
def normalize_text(text: str) -> str:
    if not text:
        return ""
    t = unicodedata.normalize("NFKC", text)
    t = t.replace("：", ":")
    t = t.replace("\t", " ").replace("\r\n", "\n").replace("\r", "\n")
    return t

def _search_one(pattern: str, text: str, flags=0) -> Optional[str]:
    m = re.search(pattern, text, flags)
    return m.group(1).strip() if m else None

def _search_span_between(labels: Dict[str, str], key: str, text: str) -> Optional[str]:
    lab = labels[key]
    others = [v for k, v in labels.items() if k != key]
    boundary = "|".join([f"(?:{v})" for v in others]) if others else r"$"
    pattern = rf"{lab}\s*(.+?)(?=\n(?:{boundary})|\Z)"
    m = re.search(pattern, text, flags=re.DOTALL | re.IGNORECASE)
    return m.group(1).strip() if m else None

def _try_parse_datetime(s: Optional[str]) -> Optional[datetime]:
    if not s:
        return None
    cand = s.strip()
    cand = cand.replace("年", "/").replace("月", "/").replace("日", "")
    cand = cand.replace("-", "/").replace("　", " ")
    for fmt in ("%Y/%m/%d %H:%M:%S", "%Y/%m/%d %H:%M", "%Y/%m/%d"):
        try:
            # naive -> JST
            dt = datetime.strptime(cand, fmt)
            return dt.replace(tzinfo=JST)
        except Exception:
            pass
    # pandas非依存で完結させるため、ここではこれ以上無理にパースしない
    return None

def _split_dt_components(dt: Optional[datetime]) -> Tuple[Optional[int], Optional[int], Optional[int], Optional[str], Optional[int], Optional[int]]:
    if not dt:
        return None, None, None, None, None, None
    dt = dt.astimezone(JST)
    y = dt.year
    m = dt.month
    d = dt.day
    wd = WEEKDAYS_JA[dt.weekday()]
    hh = dt.hour
    mm = dt.minute
    return y, m, d, wd, hh, mm

def _first_date_yyyymmdd(*vals) -> str:
    for v in vals:
        dt = _try_parse_datetime(v)
        if dt:
            return dt.strftime("%Y%m%d")
    return datetime.now(JST).strftime("%Y%m%d")

def minutes_between(a: Optional[str], b: Optional[str]) -> Optional[int]:
    s = _try_parse_datetime(a)
    e = _try_parse_datetime(b)
    if s and e:
        return int((e - s).total_seconds() // 60)
    return None

def _split_lines(text: Optional[str], max_lines: int = 5) -> List[str]:
    if not text:
        return []
    lines = [ln.strip() for ln in text.splitlines() if ln.strip() != ""]
    if len(lines) <= max_lines:
        return lines
    kept = lines[: max_lines - 1] + [lines[max_lines - 1] + "…"]
    return kept

# ====== 正規表現 抽出 ======
def extract_fields(raw_text: str) -> Dict[str, Optional[str]]:
    t = normalize_text(raw_text)

    # 件名由来の補助抽出
    subject_case = _search_one(r"件名:\s*【\s*([^】]+)\s*】", t, flags=re.IGNORECASE)
    subject_manageno = _search_one(r"件名:.*?【[^】]+】\s*([A-Z0-9\-]+)", t, flags=re.IGNORECASE)

    # 1行想定
    single_line = {
        "管理番号": r"管理番号\s*:\s*([A-Za-z0-9\-]+)",
        "物件名": r"物件名\s*:\s*(.+)",
        "住所": r"住所\s*:\s*(.+)",
        "窓口会社": r"窓口\s*:\s*(.+)",
        "メーカー": r"メーカー\s*:\s*(.+)",
        "制御方式": r"制御方式\s*:\s*(.+)",
        "契約種別": r"契約種別\s*:\s*(.+)",
        "受信時刻": r"受信時刻\s*:\s*([0-9/\-:\s]+)",
        "通報者": r"通報者\s*:\s*(.+)",
        "現着時刻": r"現着時刻\s*:\s*([0-9/\-:\s]+)",
        "完了時刻": r"完了時刻\s*:\s*([0-9/\-:\s]+)",
        "対応者": r"対応者\s*:\s*(.+)",
        "送信者": r"送信者\s*:\s*(.+)",
        "受付番号": r"受付番号\s*:\s*([0-9]+)",
        "受付URL": r"詳細はこちら\s*:\s*.*?(https?://\S+)",
        "現着完了登録URL": r"現着・完了登録はこちら\s*:\s*(https?://\S+)",
    }

    # 複数行想定（境界抽出）
    multiline_labels = {
        "受信内容": r"受信内容\s*:",
        "現着状況": r"現着状況\s*:",
        "原因": r"原因\s*:",
        "処置内容": r"処置内容\s*:",
        # 下記はフォーマット依存で複数行になりうるため残す
        "通報者": r"通報者\s*:",
        "対応者": r"対応者\s*:",
        "送信者": r"送信者\s*:",
        "現着時刻": r"現着時刻\s*:",
        "完了時刻": r"完了時刻\s*:",
    }

    out: Dict[str, Optional[str]] = {k: None for k in set(single_line.keys()) | set(multiline_labels.keys())}
    out.update({
        "案件種別(件名)": subject_case,
        "受付URL": None,
        "現着完了登録URL": None,
    })

    for k, pat in single_line.items():
        out[k] = _search_one(pat, t, flags=re.IGNORECASE | re.MULTILINE)

    if not out.get("管理番号") and subject_manageno:
        out["管理番号"] = subject_manageno

    for k in multiline_labels:
        span = _search_span_between(multiline_labels, k, t)
        if span:  # スパン抽出があれば優先（原文保持）
            out[k] = span

    dur = minutes_between(out.get("現着時刻"), out.get("完了時刻"))
    out["作業時間_分"] = str(dur) if dur is not None and dur >= 0 else None
    return out

# ====== テンプレ書き込み ======
def fill_template_xlsx(template_bytes: bytes, data: Dict[str, Optional[str]]) -> bytes:
    if not template_bytes:
        raise ValueError("テンプレートのバイト列が空です。")

    try:
        wb = load_workbook(io.BytesIO(template_bytes), keep_vba=True)
    except Exception as e:
        raise RuntimeError(f"テンプレートの読み込みに失敗しました（破損の可能性）: {e}") from e

    ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.active

    def fill_multiline(col_letter: str, start_row: int, text: Optional[str], max_lines: int = 5):
        # 事前にクリア
        for i in range(max_lines):
            ws[f"{col_letter}{start_row + i}"] = ""
        if not text:
            return
        lines = _split_lines(text, max_lines=max_lines)
        for idx, line in enumerate(lines[:max_lines]):
            ws[f"{col_letter}{start_row + idx}"] = line

    # ---- 単項目
    if data.get("管理番号"): ws["C12"] = data["管理番号"]
    if data.get("メーカー"): ws["J12"] = data["メーカー"]
    if data.get("制御方式"): ws["M12"] = data["制御方式"]
    if data.get("通報者"): ws["C14"] = data["通報者"]
    if data.get("対応者"): ws["L37"] = data["対応者"]

    # 任意：処理修理後
    pa = (st.session_state.get("processing_after") or data.get("処理修理後") or "").strip()
    if pa:
        ws["C35"] = pa

    # 所属
    if data.get("所属"): ws["C37"] = data["所属"]

    # B5/D5/F5 に現在日付（JST）
    now = datetime.now(JST)
    ws["B5"], ws["D5"], ws["F5"] = now.year, now.month, now.day

    # ---- 日時分解ブロック
    def write_dt_block(base_row: int, src_key: str):
        dt = _try_parse_datetime(data.get(src_key))
        y, m, d, wd, hh, mm = _split_dt_components(dt)
        cellmap = {"Y": f"C{base_row}", "Mo": f"F{base_row}", "D": f"H{base_row}",
                   "W": f"J{base_row}", "H": f"M{base_row}", "Min": f"O{base_row}"}
        if y is not None: ws[cellmap["Y"]] = y
        if m is not None: ws[cellmap["Mo"]] = m
        if d is not None: ws[cellmap["D"]] = d
        if wd is not None: ws[cellmap["W"]] = wd
        if hh is not None: ws[cellmap["H"]] = f"{hh:02d}"
        if mm is not None: ws[cellmap["Min"]] = f"{mm:02d}"

    write_dt_block(13, "受信時刻")
    write_dt_block(19, "現着時刻")
    write_dt_block(36, "完了時刻")

    # ---- 複数行
    fill_multiline("C", 15, data.get("受信内容"), max_lines=4)
    fill_multiline("C", 20, data.get("現着状況"))
    fill_multiline("C", 25, data.get("原因"))
    fill_multiline("C", 30, data.get("処置内容"))

    out = io.BytesIO()
    try:
        wb.save(out)
    except Exception as e:
        raise RuntimeError(f"Excel保存時に失敗しました: {e}") from e

    return out.getvalue()

def _sanitize_filename(name: str) -> str:
    # Windows等で不正な文字を避ける
    return re.sub(r'[\\/:*?"<>|]+', "_", name)

def build_filename(data: Dict[str, Optional[str]]) -> str:
    base_day = _first_date_yyyymmdd(data.get("現着時刻"), data.get("完了時刻"), data.get("受信時刻"))
    manageno = _sanitize_filename((data.get("管理番号") or "UNKNOWN").strip().replace("/", "_"))
    bname = _sanitize_filename((data.get("物件名") or "").strip().replace("/", "_"))
    if bname:
        return f"緊急出動報告書_{manageno}_{bname}_{base_day}.xlsm"
    return f"緊急出動報告書_{manageno}_{base_day}.xlsm"

# ====== Streamlit UI ======
st.set_page_config(page_title=APP_TITLE, layout="centered")
# タイトル非表示＋上部余白を最小化
st.markdown(
    """
    <style>
    header {visibility: hidden;}
    .block-container {padding-top: 0rem;}
    </style>
    """,
    unsafe_allow_html=True,
)

# ---- セッション初期化 ----
if "step" not in st.session_state:
    st.session_state.step = 1
if "authed" not in st.session_state:
    st.session_state.authed = False
if "extracted" not in st.session_state:
    st.session_state.extracted = None
if "affiliation" not in st.session_state:
    st.session_state.affiliation = ""
if "template_xlsx_bytes" not in st.session_state:
    st.session_state.template_xlsx_bytes = None

PASSCODE = _get_passcode()

# Step1: 認証
if st.session_state.step == 1:
    st.subheader("Step 1. パスコード認証")
    # Secrets未設定のときの注意喚起
    if not PASSCODE:
        st.info("（注意）現在、PASSCODEがSecrets/環境変数に未設定です。開発モード想定で空文字として扱います。")
    pw = st.text_input("パスコードを入力してください", type="password")
    if st.button("次へ", use_container_width=True):
        if pw == PASSCODE:
            st.session_state.authed = True
            st.session_state.step = 2
            st.rerun()
        else:
            st.error("パスコードが違います。")

# Step2: 入力
elif st.session_state.step == 2 and st.session_state.authed:
    st.subheader("Step 2. メール本文の貼り付け / 所属 / テンプレ選択")

    # --- テンプレ選択（既定ファイル or アップロード）
    template_path = "template.xlsm"
    tpl_col1, tpl_col2 = st.columns([0.55, 0.45])
    with tpl_col1:
        st.caption("① 既定：template.xlsm を探します")
        if os.path.exists(template_path) and not st.session_state.template_xlsx_bytes:
            try:
                with open(template_path, "rb") as f:
                    st.session_state.template_xlsx_bytes = f.read()
                st.success(f"テンプレートを読み込みました: {template_path}")
            except Exception as e:
                st.error(f"テンプレートの読み込みに失敗: {e}")
        elif st.session_state.template_xlsx_bytes:
            st.success("テンプレートは読み込み済みです。")
        else:
            st.warning("既定テンプレートが見つかりません。②のアップロードをご利用ください。")

    with tpl_col2:
        st.caption("② またはテンプレ.xlsmをアップロード")
        up = st.file_uploader("テンプレート（.xlsm）", type=["xlsm"], accept_multiple_files=False)
        if up is not None:
            st.session_state.template_xlsx_bytes = up.read()
            st.success(f"アップロード済み: {up.name}")

    # どちらも用意できない場合は処理停止
    if not st.session_state.template_xlsx_bytes:
        st.error("テンプレートが未準備です。template.xlsm を配置するか、上でアップロードしてください。")
        st.stop()

    # 所属
    aff = st.text_input("所属", value=st.session_state.affiliation)
    st.session_state.affiliation = aff

    # 任意の補足（処理修理後）
    processing_after = st.text_input("処理修理後（任意）")
    if processing_after:
        st.session_state["processing_after"] = processing_after

    # 本文
    text = st.text_area("故障完了メール（本文）を貼り付け", height=240, placeholder="ここにメール本文を貼り付け...")

    c1, c2 = st.columns(2)
    with c1:
        if st.button("抽出する", use_container_width=True):
            if not text.strip():
                st.warning("本文が空です。")
            else:
                st.session_state.extracted = extract_fields(text)
                st.session_state.extracted["所属"] = st.session_state.affiliation
                st.session_state.step = 3
                st.rerun()
    with c2:
        if st.button("クリア", use_container_width=True):
            st.session_state.extracted = None
            st.session_state.affiliation = ""
            st.session_state.processing_after = ""
            st.rerun()

# Step3: 抽出確認→Excel生成
elif st.session_state.step == 3 and st.session_state.authed:
    st.subheader("Step 3. 抽出結果の確認・編集 → Excel生成")

    # Step2の処理修理後を初回のみ反映
    if st.session_state.get("processing_after") and st.session_state.extracted is not None:
        if not st.session_state.extracted.get("_processing_after_initialized"):
            st.session_state.extracted["処理修理後"] = st.session_state["processing_after"]
            st.session_state.extracted["_processing_after_initialized"] = True

    data = st.session_state.extracted or {}

    # ① 基本情報
    with st.expander("① 基本情報", expanded=True):
        base_fields = ["管理番号", "物件名", "住所", "窓口会社"]
        for key in base_fields:
            val = data.get(key) or ""
            st.markdown(f"**{key}：** {val}")

    # ② 通報・受付情報
    with st.expander("② 通報・受付情報", expanded=True):
        st.markdown(f"**受信時刻：** {data.get('受信時刻') or ''}")
        editable_field("通報者", "通報者", 1)
        editable_field("受信内容", "受信内容", 4)

    # ③ 現着・作業・完了情報
    with st.expander("③ 現着・作業・完了情報", expanded=True):
        st.markdown(f"**現着時刻：** {data.get('現着時刻') or ''}")
        st.markdown(f"**完了時刻：** {data.get('完了時刻') or ''}")
        dur = data.get("作業時間_分")
        if dur:
            st.info(f"作業時間（概算）：{dur} 分")
        editable_field("現着状況", "現着状況", 5)
        editable_field("原因", "原因", 5)
        editable_field("処置内容", "処置内容", 5)
        editable_field("処理修理後（Step2入力値）", "処理修理後", 1)

    # ④ 技術情報
    with st.expander("④ 技術情報", expanded=False):
        tech_fields = ["制御方式", "契約種別", "メーカー"]
        for key in tech_fields:
            val = data.get(key) or ""
            st.markdown(f"**{key}：** {val}")

    # ⑤ その他情報
    with st.expander("⑤ その他情報", expanded=False):
        other_fields = ["所属", "対応者", "送信者", "受付番号", "受付URL", "現着完了登録URL"]
        for key in other_fields:
            val = data.get(key) or ""
            st.markdown(f"**{key}：** {val}")

    st.divider()

    # --- Excel出力 ---
    try:
        xlsx_bytes = fill_template_xlsx(st.session_state.template_xlsx_bytes, data)
        fname = build_filename(data)
        st.download_button(
            "Excelを生成（.xlsm）",
            data=xlsx_bytes,
            file_name=fname,
            mime="application/vnd.ms-excel.sheet.macroEnabled.12",
            use_container_width=True,
        )
    except Exception as e:
        st.error(f"テンプレート書き込み中にエラーが発生しました: {e}")
        with st.expander("詳細（開発者向け）"):
            st.code("".join(traceback.format_exception(*sys.exc_info())), language="python")

    # --- 戻るボタン群 ---
    c1, c2 = st.columns(2)
    with c1:
        if st.button("Step2に戻る", use_container_width=True):
            st.session_state.step = 2
            st.rerun()
    with c2:
        if st.button("最初に戻る", use_container_width=True):
            st.session_state.step = 1
            st.session_state.extracted = None
            st.session_state.affiliation = ""
            st.session_state.processing_after = ""
            st.rerun()

# 認証未完了時のフォールバック
else:
    st.warning("認証が必要です。Step1に戻ります。")
    st.session_state.step = 1
