# app.py

# ————————————————————

# 故障メール → 正規表現抽出 → 既存テンプレ(.xlsm)へ書込み → ダウンロード

# 3ステップUI / パスコード認証 / 編集可能 / 折りたたみ表示（時系列）

# ————————————————————

import io
import re
import unicodedata
from datetime import datetime, timedelta, timezone
from typing import Dict, Optional, Tuple, List
import os
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
import streamlit as st

JST = timezone(timedelta(hours=9))

APP_TITLE = “故障報告メール → Excel自動生成（マクロ対応）”
PASSCODE_DEFAULT = “1357”
try:
PASSCODE = st.secrets.get(“APP_PASSCODE”, PASSCODE_DEFAULT)
except:
PASSCODE = PASSCODE_DEFAULT

SHEET_NAME = “緊急出動報告書（リンク付き）”
WEEKDAYS_JA = [“月”, “火”, “水”, “木”, “金”, “土”, “日”]

# ====== テキスト整形・抽出ユーティリティ ======

def normalize_text(text: str) -> str:
if not text:
return “”
t = unicodedata.normalize(“NFKC”, text)
t = t.replace(”：”, “:”)
t = t.replace(”\t”, “ “).replace(”\r\n”, “\n”).replace(”\r”, “\n”)
return t

def _search_one(pattern: str, text: str, flags=0) -> Optional[str]:
m = re.search(pattern, text, flags)
return m.group(1).strip() if m else None

def _search_span_between(labels: Dict[str, str], key: str, text: str) -> Optional[str]:
lab = labels[key]
others = [v for k, v in labels.items() if k != key]
boundary = “|”.join([f”(?:{v})” for v in others]) if others else r”$”
pattern = rf”{lab}\s*(.+?)(?=\n(?:{boundary})|\Z)”
m = re.search(pattern, text, flags=re.DOTALL | re.IGNORECASE)
return m.group(1).strip() if m else None

def _try_parse_datetime(s: Optional[str]) -> Optional[datetime]:
if not s:
return None
cand = s.strip()
cand = cand.replace(“年”, “/”).replace(“月”, “/”).replace(“日”, “”)
cand = cand.replace(”-”, “/”)
for fmt in (”%Y/%m/%d %H:%M:%S”, “%Y/%m/%d %H:%M”, “%Y/%m/%d”):
try:
return datetime.strptime(cand, fmt)
except Exception:
pass
return None

def _split_dt_components(dt: Optional[datetime]) -> Tuple[Optional[int], Optional[int], Optional[int], Optional[str], Optional[int], Optional[int]]:
if not dt:
return None, None, None, None, None, None
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
return dt.strftime(”%Y%m%d”)
return datetime.now().strftime(”%Y%m%d”)

def minutes_between(a: Optional[str], b: Optional[str]) -> Optional[int]:
s = _try_parse_datetime(a)
e = _try_parse_datetime(b)
if s and e:
return int((e - s).total_seconds() // 60)
return None

def _split_lines(text: Optional[str], max_lines: int = 5) -> List[str]:
if not text:
return []
lines = [ln.strip() for ln in text.splitlines() if ln.strip() != “”]
if len(lines) <= max_lines:
return lines
kept = lines[: max_lines - 1] + [lines[max_lines - 1] + “…”]
return kept

# ====== 正規表現 抽出 ======

def extract_fields(raw_text: str) -> Dict[str, Optional[str]]:
t = normalize_text(raw_text)
subject_case = _search_one(r”件名:\s*【\s*([^】]+)\s*】”, t, flags=re.IGNORECASE)
subject_manageno = _search_one(r”件名:.*?【[^】]+】\s*([A-Z0-9-]+)”, t, flags=re.IGNORECASE)

```
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

multiline_labels = {
    "受信内容": r"受信内容\s*:",
    "現着状況": r"現着状況\s*:",
    "原因": r"原因\s*:",
    "処置内容": r"処置内容\s*:",
    "通報者": r"通報者\s*:",
    "対応者": r"対応者\s*:",
    "送信者": r"送信者\s*:",
    "現着時刻": r"現着時刻\s*:",
    "完了時刻": r"完了時刻\s*:",
}

out = {k: None for k in single_line.keys() | multiline_labels.keys()}
out.update({
    "案件種別(件名)": subject_case,
    "受付URL": None,
    "現着完了登録URL": None,
})

for k, pat in single_line.items():
    out[k] = _search_one(pat, t, flags=re.IGNORECASE | re.MULTILINE)

if not out["管理番号"] and subject_manageno:
    out["管理番号"] = subject_manageno

for k in multiline_labels:
    out[k] = _search_span_between(multiline_labels, k, t)

dur = minutes_between(out["現着時刻"], out["完了時刻"])
out["作業時間_分"] = str(dur) if dur is not None and dur >= 0 else None
return out
```

# ====== テンプレ書き込み ======

def fill_template_xlsx(template_bytes: bytes, data: Dict[str, Optional[str]]) -> bytes:
wb = load_workbook(io.BytesIO(template_bytes), keep_vba=True)
ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.active

```
# --- 基本情報 ---
if data.get("管理番号"): ws["C12"] = data["管理番号"]
if data.get("メーカー"): ws["J12"] = data["メーカー"]
if data.get("制御方式"): ws["M12"] = data["制御方式"]
if data.get("受信内容"): ws["C15"] = data["受信内容"]
if data.get("通報者"): ws["C14"] = data["通報者"]
if data.get("対応者"): ws["L37"] = data["対応者"]

if data.get("処理修理後"):
    ws["C35"] = data["処理修理後"]

if data.get("所属"): 
    ws["C37"] = data["所属"]

now = datetime.now(JST)
ws["B5"], ws["D5"], ws["F5"] = now.year, now.month, now.day

# --- 日付・時刻ブロック ---
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

# --- 複数行ブロック ---
def fill_multiline(col_letter: str, start_row: int, text: Optional[str], max_lines: int = 5):
    lines = _split_lines(text, max_lines=max_lines)
    for i in range(max_lines):
        ws[f"{col_letter}{start_row + i}"] = ""
    for idx, line in enumerate(lines[:max_lines]):
        ws[f"{col_letter}{start_row + idx}"] = line

fill_multiline("C", 20, data.get("現着状況"))
fill_multiline("C", 25, data.get("原因"))
fill_multiline("C", 30, data.get("処置内容"))

out = io.BytesIO()
wb.save(out)
return out.getvalue()
```

def build_filename(data: Dict[str, Optional[str]]) -> str:
base_day = *first_date_yyyymmdd(data.get(“現着時刻”), data.get(“完了時刻”), data.get(“受信時刻”))
manageno = (data.get(“管理番号”) or “UNKNOWN”).replace(”/”, “*”)
bname = (data.get(“物件名”) or “”).strip().replace(”/”, “*”)
if bname:
return f”緊急出動報告書*{manageno}*{bname}*{base_day}.xlsm”
return f”緊急出動報告書_{manageno}_{base_day}.xlsm”

# ====== 編集フィールド共通関数 ======

def editable_field(label, key, max_lines=1):
“”“共通：編集可能なフィールド表示”””
data = st.session_state.extracted
edit_key = f”edit_{key}”

```
if edit_key not in st.session_state:
    st.session_state[edit_key] = False

if not st.session_state[edit_key]:
    # 表示モード
    value = data.get(key) or ""
    
    col1, col2, col3 = st.columns([0.85, 0.1, 0.05])
    with col1:
        if max_lines == 1:
            st.text_input(label, value=value, disabled=True, key=f"display_{key}")
        else:
            st.text_area(label, value=value, height=max_lines * 30, disabled=True, key=f"display_{key}")
    with col2:
        if st.button("✏️", key=f"btn_{key}", help=f"{label}を編集", use_container_width=True):
            st.session_state[edit_key] = True
            st.rerun()
else:
    # 編集モード
    value = data.get(key) or ""
    
    st.markdown(f"**✏️ {label} を編集中**")
    if max_lines == 1:
        new_val = st.text_input(f"{label}", value=value, key=f"in_{key}", label_visibility="collapsed")
    else:
        new_val = st.text_area(f"{label}", value=value, height=max_lines * 30, key=f"ta_{key}", label_visibility="collapsed")
    
    col1, col2 = st.columns(2)
    with col1:
        if st.button("💾 保存", key=f"save_{key}", use_container_width=True):
            st.session_state.extracted[key] = new_val
            st.session_state[edit_key] = False
            st.success(f"{label}を更新しました")
            st.rerun()
    with col2:
        if st.button("❌ キャンセル", key=f"cancel_{key}", use_container_width=True):
            st.session_state[edit_key] = False
            st.rerun()
```

# ====== Streamlit UI ======

st.set_page_config(
page_title=APP_TITLE,
layout=“centered”,
initial_sidebar_state=“collapsed”
)

# カスタムCSS

st.markdown(”””

<style>
    .main > div {
        padding-top: 2rem;
    }
    .stButton button {
        font-weight: 500;
    }
    h1 {
        padding-bottom: 1rem;
        border-bottom: 3px solid #1f77b4;
    }
    h2 {
        color: #1f77b4;
        margin-top: 2rem;
    }
    .step-indicator {
        text-align: center;
        padding: 1rem;
        background: linear-gradient(90deg, #e3f2fd 0%, #bbdefb 100%);
        border-radius: 10px;
        margin-bottom: 2rem;
        font-weight: 600;
        font-size: 1.1rem;
    }
</style>

“””, unsafe_allow_html=True)

st.title(“📋 “ + APP_TITLE)

# セッション状態の初期化

if “step” not in st.session_state:
st.session_state.step = 1
if “authed” not in st.session_state:
st.session_state.authed = False
if “extracted” not in st.session_state:
st.session_state.extracted = None
if “affiliation” not in st.session_state:
st.session_state.affiliation = “”
if “processing_after” not in st.session_state:
st.session_state.processing_after = “”

# ステップインジケーター

step_names = [“🔐 認証”, “📝 入力”, “✅ 確認・生成”]
current_step = st.session_state.step
progress_html = f”””

<div class="step-indicator">
    {"  →  ".join([f"<span style='color: {'#1f77b4' if i+1 == current_step else '#999'};'>{'<b>' if i+1 == current_step else ''}{step_names[i]}{'</b>' if i+1 == current_step else ''}</span>" for i in range(3)])}
</div>
"""
st.markdown(progress_html, unsafe_allow_html=True)

# Step 1: パスコード認証

if st.session_state.step == 1:
st.markdown(”### 🔐 Step 1. パスコード認証”)

```
with st.container():
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        pw = st.text_input("パスコードを入力", type="password", placeholder="パスコードを入力してください")
        
        if st.button("🔓 ログイン", use_container_width=True, type="primary"):
            if pw == PASSCODE:
                st.session_state.authed = True
                st.session_state.step = 2
                st.rerun()
            else:
                st.error("⚠️ パスコードが正しくありません")
```

elif st.session_state.step == 2 and st.session_state.authed:
# Step 2: メール本文＋テンプレ自動読み込み
st.markdown(”### 📝 Step 2. メール本文の貼り付け”)

```
# テンプレート読み込み
template_path = "template.xlsm"
if os.path.exists(template_path):
    with open(template_path, "rb") as f:
        st.session_state.template_xlsx_bytes = f.read()
    st.success(f"✅ テンプレートを読み込みました: `{template_path}`")
else:
    st.error(f"❌ テンプレートファイルが見つかりません: `{template_path}`")
    st.stop()

# 入力フォーム
with st.container():
    col1, col2 = st.columns(2)
    with col1:
        aff = st.text_input("🏢 所属", value=st.session_state.affiliation, placeholder="例：札幌支店 / 本社")
        st.session_state.affiliation = aff
    with col2:
        processing_after = st.text_input("🔧 処理修理後（任意）", value=st.session_state.processing_after, placeholder="任意項目")
        st.session_state.processing_after = processing_after

text = st.text_area(
    "📧 故障完了メール（本文）", 
    height=300,
    placeholder="メール本文をここに貼り付けてください...",
    help="完了メールの本文をそのまま貼り付けてください"
)

st.divider()

col1, col2, col3 = st.columns([2, 2, 1])
with col1:
    if st.button("🔍 抽出して次へ", use_container_width=True, type="primary"):
        if not text.strip():
            st.warning("⚠️ メール本文を入力してください")
        else:
            with st.spinner("データを抽出中..."):
                st.session_state.extracted = extract_fields(text)
                st.session_state.extracted["所属"] = st.session_state.affiliation
                if st.session_state.processing_after:
                    st.session_state.extracted["処理修理後"] = st.session_state.processing_after
                st.session_state.step = 3
            st.rerun()
with col2:
    if st.button("🗑️ クリア", use_container_width=True):
        st.session_state.extracted = None
        st.session_state.affiliation = ""
        st.session_state.processing_after = ""
        st.rerun()
with col3:
    if st.button("⬅️ 戻る", use_container_width=True):
        st.session_state.step = 1
        st.rerun()
```

elif st.session_state.step == 3 and st.session_state.authed:
# Step 3: 抽出結果の確認・編集 → Excel生成
st.markdown(”### ✅ Step 3. 抽出結果の確認・編集”)

```
data = st.session_state.extracted or {}

# 基本情報サマリー（折りたたみ不可）
with st.container():
    st.markdown("#### 📊 基本情報")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("管理番号", data.get("管理番号") or "未取得")
    with col2:
        st.metric("物件名", data.get("物件名") or "未取得")
    with col3:
        st.metric("メーカー", data.get("メーカー") or "未取得")

st.divider()

# 編集可能フィールド
with st.expander("📞 通報・受付情報", expanded=True):
    editable_field("通報者", "通報者", 1)
    editable_field("受信内容", "受信内容", 4)

with st.expander("🔧 現着・作業・完了情報", expanded=True):
    editable_field("現着状況", "現着状況", 5)
    editable_field("原因", "原因", 5)
    editable_field("処置内容", "処置内容", 5)
    editable_field("処理修理後", "処理修理後", 1)

st.divider()

# Excel生成
st.markdown("#### 📥 Excel出力")
try:
    xlsx_bytes = fill_template_xlsx(st.session_state.template_xlsx_bytes, data)
    fname = build_filename(data)
    
    col1, col2 = st.columns([3, 1])
    with col1:
        st.download_button(
            "📥 Excel ファイルをダウンロード (.xlsm)",
            data=xlsx_bytes,
            file_name=fname,
            mime="application/vnd.ms-excel.sheet.macroEnabled.12",
            use_container_width=True,
            type="primary"
        )
        st.caption(f"ファイル名: `{fname}`")
except Exception as e:
    st.error(f"❌ エラーが発生しました: {e}")

st.divider()

# ナビゲーション
col1, col2 = st.columns(2)
with col1:
    if st.button("⬅️ Step2に戻る", use_container_width=True):
        st.session_state.step = 2
        st.rerun()
with col2:
    if st.button("🔄 最初に戻る", use_container_width=True):
        st.session_state.step = 1
        st.session_state.extracted = None
        st.session_state.affiliation = ""
        st.session_state.processing_after = ""
        st.rerun()
```

else:
# 認証なし状態
st.warning(“⚠️ 認証が必要です”)
if st.button(“🔐 ログイン画面へ”):
st.session_state.step = 1
st.session_state.authed = False
st.rerun()