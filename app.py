"""
Company Name Fuzzy Matching — Streamlit App
=============================================
Upload either 1 Excel file (multiple tabs) OR 2 separate Excel files.
The app automatically detects which tab/file is "Indue" (the reference
list) based on the file or tab name containing "indue" (case-insensitive).
If it can't auto-detect, the user can choose manually.

Output: a preview table in the app + a downloadable Excel file with 3 tabs:
  - Result: Indue column, best-match column, match score
  - Source - <Indue name>: full original data
  - Source - <other name>: full original data
"""

import io
import re
import unicodedata

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from rapidfuzz import fuzz, process

# ============================================================================
# PAGE CONFIG
# ============================================================================
st.set_page_config(page_title="Company Name Matcher", page_icon="🔎", layout="wide")


# ============================================================================
# 1. NORMALISE COMPANY NAMES
# ============================================================================
def clean(name: str) -> str:
    """Normalise a company name: lowercase, strip accents, strip legal/status suffixes, strip punctuation."""
    if not isinstance(name, str):
        return ""
    s = name.lower()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = re.sub(
        r"\b(limited|ltd|llp|llc|plc|inc|incorporated|gmbh|srl|bv|nv|sa"
        r"|lp|corp|corporation|company|ag|kg|oy|ab|pty|pte"
        r"|in administration|in liquidation|in receivership|dissolved)\b",
        " ", s
    )
    s = re.sub(r"\bt\/as?\b.*$", " ", s)
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def combined_score(a_clean: str, b_clean: str) -> float:
    """Blend token_sort_ratio (handles reordered words) with ratio (penalises length mismatches)."""
    s1 = fuzz.token_sort_ratio(a_clean, b_clean)
    s2 = fuzz.ratio(a_clean, b_clean)
    return 0.7 * s1 + 0.3 * s2


# ============================================================================
# 2. DETECT COMPANY-NAME COLUMN & EXTRACT NAMES
# ============================================================================
def find_company_column(df: pd.DataFrame) -> int:
    headers = [str(c).strip().lower() for c in df.columns]
    for i, h in enumerate(headers):
        if "company" in h and "name" in h:
            return i
    for i, h in enumerate(headers):
        if "name" in h or "company" in h:
            return i
    for i, col in enumerate(df.columns):
        non_null = df[col].dropna()
        if len(non_null) > 0 and non_null.astype(str).str.match(r"^[A-Za-z]").mean() > 0.5:
            return i
    return 0


def extract_names(df_raw: pd.DataFrame) -> list[str]:
    df = df_raw.dropna(how="all")
    if df.empty:
        return []
    col_idx = find_company_column(df)
    col = df.iloc[:, col_idx].dropna().astype(str).str.strip()
    return [c for c in col.tolist() if c and c.lower() not in ("name", "company name", "nan")]


def is_indue_label(label: str) -> bool:
    return "indue" in label.lower()


# ============================================================================
# 3. MATCHING
# ============================================================================
def match_all(indue_names: list[str], other_names: list[str], threshold: float):
    other_cleaned = [clean(n) for n in other_names]
    valid_idx = [i for i, c in enumerate(other_cleaned) if c]
    valid_other_clean = [other_cleaned[i] for i in valid_idx]
    valid_other_orig = [other_names[i] for i in valid_idx]

    results = []
    for ind_orig in indue_names:
        ind_clean = clean(ind_orig)
        if not ind_clean or not valid_other_clean:
            results.append((ind_orig, "", 0.0))
            continue

        coarse = process.extract(
            ind_clean, valid_other_clean, scorer=fuzz.token_sort_ratio, limit=5
        )
        best_score, best_name = 0.0, ""
        for _, coarse_score, idx in coarse:
            if coarse_score < threshold - 8:
                continue
            sc = combined_score(ind_clean, valid_other_clean[idx])
            if sc > best_score:
                best_score, best_name = sc, valid_other_orig[idx]

        results.append((ind_orig, best_name if best_score >= threshold else "", round(best_score, 1)))

    return results


# ============================================================================
# 4. BUILD OUTPUT EXCEL FILE (returns bytes for Streamlit download)
# ============================================================================
HEADER_FILL = PatternFill("solid", fgColor="4472C4")
HEADER_FONT = Font(bold=True, color="FFFFFF", name="Arial", size=11)
ALT_FILL = PatternFill("solid", fgColor="DCE6F1")
DATA_FONT = Font(name="Arial", size=10)


def write_header(ws, headers):
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=col, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center")


def autosize_columns(ws, widths):
    for i, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = min(max(width, 12), 70)


def write_result_sheet(wb, results, indue_label, other_label):
    ws = wb.create_sheet("Result")
    write_header(ws, [indue_label, other_label, "Match score"])
    for row_idx, (ind, match, score) in enumerate(results, start=2):
        fill = ALT_FILL if row_idx % 2 == 0 else None
        for col, val in enumerate([ind, match, score if match else ""], 1):
            cell = ws.cell(row=row_idx, column=col, value=val)
            cell.font = DATA_FONT
            if fill:
                cell.fill = fill
    autosize_columns(ws, [55, 55, 14])
    ws.freeze_panes = "A2"


def write_source_sheet(wb, df_raw, sheet_name):
    ws = wb.create_sheet(sheet_name[:31])
    headers = [str(c) for c in df_raw.columns]
    write_header(ws, headers)
    for row_idx, (_, row) in enumerate(df_raw.iterrows(), start=2):
        fill = ALT_FILL if row_idx % 2 == 0 else None
        for col, val in enumerate(row.tolist(), 1):
            cell = ws.cell(row=row_idx, column=col, value=None if pd.isna(val) else val)
            cell.font = DATA_FONT
            if fill:
                cell.fill = fill
    autosize_columns(ws, [max(18, len(h) + 2) for h in headers])
    ws.freeze_panes = "A2"


def build_output_excel(results, indue_label, other_label, indue_raw, other_raw) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)
    write_result_sheet(wb, results, indue_label, other_label)
    write_source_sheet(wb, indue_raw, f"Source - {indue_label}")
    write_source_sheet(wb, other_raw, f"Source - {other_label}")
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ============================================================================
# 5. READ UPLOADED FILE → DICT {sheet_name: df}
# ============================================================================
@st.cache_data(show_spinner=False)
def read_excel_sheets(file_bytes: bytes) -> dict[str, pd.DataFrame]:
    return pd.read_excel(io.BytesIO(file_bytes), sheet_name=None, dtype=str)


# ============================================================================
# UI
# ============================================================================
st.title("🔎 Company Name Matcher")
st.write(
    "Match company names between an **Indue** (reference) list and another list using "
    "fuzzy matching. Supports uploading **1 Excel file with multiple tabs** or **2 separate Excel files**."
)

mode = st.radio(
    "Choose how to upload your data:",
    ["1 Excel file (multiple tabs)", "2 separate Excel files"],
    horizontal=True,
)

threshold = st.slider(
    "Similarity threshold to accept a match (0–100)",
    min_value=50, max_value=100, value=86, step=1,
    help="Scores below this threshold leave the match cell blank (treated as no match found)."
)

indue_raw, other_raw, indue_label, other_label = None, None, None, None

if mode == "1 Excel file (multiple tabs)":
    uploaded = st.file_uploader("Upload Excel file (.xlsx)", type=["xlsx", "xlsm"])
    if uploaded:
        sheets = read_excel_sheets(uploaded.getvalue())
        sheet_names = list(sheets.keys())
        if len(sheet_names) < 2:
            st.error("This file only has 1 tab. At least 2 tabs are needed to compare.")
        else:
            auto_indue = [n for n in sheet_names if is_indue_label(n)]
            default_indue = auto_indue[0] if auto_indue else sheet_names[0]
            col1, col2 = st.columns(2)
            with col1:
                indue_label = st.selectbox(
                    "Reference tab (Indue)", sheet_names,
                    index=sheet_names.index(default_indue),
                )
            with col2:
                remaining = [n for n in sheet_names if n != indue_label]
                other_label = st.selectbox("Comparison tab", remaining, index=0)
            indue_raw = sheets[indue_label]
            other_raw = sheets[other_label]
            if auto_indue:
                st.caption(f"✓ Automatically detected Indue tab: **{auto_indue[0]}**")
            else:
                st.caption("⚠️ Couldn't auto-detect the Indue tab from its name — please choose manually above.")

else:
    col1, col2 = st.columns(2)
    with col1:
        indue_file = st.file_uploader("Upload Indue file (reference list)", type=["xlsx", "xlsm"], key="indue")
    with col2:
        other_file = st.file_uploader("Upload the other file (to compare)", type=["xlsx", "xlsm"], key="other")

    if indue_file and other_file:
        indue_sheets = read_excel_sheets(indue_file.getvalue())
        other_sheets = read_excel_sheets(other_file.getvalue())
        # If a file has multiple tabs, prefer the first one or one containing "indue"
        def pick_sheet(sheets, prefer_indue):
            names = list(sheets.keys())
            if len(names) == 1:
                return names[0], sheets[names[0]]
            if prefer_indue:
                hits = [n for n in names if is_indue_label(n)]
                chosen = hits[0] if hits else names[0]
            else:
                chosen = names[0]
            return chosen, sheets[chosen]

        indue_sheet_name, indue_raw = pick_sheet(indue_sheets, prefer_indue=True)
        other_sheet_name, other_raw = pick_sheet(other_sheets, prefer_indue=False)
        indue_label = indue_file.name.rsplit(".", 1)[0]
        other_label = other_file.name.rsplit(".", 1)[0]
        st.caption(f"✓ Indue file: **{indue_file.name}** (tab: {indue_sheet_name})")
        st.caption(f"✓ Comparison file: **{other_file.name}** (tab: {other_sheet_name})")

# ============================================================================
# RUN MATCHING
# ============================================================================
if indue_raw is not None and other_raw is not None:
    indue_names = extract_names(indue_raw)
    other_names = extract_names(other_raw)

    st.write(f"📋 **{indue_label}**: {len(indue_names)} companies &nbsp;|&nbsp; **{other_label}**: {len(other_names)} companies")

    if st.button("🚀 Run matching", type="primary"):
        if not indue_names or not other_names:
            st.error("No company names found in one of the two data sources. Please check your file(s).")
        else:
            with st.spinner("Matching..."):
                results = match_all(indue_names, other_names, threshold)

            matched = sum(1 for _, k, _ in results if k)
            st.success(f"Done! Matched {matched} / {len(indue_names)} Indue companies.")

            result_df = pd.DataFrame(results, columns=[indue_label, other_label, "Match score"])
            result_df["Match score"] = result_df.apply(
                lambda r: r["Match score"] if r[other_label] else None, axis=1
            )
            st.dataframe(result_df, use_container_width=True, height=400)

            output_bytes = build_output_excel(results, indue_label, other_label, indue_raw, other_raw)
            st.download_button(
                label="⬇️ Download result file (.xlsx)",
                data=output_bytes,
                file_name="Matching result - OUTPUT.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            with st.expander("View original data"):
                tab1, tab2 = st.tabs([f"Source - {indue_label}", f"Source - {other_label}"])
                with tab1:
                    st.dataframe(indue_raw, use_container_width=True)
                with tab2:
                    st.dataframe(other_raw, use_container_width=True)
else:
    st.info("👆 Upload a file to get started.")
