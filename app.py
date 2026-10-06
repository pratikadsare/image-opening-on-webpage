"""
Images Viewer
--------------
Upload an Excel file, pick which sheet/tab to use, map which columns hold
the SKU, Title, and Image URL(s), and preview the image URLs as inline
thumbnails. Export the result as a standalone HTML report, or attach
per-row comments (e.g. "title is not matching with main image") and
download the SAME Excel file back with those comments written onto the
SKU cells, just like adding a comment manually in the masterfile.

Run with:
    pip install -r requirements.txt
    streamlit run app.py
"""

import io
import os

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Images Viewer", page_icon=":frame_with_picture:", layout="wide")

TABLE_CSS = """
<style>
table.iv-table { border-collapse: collapse; width: 100%; background: #fff; border-radius: 8px; overflow: hidden; }
table.iv-table th, table.iv-table td { border: 1px solid #ddd; padding: 10px; text-align: center; word-wrap: break-word; }
table.iv-table th { background: #e0e7ff; color: #1e40af; }
table.iv-table img { display: block; margin: auto; border-radius: 5px; }
</style>
"""

COMMENT_AUTHOR = "Images Viewer"


def is_url(value) -> bool:
    return isinstance(value, str) and value.strip().lower().startswith("http")


def build_table_html(df, sku_col, title_col, image_cols, thumb_size):
    header_cells = "<th>SKU</th><th>Title</th>" + "".join(f"<th>{col}</th>" for col in image_cols)
    body_rows = []
    for _, row in df.iterrows():
        cells = [f"<td>{row[sku_col]}</td>", f"<td>{row[title_col]}</td>"]
        for col in image_cols:
            val = str(row[col]) if pd.notna(row[col]) else ""
            if is_url(val):
                cells.append(
                    f'<td><img src="{val}" style="max-width:{thumb_size}px;max-height:{thumb_size}px;"></td>'
                )
            else:
                cells.append(f"<td>{val}</td>")
        body_rows.append(f"<tr>{''.join(cells)}</tr>")
    return f'<table class="iv-table"><thead><tr>{header_cells}</tr></thead><tbody>{"".join(body_rows)}</tbody></table>'


def build_full_report(table_html: str) -> str:
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Images Viewer Report</title>
{TABLE_CSS}
</head>
<body style="font-family: Arial, sans-serif; background:#f5f7fa; padding:30px;">
<h1 style="color:#1e3a8a;">Images Viewer Report</h1>
{table_html}
</body>
</html>"""


def build_commented_workbook(file_bytes, sheet_name, sku_col_pos, header_row_idx, comments_by_row):
    """Reopen the ORIGINAL uploaded file and write each comment onto the
    SKU cell of its row, preserving everything else in the workbook."""
    wb = load_workbook(io.BytesIO(file_bytes))
    ws = wb[sheet_name]
    col_letter = get_column_letter(sku_col_pos + 1)
    for row_index, note in comments_by_row.items():
        excel_row = header_row_idx + row_index + 2  # +1 for header row, +1 for 1-indexing
        cell = ws[f"{col_letter}{excel_row}"]
        cell.comment = Comment(note, COMMENT_AUTHOR)
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


st.title(":frame_with_picture: Images Viewer")
st.caption("Upload an Excel file, map your columns, and preview image URLs as thumbnails.")

uploaded = st.file_uploader("Upload Excel file (.xlsx)", type=["xlsx"])

if uploaded is None:
    st.info("Once you upload a file, you'll pick the sheet, header row, and which columns are SKU / Title / Images.")
    st.stop()

# Read the raw bytes once so we can both parse it with pandas AND, later,
# reopen the exact same original file with openpyxl to write comments back.
file_bytes = uploaded.getvalue()

if "sku_comments" not in st.session_state:
    st.session_state.sku_comments = {}
if "last_uploaded_name" not in st.session_state or st.session_state.last_uploaded_name != uploaded.name:
    st.session_state.sku_comments = {}
    st.session_state.last_uploaded_name = uploaded.name

try:
    xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
except Exception as exc:
    st.error(f"Could not read this file: {exc}")
    st.stop()

st.subheader("1. Sheet / tab")
sheet_name = st.selectbox("Which sheet has your data?", xls.sheet_names)

raw = pd.read_excel(xls, sheet_name=sheet_name, header=None, dtype=str).fillna("")

if raw.empty:
    st.warning("This sheet has no rows.")
    st.stop()

st.subheader("2. Header row")
header_row_display = st.number_input(
    "Which row holds the column headers? (1 = first row in the sheet)",
    min_value=1,
    max_value=len(raw),
    value=1,
    step=1,
)
header_row_idx = header_row_display - 1

headers = [
    str(h).strip() if str(h).strip() else f"Column {i + 1}"
    for i, h in enumerate(raw.iloc[header_row_idx].tolist())
]
data = raw.iloc[header_row_idx + 1 :].reset_index(drop=True)
data.columns = headers

with st.expander("Preview raw data (first 5 rows after the header)"):
    st.dataframe(data.head(5), use_container_width=True)

st.subheader("3. Map your columns")
col1, col2, col3 = st.columns(3)
with col1:
    sku_col = st.selectbox("SKU column", headers, index=0)
with col2:
    title_col = st.selectbox("Title column", headers, index=min(1, len(headers) - 1))
with col3:
    default_image_cols = [h for h in headers if h not in (sku_col, title_col)]
    image_cols = st.multiselect("Image URL column(s)", headers, default=default_image_cols)

if not image_cols:
    st.warning("Pick at least one image column to see thumbnails.")
    st.stop()

st.subheader("4. Display options")
thumb_size = st.slider("Thumbnail size (px)", min_value=40, max_value=240, value=80, step=10)

st.subheader("Preview")
table_html = build_table_html(data, sku_col, title_col, image_cols, thumb_size)
st.markdown(TABLE_CSS + table_html, unsafe_allow_html=True)

st.subheader("5. Add a comment for a specific SKU (optional)")
st.caption(
    "E.g. \"title is not matching with main image\" — this gets written as a real Excel "
    "comment on that SKU's cell, same as adding one by hand in the masterfile."
)

row_options = list(range(len(data)))
row_labels = {i: f"Row {i + 1} — SKU: {data.iloc[i][sku_col]}" for i in row_options}

c1, c2 = st.columns([1, 2])
with c1:
    selected_row = st.selectbox(
        "Row", row_options, format_func=lambda i: row_labels[i], key="comment_row_select"
    )
with c2:
    note_text = st.text_input("Comment", key="comment_text_input")

if st.button("Add comment", disabled=not note_text.strip()):
    st.session_state.sku_comments[selected_row] = note_text.strip()
    st.rerun()

if st.session_state.sku_comments:
    st.write("Comments added so far:")
    for row_index, note in list(st.session_state.sku_comments.items()):
        cc1, cc2, cc3 = st.columns([2, 4, 1])
        with cc1:
            st.write(row_labels.get(row_index, f"Row {row_index + 1}"))
        with cc2:
            st.write(note)
        with cc3:
            if st.button("Remove", key=f"remove_comment_{row_index}"):
                del st.session_state.sku_comments[row_index]
                st.rerun()

st.subheader("6. Download")

dl1, dl2 = st.columns(2)
with dl1:
    report_html = build_full_report(table_html)
    st.download_button(
        label="Download as HTML report",
        data=report_html.encode("utf-8"),
        file_name="images_viewer_report.html",
        mime="text/html",
    )
with dl2:
    sku_col_pos = headers.index(sku_col)
    commented_bytes = build_commented_workbook(
        file_bytes, sheet_name, sku_col_pos, header_row_idx, st.session_state.sku_comments
    )
    base, ext = os.path.splitext(uploaded.name)
    out_name = f"{base}_with_comments{ext or '.xlsx'}"
    st.download_button(
        label="Download same Excel file (with comments)",
        data=commented_bytes,
        file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        disabled=not st.session_state.sku_comments,
        help="Add at least one comment above to enable this.",
    )
