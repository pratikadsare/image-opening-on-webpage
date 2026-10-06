"""
Images Viewer
--------------
Upload an Excel file, pick which sheet/tab to use, map which columns hold
the SKU, Title, and Image URL(s), and preview the image URLs as inline
thumbnails. Export the result as a standalone HTML report.

Also handles the "image is in the wrong column" fix: turn on Swap mode,
click one thumbnail then another in the same row, and their link values
trade places. Queue up several of these, then download the same Excel
file back with those links corrected (no more copy-into-a-scratch-cell).

Run with:
    pip install -r requirements.txt
    streamlit run app.py
"""

import io
import os

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

st.set_page_config(page_title="Images Viewer", page_icon=":frame_with_picture:", layout="wide")

TABLE_CSS = """
<style>
table.iv-table { border-collapse: collapse; width: 100%; background: #fff; border-radius: 8px; overflow: hidden; }
table.iv-table th, table.iv-table td { border: 1px solid #ddd; padding: 10px; text-align: center; word-wrap: break-word; }
table.iv-table th { background: #e0e7ff; color: #1e40af; }
table.iv-table img { display: block; margin: auto; border-radius: 5px; }
</style>
"""


def is_url(value) -> bool:
    return isinstance(value, str) and value.strip().lower().startswith("http")


def apply_swaps(df, swaps):
    """Replay the queued swaps on a fresh copy of the data so the preview,
    HTML report, and Excel export all reflect the current, corrected state."""
    df = df.copy()
    for s in swaps:
        row, col_a, col_b = s["row"], s["col_a"], s["col_b"]
        val_a = df.at[row, col_a]
        val_b = df.at[row, col_b]
        df.at[row, col_a] = val_b
        df.at[row, col_b] = val_a
    return df


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


def build_two_tab_excel(base_data, data, sku_col, title_col, image_cols):
    """A plain .xlsx with two tabs:
    - "Original": SKU, Title, Image URLs exactly as uploaded, no swaps.
    - "Reorder": just the SKU (Master ID) and Image URL columns, with any
      queued swaps applied.
    """
    original_df = base_data[[sku_col, title_col] + image_cols].copy()
    original_df.columns = ["SKU", "Title"] + image_cols

    reorder_df = data[[sku_col] + image_cols].copy()
    reorder_df.columns = ["SKU"] + image_cols

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        original_df.to_excel(writer, sheet_name="Original", index=False)
        reorder_df.to_excel(writer, sheet_name="Reorder", index=False)
    return buf.getvalue()


def build_swapped_workbook(file_bytes, sheet_name, headers, header_row_idx, swaps):
    """Reopen the ORIGINAL uploaded file and swap the two cell values for
    each queued swap, preserving everything else in the workbook."""
    wb = load_workbook(io.BytesIO(file_bytes))
    ws = wb[sheet_name]
    for s in swaps:
        row, col_a, col_b = s["row"], s["col_a"], s["col_b"]
        col_a_pos = headers.index(col_a)
        col_b_pos = headers.index(col_b)
        excel_row = header_row_idx + row + 2  # +1 for header row, +1 for 1-indexing
        cell_a = ws.cell(row=excel_row, column=col_a_pos + 1)
        cell_b = ws.cell(row=excel_row, column=col_b_pos + 1)
        cell_a.value, cell_b.value = cell_b.value, cell_a.value
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


st.title(":frame_with_picture: Images Viewer")
st.caption("Upload an Excel file, map your columns, and preview image URLs as thumbnails.")

uploaded = st.file_uploader("Upload Excel file (.xlsx)", type=["xlsx"])

if uploaded is None:
    st.info("Once you upload a file, you'll pick the sheet, header row, and which columns are SKU / Title / Images.")
    st.stop()

file_bytes = uploaded.getvalue()

if st.session_state.get("last_uploaded_name") != uploaded.name:
    st.session_state.swaps = []
    st.session_state.swap_pending = None
    st.session_state.last_uploaded_name = uploaded.name
st.session_state.setdefault("swaps", [])
st.session_state.setdefault("swap_pending", None)

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
base_data = raw.iloc[header_row_idx + 1 :].reset_index(drop=True)
base_data.columns = headers

with st.expander("Preview raw data (first 5 rows after the header)"):
    st.dataframe(base_data.head(5), use_container_width=True)

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

# Re-apply every queued swap to a fresh copy of the data so the preview,
# report, and export all stay in sync no matter what else reran above.
data = apply_swaps(base_data, st.session_state.swaps)

st.subheader("Preview")
swap_mode = st.toggle("Swap", value=False, help="Click one thumbnail, then another in the same row, to swap their links.")

if swap_mode:
    pending = st.session_state.swap_pending
    if pending is not None:
        st.info(f"Selected Row {pending[0] + 1} — {pending[1]}. Click another image in the SAME row to swap with it.")
    else:
        st.caption("Click a thumbnail to start a swap, then click a second thumbnail in the same row.")

    for row_idx in range(len(data)):
        row = data.iloc[row_idx]
        row_cols = st.columns([1, 2] + [1] * len(image_cols))
        with row_cols[0]:
            st.write(f"**{row[sku_col]}**")
        with row_cols[1]:
            st.write(row[title_col])
        for i, col in enumerate(image_cols):
            with row_cols[2 + i]:
                val = str(row[col]) if pd.notna(row[col]) else ""
                if is_url(val):
                    st.image(val, width=thumb_size, caption=col)
                else:
                    st.caption(col)
                    st.write(val if val else "—")
                selection = (row_idx, col)
                is_selected = pending == selection
                if st.button(
                    "Selected" if is_selected else "Select",
                    key=f"swap_btn_{row_idx}_{col}",
                    type="primary" if is_selected else "secondary",
                ):
                    if pending is None:
                        st.session_state.swap_pending = selection
                    elif pending == selection:
                        st.session_state.swap_pending = None
                    elif pending[0] != row_idx:
                        st.warning("Pick two images from the same row.")
                        st.session_state.swap_pending = selection
                    else:
                        st.session_state.swaps.append(
                            {"row": row_idx, "col_a": pending[1], "col_b": col}
                        )
                        st.session_state.swap_pending = None
                    st.rerun()

    if st.session_state.swaps:
        st.write("Swaps queued:")
        for i, s in enumerate(st.session_state.swaps):
            lc1, lc2, lc3 = st.columns([3, 3, 1])
            with lc1:
                st.write(f"Row {s['row'] + 1} — {data.iloc[s['row']][sku_col]}")
            with lc2:
                st.write(f"{s['col_a']}  ⇄  {s['col_b']}")
            with lc3:
                if st.button("Undo", key=f"undo_swap_{i}"):
                    st.session_state.swaps.pop(i)
                    st.rerun()
else:
    table_html = build_table_html(data, sku_col, title_col, image_cols, thumb_size)
    st.markdown(TABLE_CSS + table_html, unsafe_allow_html=True)

st.subheader("5. Download")
table_html = build_table_html(data, sku_col, title_col, image_cols, thumb_size)

dl1, dl2, dl3 = st.columns(3)
with dl1:
    report_html = build_full_report(table_html)
    st.download_button(
        label="Download as HTML report",
        data=report_html.encode("utf-8"),
        file_name="images_viewer_report.html",
        mime="text/html",
    )
with dl2:
    two_tab_bytes = build_two_tab_excel(base_data, data, sku_col, title_col, image_cols)
    st.download_button(
        label="Download as Excel (Original + Reorder tabs)",
        data=two_tab_bytes,
        file_name="images_viewer_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="Tab 1 'Original': SKU, Title, Images exactly as uploaded. Tab 2 'Reorder': SKU + Image links only, with swaps applied.",
    )
with dl3:
    swapped_bytes = build_swapped_workbook(file_bytes, sheet_name, headers, header_row_idx, st.session_state.swaps)
    base, ext = os.path.splitext(uploaded.name)
    out_name = f"{base}_corrected{ext or '.xlsx'}"
    st.download_button(
        label="Download same Excel file (with swaps applied)",
        data=swapped_bytes,
        file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        disabled=not st.session_state.swaps,
        help="Queue at least one swap above to enable this.",
    )
