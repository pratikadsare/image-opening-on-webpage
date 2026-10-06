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

The upload step is split into three tabs for the three PXM file shapes:
- Simple File: a plain PIS/masterfile-style sheet (the original tool).
- PXM Filename + URL Export File: one row per image, with a "Filename"
  like 50540_P07HR2C3_WM_US_ISP_04.jpg encoding the MPN and that image's
  sequence number, plus the real CDN URL in a separate column. This tab
  parses that, groups rows by MPN, and lays them out as SKU + Image 1..N.
- PXM Media Stack File: (to be built)

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


def build_two_tab_excel(base_data, data, sku_col, title_col, image_cols, swapped_rows):
    """A plain .xlsx with two tabs:
    - "Original": SKU, Title, Image URLs exactly as uploaded, no swaps.
    - "Reorder": just the SKU (Master ID) and Image URL columns, ONLY for
      the rows that actually had a swap queued, with that swap applied.
    """
    original_df = base_data[[sku_col, title_col] + image_cols].copy()
    original_df.columns = ["SKU", "Title"] + image_cols

    reorder_rows = sorted(swapped_rows)
    reorder_df = data.loc[reorder_rows, [sku_col] + image_cols].copy()
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


def parse_pxm_filename(filename):
    """Pull the MPN and the image sequence number out of a PXM filename
    like "50540_P07HR2C3_WM_US_ISP_04.jpg" -> ("P07HR2C3", 4).
    Returns (None, None) if the filename doesn't look like that shape."""
    if not isinstance(filename, str) or not filename.strip():
        return None, None
    base = filename.rsplit(".", 1)[0]
    parts = base.split("_")
    if len(parts) < 2:
        return None, None
    mpn = parts[1]
    try:
        seq = int(parts[-1])
    except ValueError:
        seq = None
    return mpn, seq


def build_pxm_pivot(df, filename_col, url_col):
    """Group PXM export rows by MPN and lay each MPN's images out in one
    row, ordered by the sequence number encoded in the filename."""
    groups = {}
    for idx, row in df.iterrows():
        fname, url = row[filename_col], row[url_col]
        if pd.isna(fname) or pd.isna(url) or not str(url).strip():
            continue
        mpn, seq = parse_pxm_filename(str(fname))
        if mpn is None:
            continue
        groups.setdefault(mpn, []).append({"seq": seq, "url": str(url).strip(), "orig_idx": idx})

    if not groups:
        return pd.DataFrame(columns=["SKU"]), []

    for mpn in groups:
        groups[mpn].sort(key=lambda r: (r["seq"] is None, r["seq"] if r["seq"] is not None else 0, r["orig_idx"]))

    max_images = max(len(v) for v in groups.values())
    image_cols = [f"Image {i + 1}" for i in range(max_images)]

    rows_out = []
    for mpn in sorted(groups.keys()):
        entry = {"SKU": mpn}
        for i, col_name in enumerate(image_cols):
            entry[col_name] = groups[mpn][i]["url"] if i < len(groups[mpn]) else ""
        rows_out.append(entry)

    pivot_df = pd.DataFrame(rows_out, columns=["SKU"] + image_cols)
    return pivot_df, image_cols


def build_pivot_table_html(df, image_cols, thumb_size):
    header_cells = "<th>SKU</th>" + "".join(f"<th>{col}</th>" for col in image_cols)
    body_rows = []
    for _, row in df.iterrows():
        cells = [f"<td>{row['SKU']}</td>"]
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


st.title(":frame_with_picture: Images Viewer")
st.caption("Upload a file, map your columns, and preview image URLs as thumbnails.")

tab_simple, tab_pxm_url, tab_pxm_media = st.tabs(
    ["Simple File", "PXM Filename + URL Export File", "PXM Media Stack File"]
)

def render_simple_tab():
    uploaded = st.file_uploader("Upload Excel file (.xlsx)", type=["xlsx"], key="simple_uploader")

    if uploaded is None:
        st.info("Once you upload a file, you'll pick the sheet, header row, and which columns are SKU / Title / Images.")
        return

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
        return

    st.subheader("1. Sheet / tab")
    sheet_name = st.selectbox("Which sheet has your data?", xls.sheet_names)

    raw = pd.read_excel(xls, sheet_name=sheet_name, header=None, dtype=str).fillna("")

    if raw.empty:
        st.warning("This sheet has no rows.")
        return

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
        st.dataframe(base_data.head(5), width="stretch")

    st.subheader("3. Map your columns")
    col1, col2, col3 = st.columns(3)
    with col1:
        sku_col = st.selectbox("SKU column", headers, index=0)
    with col2:
        title_col = st.selectbox("Title column", headers, index=min(1, len(headers) - 1))
    with col3:
        default_image_cols = [h for h in headers if h not in (sku_col, title_col)]
        image_cols = st.multiselect("Image URL column(s)", headers, default=default_image_cols)
        # Always keep image columns in the same left-to-right order as the
        # uploaded file, regardless of the order they were clicked/selected in.
        image_cols = sorted(image_cols, key=lambda c: headers.index(c))

    if not image_cols:
        st.warning("Pick at least one image column to see thumbnails.")
        return

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
                    if not is_url(val):
                        # Not an actual image URL for this row — show it, but
                        # it's not something you can swap, so no Select button.
                        st.caption(col)
                        st.write(val if val else "—")
                        continue
                    st.image(val, width=thumb_size, caption=col)
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
        swapped_rows = {s["row"] for s in st.session_state.swaps}
        two_tab_bytes = build_two_tab_excel(base_data, data, sku_col, title_col, image_cols, swapped_rows)
        st.download_button(
            label="Download as Excel (Original + Reorder tabs)",
            data=two_tab_bytes,
            file_name="images_viewer_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Tab 1 'Original': SKU, Title, Images exactly as uploaded. Tab 2 'Reorder': only the SKUs you swapped, SKU + Image links, with the swap applied.",
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


def render_pxm_url_tab():
    uploaded = st.file_uploader(
        "Upload PXM Filename + URL export file (.xlsx)", type=["xlsx"], key="pxm_url_uploader"
    )
    if uploaded is None:
        st.info(
            "Upload the PXM 'Filename + URL' export. This reads the Filename column "
            "(e.g. 50540_P07HR2C3_WM_US_ISP_04.jpg), pulls out the MPN and that image's "
            "sequence number, and lays each MPN's images out in one row, ordered by sequence."
        )
        return

    file_bytes = uploaded.getvalue()
    try:
        xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
    except Exception as exc:
        st.error(f"Could not read this file: {exc}")
        return

    sheet_name = st.selectbox("Which sheet has your data?", xls.sheet_names, key="pxm_url_sheet")
    raw = pd.read_excel(xls, sheet_name=sheet_name, dtype=str)
    if raw.empty:
        st.warning("This sheet has no rows.")
        return

    cols = raw.columns.tolist()
    default_filename_col = next((c for c in cols if c.strip().lower() == "filename"), cols[0])
    default_url_col = next(
        (c for c in cols if "image" in c.lower() and ("master file" in c.lower() or "url" in c.lower())),
        cols[min(1, len(cols) - 1)],
    )

    c1, c2 = st.columns(2)
    with c1:
        filename_col = st.selectbox(
            "Filename column", cols, index=cols.index(default_filename_col), key="pxm_url_filename_col"
        )
    with c2:
        url_col = st.selectbox(
            "Image URL column", cols, index=cols.index(default_url_col), key="pxm_url_url_col"
        )

    pivot_df, image_cols = build_pxm_pivot(raw, filename_col, url_col)

    if pivot_df.empty:
        st.warning("Couldn't find any rows with a parseable filename + URL. Check the columns picked above.")
        return

    image_row_count = raw[filename_col].notna().sum()
    st.success(f"Found {len(pivot_df)} SKU(s) across {image_row_count} image row(s).")

    st.subheader("Display options")
    thumb_size = st.slider(
        "Thumbnail size (px)", min_value=40, max_value=240, value=80, step=10, key="pxm_url_thumb_size"
    )

    st.subheader("Preview")
    table_html = build_pivot_table_html(pivot_df, image_cols, thumb_size)
    st.markdown(TABLE_CSS + table_html, unsafe_allow_html=True)

    st.subheader("Download")
    d1, d2 = st.columns(2)
    with d1:
        report_html = build_full_report(table_html)
        st.download_button(
            label="Download as HTML report",
            data=report_html.encode("utf-8"),
            file_name="pxm_images_report.html",
            mime="text/html",
        )
    with d2:
        buf = io.BytesIO()
        pivot_df.to_excel(buf, index=False, engine="openpyxl")
        st.download_button(
            label="Download as Excel (SKU + Images)",
            data=buf.getvalue(),
            file_name="pxm_images_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


with tab_simple:
    render_simple_tab()
with tab_pxm_url:
    render_pxm_url_tab()

with tab_pxm_media:
    st.info(
        "Coming soon. Share the PXM Media Stack file and how it should be sorted and displayed, "
        "and this tab will be built out to match."
    )
