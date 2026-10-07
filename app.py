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

import html as html_lib
import io
import json
import os
from urllib.parse import urlparse

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from openpyxl import load_workbook

st.set_page_config(page_title="Images Viewer", page_icon="\U0001F5BC\ufe0f", layout="wide")

# ---------------------------------------------------------------------------
# Theme: light-blue by default, with a Day/Night toggle at the top.
#
# The base look (including native widgets like selects/sliders/buttons) comes
# from .streamlit/config.toml (a light-blue theme). The Day/Night toggle below
# layers CSS on top of that, targeting Streamlit's actual data-testid
# attributes (confirmed by inspecting the rendered app) so the whole page -
# tables, buttons, alerts, inputs - switches together.
# ---------------------------------------------------------------------------

THEME_PALETTES = {
    "light": {
        "app_bg": "#eaf4fd",
        "block_bg": "#f5faff",
        "header_bg": "rgba(234, 244, 253, 0.85)",
        "text": "#102a43",
        "muted_text": "#45617a",
        "border": "#bcdcf7",
        "accent": "#2563eb",
        "accent_text": "#ffffff",
        "table_bg": "#ffffff",
        "table_header_bg": "#dceafb",
        "table_header_text": "#1e40af",
        "table_border": "#cfe3f7",
        "input_bg": "#ffffff",
        "alert_bg": "#f5faff",
    },
    "dark": {
        "app_bg": "#0b1d2e",
        "block_bg": "#122a40",
        "header_bg": "rgba(11, 29, 46, 0.9)",
        "text": "#e6f1fb",
        "muted_text": "#9fb8cc",
        "border": "#1f3b54",
        "accent": "#60a5fa",
        "accent_text": "#0b1d2e",
        "table_bg": "#122a40",
        "table_header_bg": "#1b3a56",
        "table_header_text": "#bfe0ff",
        "table_border": "#1f3b54",
        "input_bg": "#162f47",
        "alert_bg": "#122a40",
    },
}


def get_theme_mode() -> str:
    return st.session_state.get("theme_mode", "light")


def get_table_css(mode: str = "light") -> str:
    p = THEME_PALETTES.get(mode, THEME_PALETTES["light"])
    return f"""
<style>
table.iv-table {{ border-collapse: collapse; width: 100%; background: {p['table_bg']}; border-radius: 8px; overflow: hidden; color: {p['text']}; }}
table.iv-table th, table.iv-table td {{ border: 1px solid {p['table_border']}; padding: 10px; text-align: center; word-wrap: break-word; }}
table.iv-table th {{ background: {p['table_header_bg']}; color: {p['table_header_text']}; }}
table.iv-table img {{ display: block; margin: auto; border-radius: 5px; cursor: zoom-in; }}
</style>
"""


# Kept for any lingering reference to the old constant name (defaults to light).
TABLE_CSS = get_table_css("light")


def inject_app_theme_css(mode: str) -> None:
    """Re-skin Streamlit's own chrome (background, header, buttons, inputs,
    alerts, tabs) to match the active Day/Night mode. Selectors below are the
    actual data-testid attributes Streamlit 1.65 renders."""
    p = THEME_PALETTES.get(mode, THEME_PALETTES["light"])
    st.markdown(
        f"""
<style>
[data-testid="stApp"], [data-testid="stAppViewContainer"], [data-testid="stMain"] {{
    background-color: {p['app_bg']} !important;
    color: {p['text']} !important;
}}
[data-testid="stHeader"] {{
    background-color: {p['header_bg']} !important;
}}
[data-testid="stMainBlockContainer"] {{
    background-color: transparent !important;
}}
[data-testid="stMarkdownContainer"], [data-testid="stCaptionContainer"],
[data-testid="stHeading"], [data-testid="stWidgetLabel"], label, p, span, div {{
    color: {p['text']};
}}
[data-testid="stCaptionContainer"] {{
    color: {p['muted_text']} !important;
}}
[data-testid="stTabs"] [data-testid="stTab"] {{
    color: {p['muted_text']};
}}
[data-testid="stTabs"] [aria-selected="true"] {{
    color: {p['accent']} !important;
    border-bottom-color: {p['accent']} !important;
}}
[data-testid="stExpander"], [data-testid="stFileUploaderDropzone"],
[data-testid="stDataFrame"], [data-testid="stAlertContainer"] {{
    background-color: {p['block_bg']} !important;
    border-color: {p['border']} !important;
    color: {p['text']} !important;
}}
[data-testid="stSelectbox"] div[role="group"],
[data-testid="stMultiSelect"] div[role="group"],
[data-testid="stSelectbox"] input,
[data-testid="stMultiSelect"] input,
[data-testid="stNumberInput"] input,
[data-testid="stNumberInputField"] {{
    background-color: {p['input_bg']} !important;
    color: {p['text']} !important;
    border-color: {p['border']} !important;
}}
[data-testid="stSelectbox"] button svg,
[data-testid="stMultiSelect"] button svg,
[data-testid="stNumberInputStepDown"] svg,
[data-testid="stNumberInputStepUp"] svg {{
    fill: {p['text']} !important;
}}
[data-testid="stSelectboxVirtualDropdown"],
[data-testid="stSelectboxVirtualDropdown"] div {{
    background-color: {p['block_bg']} !important;
    color: {p['text']} !important;
}}
[data-testid="stSelectboxVirtualDropdown"] [role="option"][aria-selected="true"],
[data-testid="stSelectboxVirtualDropdown"] [role="option"]:hover {{
    background-color: {p['accent']} !important;
    color: {p['accent_text']} !important;
}}
[data-testid="stSelectboxVirtualDropdown"] [role="option"][aria-selected="true"] div,
[data-testid="stSelectboxVirtualDropdown"] [role="option"]:hover div {{
    background-color: transparent !important;
    color: {p['accent_text']} !important;
}}
[data-testid="stBaseButton-secondary"] {{
    background-color: {p['block_bg']} !important;
    color: {p['text']} !important;
    border-color: {p['accent']} !important;
}}
[data-testid="stBaseButton-primary"], [data-testid="stDownloadButton"] button {{
    background-color: {p['accent']} !important;
    color: {p['accent_text']} !important;
}}
[data-testid="stSliderTickBar"] {{
    color: {p['muted_text']} !important;
}}
[data-testid="stFileUploaderDropzoneInstructions"] span,
[data-testid="stFileUploaderDropzoneInstructions"] small,
[data-testid="stWidgetLabel"] span,
small, .small-text {{
    color: {p['muted_text']} !important;
    opacity: 1 !important;
}}
hr {{
    border-color: {p['border']} !important;
}}
</style>
""",
        unsafe_allow_html=True,
    )


def render_theme_toggle() -> None:
    """A Day/Night button pinned at the top of the app."""
    st.session_state.setdefault("theme_mode", "light")
    inject_app_theme_css(st.session_state["theme_mode"])

    left, right = st.columns([5, 1])
    with left:
        st.title("\U0001F5BC\ufe0f Images Viewer")
    with right:
        is_dark = st.session_state["theme_mode"] == "dark"
        label = "☀️ Day Mode" if is_dark else "🌙 Night Mode"
        if st.button(label, use_container_width=True, key="theme_toggle_btn"):
            st.session_state["theme_mode"] = "light" if is_dark else "dark"
            st.rerun()


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



# ---------------------------------------------------------------------------
# Full-screen image viewer (lightbox)
#
# Clicking a thumbnail opens it full screen. Previous / Next move through the
# images in the SAME table row (same product / Master ID). The top bar shows
# the file name, Master ID / MPN (or SKU), and the image's real dimensions.
# The same script is embedded in the downloaded HTML report.
# ---------------------------------------------------------------------------

LIGHTBOX_JS = r"""
<script>
(function () {
  var host;
  try {
    host = (window.parent && window.parent !== window && window.parent.document) ? window.parent : window;
  } catch (e) { host = window; }
  var doc = host.document;

  if (host.__ivlbCleanup) { try { host.__ivlbCleanup(); } catch (e) {} }
  ["ivlb-overlay", "ivlb-style"].forEach(function (id) {
    var old = doc.getElementById(id); if (old) old.remove();
  });

  var style = doc.createElement("style");
  style.id = "ivlb-style";
  style.textContent = [
    "#ivlb-overlay{position:fixed;inset:0;z-index:2147483000;display:none;flex-direction:column;background:rgba(6,14,24,.94);color:#e6f1fb;font-family:'Source Sans Pro',Arial,sans-serif}",
    "#ivlb-overlay.open{display:flex}",
    "#ivlb-overlay .ivlb-top{display:flex;align-items:center;justify-content:space-between;gap:16px;padding:14px 26px;border-bottom:1px solid #1f3b54;background:rgba(18,42,64,.75)}",
    "#ivlb-overlay .ivlb-file{font-size:20px;font-weight:700;word-break:break-all;color:#ffffff !important}",
    "#ivlb-overlay .ivlb-chips{display:flex;gap:10px;flex-wrap:wrap;margin-top:6px}",
    "#ivlb-overlay .ivlb-chip{background:#1b3a56;border:1px solid #2b5075;border-radius:999px;padding:3px 12px;font-size:14px;color:#bfe0ff}",
    "#ivlb-overlay .ivlb-chip b{color:#fff}",
    "#ivlb-overlay .ivlb-right{display:flex;align-items:center;gap:16px;flex-shrink:0}",
    "#ivlb-overlay .ivlb-count{font-size:15px;color:#9fb8cc !important}",
    "#ivlb-overlay .ivlb-close{width:40px;height:40px;border-radius:50%;background:#1b3a56;border:1px solid #2b5075;color:#fff;font-size:24px;cursor:pointer;line-height:1}",
    "#ivlb-overlay .ivlb-stage{flex:1;min-height:0;display:flex;align-items:center;justify-content:center;position:relative;padding:20px 100px}",
    "#ivlb-overlay .ivlb-stage img{max-width:100%;max-height:100%;object-fit:contain;border-radius:6px;background:#fff;box-shadow:0 10px 40px rgba(0,0,0,.6)}",
    "#ivlb-overlay .ivlb-nav{position:absolute;top:50%;transform:translateY(-50%);width:58px;height:58px;border-radius:50%;border:none;background:#2563eb;color:#fff;font-size:34px;cursor:pointer;line-height:1;box-shadow:0 4px 14px rgba(0,0,0,.5)}",
    "#ivlb-overlay .ivlb-nav:hover{background:#3b7bff}",
    "#ivlb-overlay .ivlb-prev{left:26px}",
    "#ivlb-overlay .ivlb-next{right:26px}",
    "#ivlb-overlay .ivlb-hint{text-align:center;padding:8px;color:#7f9bb3 !important;font-size:13px}"
  ].join("\n");
  doc.head.appendChild(style);

  var ov = doc.createElement("div");
  ov.id = "ivlb-overlay";
  ov.innerHTML =
    '<div class="ivlb-top"><div><div class="ivlb-file"></div><div class="ivlb-chips"></div></div>' +
    '<div class="ivlb-right"><div class="ivlb-count"></div><button class="ivlb-close" title="Close (Esc)">&times;</button></div></div>' +
    '<div class="ivlb-stage"><button class="ivlb-nav ivlb-prev" title="Previous">&#8249;</button>' +
    '<img alt=""><button class="ivlb-nav ivlb-next" title="Next">&#8250;</button></div>' +
    '<div class="ivlb-hint">Use the arrow buttons or the left and right keys. Press Esc to close.</div>';
  doc.body.appendChild(ov);

  var big = ov.querySelector(".ivlb-stage img");
  var items = [], chips = [], idx = 0;

  function esc(s) {
    return String(s).replace(/[&<>"]/g, function (c) { return {"&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;"}[c]; });
  }

  function render() {
    var it = items[idx];
    ov.querySelector(".ivlb-file").textContent = it.file || it.label || "";
    var html = chips.map(function (c) { return '<span class="ivlb-chip">' + esc(c[0]) + ": <b>" + esc(c[1]) + "</b></span>"; }).join("");
    html += '<span class="ivlb-chip">Dimensions: <b class="ivlb-dim">loading...</b></span>';
    ov.querySelector(".ivlb-chips").innerHTML = html;
    ov.querySelector(".ivlb-count").textContent = (idx + 1) + " / " + items.length;
    big.onload = function () {
      var d = ov.querySelector(".ivlb-dim");
      if (d) d.textContent = big.naturalWidth + " x " + big.naturalHeight + " px";
    };
    big.onerror = function () {
      var d = ov.querySelector(".ivlb-dim");
      if (d) d.textContent = "unavailable";
    };
    big.src = it.url;
  }

  function openFrom(img) {
    var tr = img.closest("tr");
    var imgs = Array.prototype.slice.call(tr.querySelectorAll("img[data-lb]"));
    items = imgs.map(function (i) {
      return {url: i.getAttribute("src"), file: i.getAttribute("data-file"), label: i.getAttribute("data-label")};
    });
    try { chips = JSON.parse(tr.getAttribute("data-chips") || "[]"); } catch (e) { chips = []; }
    idx = Math.max(0, imgs.indexOf(img));
    ov.classList.add("open");
    render();
  }
  function close() { ov.classList.remove("open"); big.removeAttribute("src"); }
  function step(d) { if (items.length) { idx = (idx + d + items.length) % items.length; render(); } }

  function onKey(e) {
    if (!ov.classList.contains("open")) return;
    if (e.key === "ArrowRight") { e.preventDefault(); step(1); }
    else if (e.key === "ArrowLeft") { e.preventDefault(); step(-1); }
    else if (e.key === "Escape") { e.preventDefault(); close(); }
  }

  ov.querySelector(".ivlb-prev").addEventListener("click", function (e) { e.stopPropagation(); step(-1); });
  ov.querySelector(".ivlb-next").addEventListener("click", function (e) { e.stopPropagation(); step(1); });
  ov.querySelector(".ivlb-close").addEventListener("click", close);
  ov.addEventListener("click", function (e) { if (e.target === ov || e.target.classList.contains("ivlb-stage")) close(); });
  doc.addEventListener("keydown", onKey, true);
  document.addEventListener("keydown", onKey, true);
  document.addEventListener("click", function (e) {
    var t = e.target;
    if (t && t.tagName === "IMG" && t.hasAttribute("data-lb")) openFrom(t);
  });

  host.__ivlbCleanup = function () {
    doc.removeEventListener("keydown", onKey, true);
    var o = doc.getElementById("ivlb-overlay"); if (o) o.remove();
  };
})();
</script>
"""


def url_basename(url: str) -> str:
    try:
        return os.path.basename(urlparse(str(url)).path) or ""
    except ValueError:
        return ""


def build_url_filename_map(df, filename_col, url_col):
    """url -> real filename, so the viewer can show the file name even though
    the pivoted tables only keep the links. Follows the link when swapped."""
    names = {}
    for fname, url in zip(df[filename_col], df[url_col]):
        if pd.notna(fname) and pd.notna(url) and str(url).strip():
            names.setdefault(str(url).strip(), str(fname).strip())
    return names


def img_cell(url, thumb_size, file_name, label):
    a = html_lib.escape
    return (
        f'<td><img data-lb="1" src="{a(url, quote=True)}" data-file="{a(file_name, quote=True)}" '
        f'data-label="{a(label, quote=True)}" '
        f'style="max-width:{thumb_size}px;max-height:{thumb_size}px;"></td>'
    )


def row_open_tag(chips):
    data = html_lib.escape(json.dumps([[str(k), str(v)] for k, v in chips]), quote=True)
    return f'<tr data-chips="{data}">'


def show_table(table_html, n_rows, thumb_size):
    """Render a preview table in an iframe so the full-screen viewer's
    JavaScript can run (st.markdown strips scripts)."""
    p = THEME_PALETTES.get(get_theme_mode(), THEME_PALETTES["light"])
    page = (
        get_table_css(get_theme_mode())
        + f'<body style="margin:0;font-family:Arial,sans-serif;color:{p["text"]};background:transparent;">'
        + table_html
        + LIGHTBOX_JS
        + "</body>"
    )
    height = int(min(780, max(180, 70 + n_rows * (thumb_size + 30))))
    if hasattr(st, "iframe"):
        st.iframe(page, height=height)
    else:  # older Streamlit versions
        components.html(page, height=height, scrolling=True)


def build_table_html(df, mpn_col, master_col, title_col, image_cols, thumb_size):
    header_cells = "<th>MPN</th><th>Master ID</th><th>Title</th>" + "".join(f"<th>{col}</th>" for col in image_cols)
    body_rows = []
    for _, row in df.iterrows():
        title_text = str(row[title_col])
        chips = [("Master ID", row[master_col]), ("MPN", row[mpn_col]), ("Title", title_text[:80])]
        cells = [f"<td>{row[mpn_col]}</td>", f"<td>{row[master_col]}</td>", f"<td>{row[title_col]}</td>"]
        for col in image_cols:
            val = str(row[col]) if pd.notna(row[col]) else ""
            if is_url(val):
                cells.append(img_cell(val.strip(), thumb_size, url_basename(val), col))
            else:
                cells.append(f"<td>{val}</td>")
        body_rows.append(f"{row_open_tag(chips)}{''.join(cells)}</tr>")
    return f'<table class="iv-table"><thead><tr>{header_cells}</tr></thead><tbody>{"".join(body_rows)}</tbody></table>'


def build_full_report(table_html: str, mode: str = "light") -> str:
    p = THEME_PALETTES.get(mode, THEME_PALETTES["light"])
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Images Viewer Report</title>
{get_table_css(mode)}
</head>
<body style="font-family: Arial, sans-serif; background:{p['app_bg']}; color:{p['text']}; padding:30px;">
<h1 style="color:{p['table_header_text']};">Images Viewer Report</h1>
{table_html}
{LIGHTBOX_JS}
</body>
</html>"""


def build_two_tab_pivot_excel(base_pivot_df, pivot_df, id_cols, image_cols, swapped_rows):
    """Same idea as build_two_tab_excel, for the pivoted PXM tabs (PXM
    Filename + URL Export, PXM Media Stack):
    - "Original": identifier column(s) + Image columns exactly as uploaded.
    - "Reorder": identifier column(s) + Image columns, ONLY for the rows
      that actually had a swap queued, with that swap applied.
    """
    original_df = base_pivot_df[id_cols + image_cols].copy()

    reorder_rows = sorted(swapped_rows)
    reorder_df = pivot_df.loc[reorder_rows, id_cols + image_cols].copy()

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        original_df.to_excel(writer, sheet_name="Original", index=False)
        reorder_df.to_excel(writer, sheet_name="Reorder", index=False)
    return buf.getvalue()


def build_two_tab_excel(base_data, data, mpn_col, master_col, title_col, image_cols, swapped_rows):
    """A plain .xlsx with two tabs (MPN, Master ID, Title, then images):
    - "Original": every row exactly as uploaded, no swaps.
    - "Reorder": ONLY the rows that actually had a swap queued, with that
      swap applied.
    """
    id_cols = [mpn_col, master_col, title_col]
    names = ["MPN", "Master ID", "Title"]
    original_df = base_data[id_cols + image_cols].copy()
    original_df.columns = names + image_cols

    reorder_rows = sorted(swapped_rows)
    reorder_df = data.loc[reorder_rows, id_cols + image_cols].copy()
    reorder_df.columns = names + image_cols

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
    """Pull the MPN, Master ID, and image sequence number out of a PXM
    filename, following the fixed convention
    MPN_MasterID_MarketplaceCode_Region_ISP_XX
    e.g. "50540_P07HR2C3_WM_US_ISP_04.jpg" -> ("50540", "P07HR2C3", 4).
    Returns (None, None, None) if the filename doesn't look like that shape."""
    if not isinstance(filename, str) or not filename.strip():
        return None, None, None
    base = filename.rsplit(".", 1)[0]
    parts = base.split("_")
    if len(parts) < 2:
        return None, None, None
    mpn = parts[0]
    master_id = parts[1]
    try:
        seq = int(parts[-1])
    except ValueError:
        seq = None
    return mpn, master_id, seq


PXM_ID_COL = "SKU / Identifier"
PXM_MPN_COL = "MPN"
PXM_MASTER_ID_COL = "Master ID"


def build_pxm_pivot(df, filename_col, url_col):
    """Group PXM export rows by MPN and lay each MPN's images out in one
    row (MPN, Master ID, then Image 1..N), ordered by the sequence number
    encoded in the filename."""
    groups = {}
    master_ids = {}
    for idx, row in df.iterrows():
        fname, url = row[filename_col], row[url_col]
        if pd.isna(fname) or pd.isna(url) or not str(url).strip():
            continue
        mpn, master_id, seq = parse_pxm_filename(str(fname))
        if mpn is None:
            continue
        groups.setdefault(mpn, []).append({"seq": seq, "url": str(url).strip(), "orig_idx": idx})
        master_ids.setdefault(mpn, master_id or "")

    if not groups:
        return pd.DataFrame(columns=[PXM_MPN_COL, PXM_MASTER_ID_COL]), []

    for mpn in groups:
        groups[mpn].sort(key=lambda r: (r["seq"] is None, r["seq"] if r["seq"] is not None else 0, r["orig_idx"]))

    max_images = max(len(v) for v in groups.values())
    image_cols = [f"Image {i + 1}" for i in range(max_images)]

    rows_out = []
    for mpn in sorted(groups.keys()):
        entry = {PXM_MPN_COL: mpn, PXM_MASTER_ID_COL: master_ids.get(mpn, "")}
        for i, col_name in enumerate(image_cols):
            entry[col_name] = groups[mpn][i]["url"] if i < len(groups[mpn]) else ""
        rows_out.append(entry)

    pivot_df = pd.DataFrame(rows_out, columns=[PXM_MPN_COL, PXM_MASTER_ID_COL] + image_cols)
    return pivot_df, image_cols


def build_media_stack_pivot(df, filename_col, position_col, link_col):
    """Group already-filtered Media Stack rows by MPN (parsed from the
    Filename column, same compulsory rule as the PXM export: everything
    before the first "_" is the MPN, the segment right after it is the
    Master ID) and lay each MPN's images out in one row, ordered by the
    Image Position column."""
    groups = {}
    master_ids = {}
    for idx, row in df.iterrows():
        fname, pos, url = row[filename_col], row[position_col], row[link_col]
        if pd.isna(fname) or pd.isna(url) or not str(url).strip():
            continue
        mpn, master_id, _seq = parse_pxm_filename(str(fname))
        if mpn is None:
            continue
        try:
            pos_val = float(pos)
        except (TypeError, ValueError):
            pos_val = None
        groups.setdefault(mpn, []).append({"pos": pos_val, "url": str(url).strip(), "orig_idx": idx})
        master_ids.setdefault(mpn, master_id or "")

    if not groups:
        return pd.DataFrame(columns=[PXM_MPN_COL, PXM_MASTER_ID_COL]), []

    for mpn in groups:
        groups[mpn].sort(key=lambda r: (r["pos"] is None, r["pos"] if r["pos"] is not None else 0, r["orig_idx"]))

    max_images = max(len(v) for v in groups.values())
    image_cols = [f"Image {i + 1}" for i in range(max_images)]

    rows_out = []
    for mpn in sorted(groups.keys()):
        entry = {PXM_MPN_COL: mpn, PXM_MASTER_ID_COL: master_ids.get(mpn, "")}
        for i, col_name in enumerate(image_cols):
            entry[col_name] = groups[mpn][i]["url"] if i < len(groups[mpn]) else ""
        rows_out.append(entry)

    pivot_df = pd.DataFrame(rows_out, columns=[PXM_MPN_COL, PXM_MASTER_ID_COL] + image_cols)
    return pivot_df, image_cols


def build_pivot_table_html(df, id_cols, image_cols, thumb_size, url_names=None):
    url_names = url_names or {}
    header_cells = "".join(f"<th>{c}</th>" for c in id_cols) + "".join(f"<th>{col}</th>" for col in image_cols)
    body_rows = []
    for _, row in df.iterrows():
        # Master ID first, then MPN, in the viewer's top bar.
        chips = [(c, row[c]) for c in reversed(id_cols)]
        cells = [f"<td>{row[c]}</td>" for c in id_cols]
        for col in image_cols:
            val = str(row[col]) if pd.notna(row[col]) else ""
            if is_url(val):
                url = val.strip()
                cells.append(img_cell(url, thumb_size, url_names.get(url) or url_basename(url), col))
            else:
                cells.append(f"<td>{val}</td>")
        body_rows.append(f"{row_open_tag(chips)}{''.join(cells)}</tr>")
    return f'<table class="iv-table"><thead><tr>{header_cells}</tr></thead><tbody>{"".join(body_rows)}</tbody></table>'


render_theme_toggle()
st.caption("Upload a file, map your columns, and preview image URLs as thumbnails.")

tab_simple, tab_pxm_url, tab_pxm_media = st.tabs(
    ["Simple File", "PXM Filename + URL Export File", "PXM Media Stack File"]
)

def render_simple_tab():
    uploaded = st.file_uploader("Upload Excel file (.xlsx)", type=["xlsx"], key="simple_uploader")

    if uploaded is None:
        st.info("Once you upload a file, you'll pick the sheet, header row, and which columns are MPN / Master ID / Title / Images.")
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

    st.subheader("3. Map your columns")
    def guess(names, fallback):
        norm = lambda t: "".join(ch for ch in str(t).lower() if ch.isalnum())
        for i, h in enumerate(headers):
            if norm(h) in names:
                return i
        return min(fallback, len(headers) - 1)

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        mpn_col = st.selectbox(
            "MPN column", headers, index=guess({"mpn"}, 0), key=f"simple_mpn_{uploaded.name}"
        )
    with col2:
        master_col = st.selectbox(
            "Master ID column", headers, index=guess({"masterid"}, 1), key=f"simple_master_{uploaded.name}"
        )
    with col3:
        title_col = st.selectbox(
            "Title column", headers, index=guess({"title"}, 2), key=f"simple_title_{uploaded.name}"
        )
    with col4:
        default_image_cols = [h for h in headers if h not in (mpn_col, master_col, title_col)]
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
            row_cols = st.columns([1, 1, 2] + [1] * len(image_cols))
            with row_cols[0]:
                st.write(f"**{row[mpn_col]}**")
            with row_cols[1]:
                st.write(row[master_col])
            with row_cols[2]:
                st.write(row[title_col])
            for i, col in enumerate(image_cols):
                with row_cols[3 + i]:
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
                    st.write(f"Row {s['row'] + 1} — {data.iloc[s['row']][mpn_col]}")
                with lc2:
                    st.write(f"{s['col_a']}  ⇄  {s['col_b']}")
                with lc3:
                    if st.button("Undo", key=f"undo_swap_{i}"):
                        st.session_state.swaps.pop(i)
                        st.rerun()
    else:
        table_html = build_table_html(data, mpn_col, master_col, title_col, image_cols, thumb_size)
        show_table(table_html, len(data), thumb_size)

    st.subheader("5. Download")
    table_html = build_table_html(data, mpn_col, master_col, title_col, image_cols, thumb_size)

    dl1, dl2, dl3 = st.columns(3)
    with dl1:
        report_html = build_full_report(table_html, get_theme_mode())
        st.download_button(
            label="Download as HTML report",
            data=report_html.encode("utf-8"),
            file_name="images_viewer_report.html",
            mime="text/html",
        )
    with dl2:
        swapped_rows = {s["row"] for s in st.session_state.swaps}
        two_tab_bytes = build_two_tab_excel(base_data, data, mpn_col, master_col, title_col, image_cols, swapped_rows)
        st.download_button(
            label="Download as Excel (Original + Reorder tabs)",
            data=two_tab_bytes,
            file_name="images_viewer_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Tab 1 'Original': MPN, Master ID, Title, Images exactly as uploaded. Tab 2 'Reorder': only the rows you swapped, with that swap applied.",
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

    base_pivot_df, image_cols = build_pxm_pivot(raw, filename_col, url_col)
    url_names = build_url_filename_map(raw, filename_col, url_col)

    if base_pivot_df.empty:
        st.warning("Couldn't find any rows with a parseable filename + URL. Check the columns picked above.")
        return

    image_row_count = raw[filename_col].notna().sum()
    st.success(f"Found {len(base_pivot_df)} SKU(s) across {image_row_count} image row(s).")

    if st.session_state.get("pxm_last_uploaded_name") != uploaded.name:
        st.session_state.pxm_swaps = []
        st.session_state.pxm_swap_pending = None
        st.session_state.pxm_last_uploaded_name = uploaded.name
    st.session_state.setdefault("pxm_swaps", [])
    st.session_state.setdefault("pxm_swap_pending", None)

    st.subheader("Display options")
    thumb_size = st.slider(
        "Thumbnail size (px)", min_value=40, max_value=240, value=80, step=10, key="pxm_url_thumb_size"
    )

    # Re-apply every queued swap to a fresh copy so the preview and
    # downloads stay in sync no matter what else reran above.
    pivot_df = apply_swaps(base_pivot_df, st.session_state.pxm_swaps)

    st.subheader("Preview")
    swap_mode = st.toggle(
        "Swap", value=False, key="pxm_swap_toggle",
        help="Click one thumbnail, then another in the same row, to swap their links.",
    )

    if swap_mode:
        pending = st.session_state.pxm_swap_pending
        if pending is not None:
            st.info(f"Selected Row {pending[0] + 1} — {pending[1]}. Click another image in the SAME row to swap with it.")
        else:
            st.caption("Click a thumbnail to start a swap, then click a second thumbnail in the same row.")

        for row_idx in range(len(pivot_df)):
            row = pivot_df.iloc[row_idx]
            row_cols = st.columns([2] + [1] * len(image_cols))
            with row_cols[0]:
                st.write(f"**{row[PXM_MPN_COL]}**")
                st.caption(f"Master ID: {row[PXM_MASTER_ID_COL]}")
            for i, col in enumerate(image_cols):
                with row_cols[1 + i]:
                    val = str(row[col]) if pd.notna(row[col]) else ""
                    if not is_url(val):
                        st.caption(col)
                        st.write(val if val else "—")
                        continue
                    st.image(val, width=thumb_size, caption=col)
                    selection = (row_idx, col)
                    is_selected = pending == selection
                    if st.button(
                        "Selected" if is_selected else "Select",
                        key=f"pxm_swap_btn_{row_idx}_{col}",
                        type="primary" if is_selected else "secondary",
                    ):
                        if pending is None:
                            st.session_state.pxm_swap_pending = selection
                        elif pending == selection:
                            st.session_state.pxm_swap_pending = None
                        elif pending[0] != row_idx:
                            st.warning("Pick two images from the same row.")
                            st.session_state.pxm_swap_pending = selection
                        else:
                            st.session_state.pxm_swaps.append(
                                {"row": row_idx, "col_a": pending[1], "col_b": col}
                            )
                            st.session_state.pxm_swap_pending = None
                        st.rerun()

        if st.session_state.pxm_swaps:
            st.write("Swaps queued:")
            for i, s in enumerate(st.session_state.pxm_swaps):
                lc1, lc2, lc3 = st.columns([3, 3, 1])
                with lc1:
                    st.write(f"Row {s['row'] + 1} — {pivot_df.iloc[s['row']][PXM_MPN_COL]}")
                with lc2:
                    st.write(f"{s['col_a']}  ⇄  {s['col_b']}")
                with lc3:
                    if st.button("Undo", key=f"pxm_undo_swap_{i}"):
                        st.session_state.pxm_swaps.pop(i)
                        st.rerun()
    else:
        table_html = build_pivot_table_html(pivot_df, [PXM_MPN_COL, PXM_MASTER_ID_COL], image_cols, thumb_size, url_names)
        show_table(table_html, len(pivot_df), thumb_size)

    st.subheader("Download")
    table_html = build_pivot_table_html(pivot_df, [PXM_MPN_COL, PXM_MASTER_ID_COL], image_cols, thumb_size, url_names)
    d1, d2, d3 = st.columns(3)
    with d1:
        report_html = build_full_report(table_html, get_theme_mode())
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
            label="Download as Excel (MPN + Master ID + Images)",
            data=buf.getvalue(),
            file_name="pxm_images_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with d3:
        pxm_swapped_rows = {s["row"] for s in st.session_state.pxm_swaps}
        two_tab_bytes = build_two_tab_pivot_excel(
            base_pivot_df, pivot_df, [PXM_MPN_COL, PXM_MASTER_ID_COL], image_cols, pxm_swapped_rows
        )
        st.download_button(
            label="Download as Excel (Original + Reorder tabs)",
            data=two_tab_bytes,
            file_name="pxm_images_report_reorder.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Tab 1 'Original': MPN, Master ID, Images exactly as uploaded. Tab 2 'Reorder': only the rows you swapped, with that swap applied.",
        )


def render_media_stack_tab():
    uploaded = st.file_uploader(
        "Upload PXM Media Stack file (.xlsx)", type=["xlsx"], key="media_stack_uploader"
    )
    if uploaded is None:
        st.info(
            "Upload the PXM Media Stack export. Pick which column holds the Stack Group name "
            "(e.g. Amazon US / Target US / Walmart US / Default), the Filename, the image position, "
            "and the image link, then pick ONE stack group to view. The MPN and Master ID are parsed "
            "from the Filename column, and that group's products are laid out as MPN + Master ID + Image 1..N."
        )
        return

    file_bytes = uploaded.getvalue()
    try:
        xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
    except Exception as exc:
        st.error(f"Could not read this file: {exc}")
        return

    sheet_name = st.selectbox("Which sheet has your data?", xls.sheet_names, key="media_stack_sheet")
    raw = pd.read_excel(xls, sheet_name=sheet_name, dtype=str)
    if raw.empty:
        st.warning("This sheet has no rows.")
        return

    cols = raw.columns.tolist()

    def default_or_first(name):
        return name if name in cols else cols[0]

    st.subheader("Map your columns")
    c1, c2 = st.columns(2)
    with c1:
        group_col = st.selectbox(
            "Media Stack Group Name column", cols,
            index=cols.index(default_or_first("Media Stack Group")), key="media_stack_group_col",
        )
    with c2:
        filename_col = st.selectbox(
            "Filename column", cols,
            index=cols.index(default_or_first("Filename")), key="media_stack_filename_col",
            help='Anything before the first "_" is read as the MPN, and the segment right after it as the Master ID.',
        )
    c3, c4 = st.columns(2)
    with c3:
        position_col = st.selectbox(
            "Image Position column", cols,
            index=cols.index(default_or_first("Media Stack Order")), key="media_stack_position_col",
        )
    with c4:
        link_col = st.selectbox(
            "Image Links column", cols,
            index=cols.index(default_or_first("Media")), key="media_stack_link_col",
        )

    group_values = sorted(raw[group_col].dropna().unique().tolist())
    if not group_values:
        st.warning("No values found in the Stack Group column picked above.")
        return
    selected_group = st.selectbox("Which Stack Group?", group_values, key="media_stack_group_value")

    filtered = raw[raw[group_col] == selected_group]
    base_pivot_df, image_cols = build_media_stack_pivot(filtered, filename_col, position_col, link_col)
    url_names = build_url_filename_map(filtered, filename_col, link_col)

    if base_pivot_df.empty:
        st.warning("Couldn't find any rows with a parseable filename + position + link for this group.")
        return

    st.success(f"Found {len(base_pivot_df)} SKU(s) in '{selected_group}' across {len(filtered)} image row(s).")

    # Swaps are scoped to this file + this stack group, since the pivoted
    # row positions only make sense for the group currently being viewed.
    reset_key = f"{uploaded.name}::{selected_group}"
    if st.session_state.get("media_last_reset_key") != reset_key:
        st.session_state.media_swaps = []
        st.session_state.media_swap_pending = None
        st.session_state.media_last_reset_key = reset_key
    st.session_state.setdefault("media_swaps", [])
    st.session_state.setdefault("media_swap_pending", None)

    st.subheader("Display options")
    thumb_size = st.slider(
        "Thumbnail size (px)", min_value=40, max_value=240, value=80, step=10, key="media_stack_thumb_size"
    )

    pivot_df = apply_swaps(base_pivot_df, st.session_state.media_swaps)

    st.subheader("Preview")
    swap_mode = st.toggle(
        "Swap", value=False, key="media_swap_toggle",
        help="Click one thumbnail, then another in the same row, to swap their links.",
    )

    if swap_mode:
        pending = st.session_state.media_swap_pending
        if pending is not None:
            st.info(f"Selected Row {pending[0] + 1} — {pending[1]}. Click another image in the SAME row to swap with it.")
        else:
            st.caption("Click a thumbnail to start a swap, then click a second thumbnail in the same row.")

        for row_idx in range(len(pivot_df)):
            row = pivot_df.iloc[row_idx]
            row_cols = st.columns([2] + [1] * len(image_cols))
            with row_cols[0]:
                st.write(f"**{row[PXM_MPN_COL]}**")
                st.caption(f"Master ID: {row[PXM_MASTER_ID_COL]}")
            for i, col in enumerate(image_cols):
                with row_cols[1 + i]:
                    val = str(row[col]) if pd.notna(row[col]) else ""
                    if not is_url(val):
                        st.caption(col)
                        st.write(val if val else "—")
                        continue
                    st.image(val, width=thumb_size, caption=col)
                    selection = (row_idx, col)
                    is_selected = pending == selection
                    if st.button(
                        "Selected" if is_selected else "Select",
                        key=f"media_swap_btn_{row_idx}_{col}",
                        type="primary" if is_selected else "secondary",
                    ):
                        if pending is None:
                            st.session_state.media_swap_pending = selection
                        elif pending == selection:
                            st.session_state.media_swap_pending = None
                        elif pending[0] != row_idx:
                            st.warning("Pick two images from the same row.")
                            st.session_state.media_swap_pending = selection
                        else:
                            st.session_state.media_swaps.append(
                                {"row": row_idx, "col_a": pending[1], "col_b": col}
                            )
                            st.session_state.media_swap_pending = None
                        st.rerun()

        if st.session_state.media_swaps:
            st.write("Swaps queued:")
            for i, s in enumerate(st.session_state.media_swaps):
                lc1, lc2, lc3 = st.columns([3, 3, 1])
                with lc1:
                    st.write(f"Row {s['row'] + 1} — {pivot_df.iloc[s['row']][PXM_MPN_COL]}")
                with lc2:
                    st.write(f"{s['col_a']}  ⇄  {s['col_b']}")
                with lc3:
                    if st.button("Undo", key=f"media_undo_swap_{i}"):
                        st.session_state.media_swaps.pop(i)
                        st.rerun()
    else:
        table_html = build_pivot_table_html(pivot_df, [PXM_MPN_COL, PXM_MASTER_ID_COL], image_cols, thumb_size, url_names)
        show_table(table_html, len(pivot_df), thumb_size)

    st.subheader("Download")
    table_html = build_pivot_table_html(pivot_df, [PXM_MPN_COL, PXM_MASTER_ID_COL], image_cols, thumb_size, url_names)
    d1, d2, d3 = st.columns(3)
    with d1:
        report_html = build_full_report(table_html, get_theme_mode())
        st.download_button(
            label="Download as HTML report",
            data=report_html.encode("utf-8"),
            file_name=f"media_stack_{selected_group.replace(' ', '_')}_report.html",
            mime="text/html",
        )
    with d2:
        buf = io.BytesIO()
        pivot_df.to_excel(buf, index=False, engine="openpyxl")
        st.download_button(
            label="Download as Excel (MPN + Master ID + Images)",
            data=buf.getvalue(),
            file_name=f"media_stack_{selected_group.replace(' ', '_')}_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with d3:
        media_swapped_rows = {s["row"] for s in st.session_state.media_swaps}
        two_tab_bytes = build_two_tab_pivot_excel(
            base_pivot_df, pivot_df, [PXM_MPN_COL, PXM_MASTER_ID_COL], image_cols, media_swapped_rows
        )
        st.download_button(
            label="Download as Excel (Original + Reorder tabs)",
            data=two_tab_bytes,
            file_name=f"media_stack_{selected_group.replace(' ', '_')}_reorder.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Tab 1 'Original': SKU/Identifier + Images exactly as uploaded. Tab 2 'Reorder': only the rows you swapped, with that swap applied.",
        )


with tab_simple:
    render_simple_tab()
with tab_pxm_url:
    render_pxm_url_tab()

with tab_pxm_media:
    render_media_stack_tab()
