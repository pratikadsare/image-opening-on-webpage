# Images Viewer

Upload an Excel file and preview image URL columns as inline thumbnails, with custom column mapping, adjustable thumbnail size, sheet/tab selection, and export to a standalone HTML report.

This is a Python (Streamlit) rebuild of a browser-only HTML/JS tool, with added flexibility: you are no longer locked into "column 1 = SKU, column 2 = Title, column 3+ = images" — you pick which columns map to what after upload.

## Setup

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then open the local URL Streamlit prints (usually http://localhost:8501).

## How to use

1. Upload an `.xlsx` file.
2. Pick which sheet/tab has your data.
3. Tell it which row holds the column headers.
4. Map which column is SKU, which is Title, and which column(s) hold image URLs (any cell starting with `http` renders as a thumbnail; everything else shows as plain text).
5. Adjust the thumbnail size if you want bigger/smaller previews.
6. Download the result as a standalone HTML report to share.
