import streamlit as st
import pandas as pd
from urllib.parse import urlparse
import re
from io import BytesIO
from openpyxl import Workbook
import zipfile

# Project code for naming
PROJECT_CODE = "GTS2500XX"

# Mapping from locale to region path
LANGUAGE_MAP = {
    "de-DE": "/content/lifetech/europe/en-de",
    "es-ES": "/content/lifetech/europe/en-es",
    "fr-FR": "/content/lifetech/europe/en-fr",
    "ja-JP": "/content/lifetech/japan/en-jp",
    "ko-KR": "/content/lifetech/ipac/en-kr",
    "zh-CN": "/content/lifetech/greater-china/en-cn",
    "zh-TW": "/content/lifetech/ipac/en-tw",
    "pt-BR": "/content/lifetech/latin-america/en-br",
    "es-LATAM": "/content/lifetech/latin-america/en-mx"
}


def detect_first_url(row):
    # Try first two columns first (0-indexed)
    for idx in [0, 1]:
        try:
            cell = row.iloc[idx]
        except Exception:
            continue
        if isinstance(cell, str) and cell.lower().startswith("http"):
            return cell
    # Fallback: any cell
    for cell in row:
        if isinstance(cell, str) and cell.lower().startswith("http"):
            return cell
    return None


def extract_id_from_url(url):
    # Extract the segment after '/product/' in the path
    parsed = urlparse(url)
    m = re.search(r'/product/([^/?#]+)', parsed.path)
    return m.group(1) if m else None


def process_file(file):
    # Read the 'Product' sheet with header on the third row (index 2)
    try:
        df = pd.read_excel(file, sheet_name="Product", header=2)
    except Exception:
        return pd.DataFrame()

    # Clean column names
    df.columns = [str(c).strip() for c in df.columns]

    # Identify language columns based on code in header
    language_columns = {}
    for col in df.columns:
        match = re.match(r'([a-z]{2}-[A-Z]{2})', col)
        if match:
            code = match.group(1)
            if code in LANGUAGE_MAP:
                language_columns[col] = code

    results = []
    for _, row in df.iterrows():
        url = detect_first_url(row)
        if not url:
            continue
        for col, lang in language_columns.items():
            val = row.get(col, "")
            if pd.notna(val) and str(val).strip().lower() in ["x", "yes", "✓", "✔"]:
                pid = extract_id_from_url(url)
                if pid:
                    results.append({"Product ID": pid, "Language": lang})

    return pd.DataFrame(results)


def make_excel_buffer(df_ids):
    buf = BytesIO()
    # Use Pandas ExcelWriter for simplicity
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        df_ids.to_excel(writer, index=False)
    buf.seek(0)
    return buf


# --- Streamlit UI ---
st.title("🌐 URL Converter Web App")

uploaded_file = st.file_uploader("Upload an Excel File", type=["xlsx"])
if uploaded_file:
    df_result = process_file(uploaded_file)
    if df_result.empty:
        st.warning("No valid data found on the Product sheet.")
    else:
        st.success("✅ File processed successfully.")
        st.dataframe(df_result)

        # Main section for product list
        st.header("Product Inclusion List")
        buffers = {}
        # Separate tables and downloads per language
        for lang in sorted(df_result['Language'].unique()):
            df_lang = df_result[df_result['Language'] == lang][['Product ID']]
            st.subheader(lang)
            st.table(df_lang)

            excel_buf = make_excel_buffer(df_lang)
            btn_key = f"dl_{lang}"
            st.download_button(
                label="Download Inclusion List",
                data=excel_buf,
                file_name=f"Product Inclusion List_{PROJECT_CODE}_{lang}.xlsx",
                key=btn_key
            )
            buffers[lang] = excel_buf.getvalue()

        # Download all as ZIP
        zip_buf = BytesIO()
        with zipfile.ZipFile(zip_buf, 'w') as zf:
            for lang, data in buffers.items():
                zf.writestr(
                    f"Product Inclusion List_{PROJECT_CODE}_{lang}.xlsx",
                    data
                )
        zip_buf.seek(0)
        st.download_button(
            label="Download All",
            data=zip_buf,
            file_name=f"Product Inclusion Lists_{PROJECT_CODE}.zip",
            mime="application/zip"
        )
