import streamlit as st
import pandas as pd
from urllib.parse import urlparse
import re
from io import BytesIO
from openpyxl.styles import Alignment
from openpyxl import load_workbook
import zipfile

# Project code for naming
PROJECT_CODE = "GTS2500XX"

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

def clean_url(url):
    if not isinstance(url, str):
        return None
    parsed = urlparse(url)
    path = parsed.path
    if "/home/" not in path and "/product/" not in path:
        return None
    # handle both home and product paths
    return path


def detect_first_url(row):
    for cell in row:
        if isinstance(cell, str) and ("/home/" in cell or "/product/" in cell):
            return cell
    return None


def detect_header_row(df):
    for i, row in df.iterrows():
        row_str = row.astype(str).str.lower()
        if any("page title" in cell for cell in row_str) and any("url in aem" in cell for cell in row_str):
            return i
    return 0


def normalize_lang_column(colname):
    match = re.match(r'([a-z]{2}-[A-Z]{2})', str(colname))
    return match.group(1) if match else None


def process_file(file):
    df_preview = pd.read_excel(file, sheet_name=0, header=None)
    header_row = detect_header_row(df_preview)
    df = pd.read_excel(file, sheet_name=0, header=header_row)
    df.columns = [str(c).strip().replace("\n", " ") for c in df.columns]

    language_columns = {col: normalize_lang_column(col) for col in df.columns if normalize_lang_column(col) in LANGUAGE_MAP}
    results = []
    for _, row in df.iterrows():
        original_url = detect_first_url(row)
        if not original_url:
            continue
        cleaned_path = clean_url(original_url)
        for col, lang in language_columns.items():
            val = row.get(col, "")
            if pd.notna(val) and str(val).strip().lower() in ["x", "yes", "✓", "✔"]:
                base = LANGUAGE_MAP.get(lang)
                if base:
                    full = base + cleaned_path
                    results.append({"Original URL": full, "Language": lang})
    return pd.DataFrame(results)


def extract_id_from_url(url):
    # find segment after 'product/'
    m = re.search(r'/product/([^/]+)', url)
    return m.group(1) if m else None


def make_excel(df_ids):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        df_ids.to_excel(writer, index=False)
    buf.seek(0)
    return buf


# Streamlit UI
st.title("🌐 URL Converter Web App")

uploaded_file = st.file_uploader("Upload an Excel File", type=["xlsx"])
if uploaded_file:
    df_result = process_file(uploaded_file)
    if df_result.empty:
        st.warning("No valid data found in the file.")
    else:
        st.success("✅ File processed successfully.")
        st.dataframe(df_result)

        # Production Inclusion List section
        st.header("Production Inclusion List")
        downloads = {}
        for lang in df_result['Language'].unique():
            df_lang = df_result[df_result['Language'] == lang]
            ids = df_lang['Original URL'].apply(extract_id_from_url)
            df_ids = pd.DataFrame({"Product ID": ids})

            st.subheader(lang)
            st.table(df_ids)

            excel_buf = make_excel(df_ids)
            btn_key = f"dl_{lang}"
            st.download_button(
                label="Download Inclusion List",
                data=excel_buf,
                file_name=f"Product Inclusion List_{PROJECT_CODE}_{lang}.xlsx",
                key=btn_key
            )
            downloads[lang] = excel_buf.getvalue()

        # Download All as ZIP
        zip_buf = BytesIO()
        with zipfile.ZipFile(zip_buf, 'w') as zf:
            for lang, data in downloads.items():
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
