import streamlit as st
import pandas as pd
from urllib.parse import urlparse
import re
import os
from io import BytesIO
from openpyxl.styles import Alignment
from openpyxl import load_workbook
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

# --- Original URL Conversion Functions ---

def clean_url(url):
    if not isinstance(url, str):
        return None
    parsed = urlparse(url)
    path = parsed.path
    if "/home/" not in path:
        return None
    cleaned = path.split("/home/", 1)[1]
    cleaned = re.sub(r'\.html($|[\?#])', r'\1', cleaned)
    return "/home/" + cleaned


def detect_header_row(df):
    for i, row in df.iterrows():
        row_str = row.astype(str).str.lower()
        if any("page title" in cell for cell in row_str) and any("url in aem" in cell for cell in row_str):
            return i
    return 0


def normalize_lang_column(colname):
    match = re.match(r'([a-z]{2}-[A-Z]{2})', str(colname))
    return match.group(1) if match else None


def process_file_original(file):
    df_preview = pd.read_excel(file, sheet_name=0, header=None)
    header_row = detect_header_row(df_preview)
    df = pd.read_excel(file, sheet_name=0, header=header_row)
    df.columns = [str(c).strip().replace("\n", " ") for c in df.columns]

    results = []
    language_columns = {col: normalize_lang_column(col)
                        for col in df.columns
                        if normalize_lang_column(col) in LANGUAGE_MAP}

    for _, row in df.iterrows():
        url = next((cell for cell in row if isinstance(cell, str) and "/home/" in cell), None)
        cleaned = clean_url(url)
        if not cleaned:
            continue
        for col, lang in language_columns.items():
            val = row.get(col, "")
            if pd.notna(val) and str(val).strip().lower() in ["x","yes","✓","✔"]:
                results.append({
                    "Original URL": url,
                    "Language": lang,
                    "Localized Path": LANGUAGE_MAP[lang] + cleaned
                })
    return pd.DataFrame(results).sort_values(by=["Language"]) if results else pd.DataFrame()


def style_and_save_excel(df):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
        ws = writer.sheets['Sheet1']
        ws.column_dimensions['A'].width = 60
        ws.column_dimensions['B'].width = 20
        ws.column_dimensions['C'].width = 80
        for row in ws.iter_rows():
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical='top')
        for cell in ws[1]:
            cell.alignment = Alignment(horizontal='center', vertical='center')
        for row in ws.iter_rows(min_row=2, min_col=2, max_col=2):
            for cell in row:
                cell.alignment = Alignment(horizontal='center', vertical='top')
    buf.seek(0)
    return buf

# --- Product Inclusion List Functions ---

def detect_first_url_product(row):
    for idx in [0,1]:
        try:
            cell = row.iloc[idx]
            if isinstance(cell, str) and cell.lower().startswith('http'):
                return cell
        except:
            pass
    for cell in row:
        if isinstance(cell, str) and cell.lower().startswith('http'):
            return cell
    return None


def extract_id_from_url(url):
    parsed = urlparse(url)
    m = re.search(r'/product/([^/?#]+)', parsed.path)
    return m.group(1) if m else None


def process_file_product(file, sheet_name='Product'):
    try:
        df = pd.read_excel(file, sheet_name=sheet_name, header=2)
    except:
        return pd.DataFrame()
    df.columns = [str(c).strip() for c in df.columns]
    langs = {col: re.match(r'([a-z]{2}-[A-Z]{2})', col).group(1)
             for col in df.columns
             if re.match(r'([a-z]{2}-[A-Z]{2})', col) and re.match(r'([a-z]{2}-[A-Z]{2})', col).group(1) in LANGUAGE_MAP}
    results = []
    for _, row in df.iterrows():
        url = detect_first_url_product(row)
        if not url:
            continue
        for col, lang in langs.items():
            val = row.get(col, '')
            if pd.notna(val) and str(val).strip().lower() in ['x','yes','✓','✔']:
                pid = extract_id_from_url(url)
                if pid:
                    results.append({'Product ID': pid, 'Language': lang})
    return pd.DataFrame(results)


def make_excel_buffer(df):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    buf.seek(0)
    return buf

# --- Streamlit UI ---
st.title("🌐 URL Converter Web App")

uploaded_file = st.file_uploader("Upload an Excel File", type=["xlsx"])
if uploaded_file:
    # Original URL conversion
    df_orig = process_file_original(uploaded_file)
    if df_orig.empty:
        st.warning("No valid data found in the file.")
    else:
        st.success("✅ File processed successfully.")
        st.dataframe(df_orig)
        st.download_button(
            label="📥 Download Excel",
            data=style_and_save_excel(df_orig),
            file_name="converted_urls.xlsx"
        )

    # Check for Product sheet presence
    try:
        sheets = pd.ExcelFile(uploaded_file).sheet_names
        if 'Product' in sheets:
            prod_sheet = 'Product'
        elif len(sheets) > 1:
            prod_sheet = sheets[1]
        else:
            prod_sheet = None
    except:
        prod_sheet = None

    # Conditionally render Product Inclusion List only if data is found
    if prod_sheet:
        df_prod = process_file_product(uploaded_file, sheet_name=prod_sheet)
        if not df_prod.empty:
            st.header("Product Inclusion List")
            buffers = {}
            for lang in sorted(df_prod['Language'].unique()):
                df_lang = df_prod[df_prod['Language'] == lang][['Product ID']]
                st.subheader(lang)
                st.table(df_lang)
                buf = make_excel_buffer(df_lang)
                st.download_button(
                    label="Download Inclusion List",
                    data=buf,
                    file_name=f"Product Inclusion List_{PROJECT_CODE}_{lang}.xlsx",
                    key=f"dl_{lang}"
                )
                buffers[lang] = buf.getvalue()
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
