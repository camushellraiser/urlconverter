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

# --- Original URL Conversion Section ---

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


def detect_first_url(row):
    for cell in row:
        if isinstance(cell, str) and "/home/" in cell:
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


def process_file_original(file):
    df_preview = pd.read_excel(file, sheet_name=0, header=None)
    header_row = detect_header_row(df_preview)
    df = pd.read_excel(file, sheet_name=0, header=header_row)
    df.columns = [str(c).strip().replace("\n", " ") for c in df.columns]
    results = []
    language_columns = {}
    for col in df.columns:
        code = normalize_lang_column(col)
        if code in LANGUAGE_MAP:
            language_columns[col] = code
    for _, row in df.iterrows():
        original_url = detect_first_url(row)
        cleaned_path = clean_url(original_url)
        if not cleaned_path:
            continue
        for col_name, lang_code in language_columns.items():
            cell_value = row.get(col_name, "")
            if pd.notna(cell_value) and str(cell_value).strip().lower() in ["x", "yes", "✓", "✔"]:
                localized_base = LANGUAGE_MAP.get(lang_code)
                if localized_base:
                    results.append({
                        "Original URL": original_url,
                        "Language": lang_code,
                        "Localized Path": localized_base + cleaned_path
                    })
    result_df = pd.DataFrame(results)
    return result_df.sort_values(by=["Language"]) if not result_df.empty else pd.DataFrame()


def style_and_save_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
        worksheet = writer.sheets["Sheet1"]
        worksheet.column_dimensions["A"].width = 60
        worksheet.column_dimensions["B"].width = 20
        worksheet.column_dimensions["C"].width = 80
        for row in worksheet.iter_rows():
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
        for cell in worksheet[1]:
            cell.alignment = Alignment(horizontal="center", vertical="center")
        for row in worksheet.iter_rows(min_row=2, min_col=2, max_col=2):
            for cell in row:
                cell.alignment = Alignment(horizontal="center", vertical="top")
    output.seek(0)
    return output

# --- Streamlit UI: Original Section ---
st.title("🌐 URL Converter Web App")
uploaded_file = st.file_uploader("Upload an Excel File", type=["xlsx"])
if uploaded_file:
    df_result = process_file_original(uploaded_file)
    if df_result.empty:
        st.warning("No valid data found in the file.")
    else:
        st.success("✅ File processed successfully.")
        st.dataframe(df_result)
        styled_excel = style_and_save_excel(df_result)
        st.download_button("📥 Download Excel", data=styled_excel, file_name="converted_urls.xlsx")

# --- Production Inclusion List Section ---

def detect_first_url_product(row):
    # Try first two columns (0,1), then fallback
    for idx in [0,1]:
        try:
            cell = row.iloc[idx]
            if isinstance(cell, str) and cell.lower().startswith("http"):
                return cell
        except:
            pass
    for cell in row:
        if isinstance(cell, str) and cell.lower().startswith("http"):
            return cell
    return None


def extract_id_from_url(url):
    parsed = urlparse(url)
    m = re.search(r'/product/([^/?#]+)', parsed.path)
    return m.group(1) if m else None


def process_file_product(file):
    try:
        df = pd.read_excel(file, sheet_name="Product", header=2)
    except:
        return pd.DataFrame()
    df.columns = [str(c).strip() for c in df.columns]
    language_columns = {}
    for col in df.columns:
        match = re.match(r'([a-z]{2}-[A-Z]{2})', col)
        if match and match.group(1) in LANGUAGE_MAP:
            language_columns[col] = match.group(1)
    results = []
    for _, row in df.iterrows():
        url = detect_first_url_product(row)
        if not url:
            continue
        for col, lang in language_columns.items():
            val = row.get(col, "")
            if pd.notna(val) and str(val).strip().lower() in ["x","yes","✓","✔"]:
                pid = extract_id_from_url(url)
                if pid:
                    results.append({"Product ID": pid, "Language": lang})
    return pd.DataFrame(results)


def make_excel_buffer(df_ids):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        df_ids.to_excel(writer, index=False)
    buf.seek(0)
    return buf

# UI for Product Inclusion List
st.header("Product Inclusion List")
if uploaded_file:
    df_prod = process_file_product(uploaded_file)
    if df_prod.empty:
        st.warning("No valid data found on the Product sheet.")
    else:
        st.success("✅ Product sheet processed.")
        st.dataframe(df_prod)
        buffers = {}
        for lang in sorted(df_prod['Language'].unique()):
            df_lang = df_prod[df_prod['Language']==lang][['Product ID']]
            st.subheader(lang)
            st.table(df_lang)
            buf = make_excel_buffer(df_lang)
            key = f"download_{lang}"
            st.download_button(
                label="Download Inclusion List",
                data=buf,
                file_name=f"Product Inclusion List_{PROJECT_CODE}_{lang}.xlsx",
                key=key
            )
            buffers[lang] = buf.getvalue()
        zip_buf = BytesIO()
        with zipfile.ZipFile(zip_buf,'w') as zf:
            for lang,data in buffers.items():
                zf.writestr(f"Product Inclusion List_{PROJECT_CODE}_{lang}.xlsx", data)
        zip_buf.seek(0)
        st.download_button(
            label="Download All",
            data=zip_buf,
            file_name=f"Product Inclusion Lists_{PROJECT_CODE}.zip",
            mime="application/zip"
        )
