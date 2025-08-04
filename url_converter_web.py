import streamlit as st
import pandas as pd
from urllib.parse import urlparse
import re
from io import BytesIO
from openpyxl.styles import Alignment
from openpyxl import load_workbook
import zipfile

# Default project code for naming
def get_default_project_code():
    return "GTS2500XX"

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

# --- Functions for Marketing (Original URL Conversion) ---

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


def process_file_marketing(file):
    # Always read the first sheet (index 0) for marketing
    df_preview = pd.read_excel(file, sheet_name=0, header=None)
    header_row = detect_header_row(df_preview)
    df = pd.read_excel(file, sheet_name=0, header=header_row)
    df.columns = [str(c).strip().replace("\n", " ") for c in df.columns]

    language_columns = {
        col: normalize_lang_column(col)
        for col in df.columns
        if normalize_lang_column(col) in LANGUAGE_MAP
    }
    results = []
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

# --- Functions for Product Inclusion List ---

def process_file_product(file):
    xls = pd.ExcelFile(file)
    sheet = 'Product' if 'Product' in xls.sheet_names else (xls.sheet_names[1] if len(xls.sheet_names) > 1 else None)
    if not sheet:
        return pd.DataFrame()
    df = pd.read_excel(file, sheet_name=sheet, header=2)
    df.columns = [str(c).strip() for c in df.columns]

    langs = {
        col: re.match(r'([a-z]{2}-[A-Z]{2})', col).group(1)
        for col in df.columns
        if re.match(r'([a-z]{2}-[A-Z]{2})', col)
           and re.match(r'([a-z]{2}-[A-Z]{2})', col).group(1) in LANGUAGE_MAP
    }
    results = []
    for _, row in df.iterrows():
        pid = None
        # scan all cells for A### pattern
        for cell in row:
            if isinstance(cell, str):
                m = re.search(r'A\d{3,6}', cell)
                if m:
                    pid = m.group(0)
                    break
        if not pid:
            continue
        for col, lang in langs.items():
            val = row.get(col, '')
            if pd.notna(val) and str(val).strip().lower() in ['x','yes','✓','✔']:
                results.append({'Product ID': pid, 'Language': lang})
    return pd.DataFrame(results)

# --- Excel buffer helper ---

def make_excel_buffer(df):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    buf.seek(0)
    return buf

# --- Streamlit UI ---

def main():
    st.title("🌐 URL Converter Web App")
    default_code = get_default_project_code()
    gts_id = st.text_input("GTS ID", value=default_code)
    project_code = gts_id.strip() if gts_id.strip() else default_code

    uploaded_file = st.file_uploader("Upload an Excel File", type=["xlsx"])
    if not uploaded_file:
        return

    # Marketing section
    df_marketing = process_file_marketing(uploaded_file)
    if not df_marketing.empty:
        st.subheader("Marketing URLs")
        st.dataframe(df_marketing)
        filename = f"{project_code} - Converted URLs.xlsx"
        st.download_button(
            label="📥 Download Converted URLs",
            data=make_excel_buffer(df_marketing),
            file_name=filename
        )
    else:
        st.warning("No valid data found in the Marketing sheet.")

    # Product section
    df_prod = process_file_product(uploaded_file)
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
                file_name=f"Product Inclusion List_{project_code}_{lang}.xlsx",
                key=f"dl_{lang}"
            )
            buffers[lang] = buf.getvalue()
        # Download all as ZIP
        zip_buf = BytesIO()
        with zipfile.ZipFile(zip_buf, 'w') as zf:
            for lang,data in buffers.items():
                zf.writestr(
                    f"Product Inclusion List_{project_code}_{lang}.xlsx",
                    data
                )
        zip_buf.seek(0)
        st.download_button(
            label="Download All",
            data=zip_buf,
            file_name=f"Product Inclusion Lists_{project_code}.zip",
            mime="application/zip"
        )

if __name__ == "__main__":
    main()
