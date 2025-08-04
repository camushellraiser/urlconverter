import streamlit as st
import pandas as pd
from urllib.parse import urlparse
import re
from io import BytesIO
from openpyxl.styles import Alignment
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

# --- Helper Functions ---

def detect_header_row(df):
    """Detect header row by finding 'page title' and any 'url in ' in the preview."""
    for i, row in df.iterrows():
        row_str = row.astype(str).str.lower()
        if any("page title" in cell for cell in row_str) and any("url in " in cell for cell in row_str):
            return i
    return 0


def normalize_lang_column(colname):
    """Normalize column names to language codes like 'en-US'."""
    match = re.match(r'([a-z]{2}-[A-Z]{2})', str(colname))
    return match.group(1) if match else None


def make_excel_buffer(df):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    buf.seek(0)
    return buf

# --- Processing Functions ---

def process_all_sheets(file):
    """
    Scan all sheets: classify each row as marketing or product.
    Marketing: rows with a '/home/' URL and no product ID.
    Product: rows with an 'A' + 3-6 digits anywhere.
    """
    xls = pd.ExcelFile(file)
    marketing_records = []
    product_records = []

    for idx, sheet in enumerate(xls.sheet_names):
        # Read sheet with appropriate header
        if idx == 0:
            preview = pd.read_excel(file, sheet_name=sheet, header=None)
            header_row = detect_header_row(preview)
            df = pd.read_excel(file, sheet_name=sheet, header=header_row)
        else:
            df = pd.read_excel(file, sheet_name=sheet, header=2)
        # Clean column names
        df.columns = [str(c).strip().replace("\n", " ") for c in df.columns]

        # Identify language columns
        lang_cols = {col: normalize_lang_column(col) for col in df.columns if normalize_lang_column(col) in LANGUAGE_MAP}
        if not lang_cols:
            continue

        for _, row in df.iterrows():
            cells = list(row)
            # Detect product ID
            pid = None
            for cell in cells:
                if isinstance(cell, str):
                    m = re.search(r'A\d{3,6}', cell)
                    if m:
                        pid = m.group(0)
                        break
            # Detect marketing URL
            murl = next((c for c in cells if isinstance(c, str) and '/home/' in urlparse(c).path), None)
            # Classify
            if pid:
                # product record
                for col, lang in lang_cols.items():
                    val = row.get(col, '')
                    if pd.notna(val) and str(val).strip().lower() in ['x','yes','✓','✔']:
                        product_records.append({'Product ID': pid, 'Language': lang})
            elif murl:
                cleaned = None
                try:
                    path = urlparse(murl).path
                    cleaned = "/home/" + re.sub(r'.*?/home/', '', path)
                    cleaned = re.sub(r'\.html($|[\?#])', r'\1', cleaned)
                except Exception:
                    cleaned = None
                if not cleaned:
                    continue
                for col, lang in lang_cols.items():
                    val = row.get(col, '')
                    if pd.notna(val) and str(val).strip().lower() in ['x','yes','✓','✔']:
                        marketing_records.append({
                            'Original URL': murl,
                            'Language': lang,
                            'Localized Path': LANGUAGE_MAP[lang] + cleaned
                        })
            # else ignore row

    df_marketing = pd.DataFrame(marketing_records)
    df_product = pd.DataFrame(product_records)
    return df_marketing, df_product

# --- Streamlit App ---

def main():
    st.title("🌐 URL Converter Web App")

    default_code = get_default_project_code()
    gts_id = st.text_input("GTS ID", value=default_code)
    project_code = gts_id.strip() if gts_id.strip() else default_code

    uploaded_file = st.file_uploader("Upload an Excel File", type=["xlsx"])
    if not uploaded_file:
        return

    df_marketing, df_product = process_all_sheets(uploaded_file)

    # Marketing URLs section
    if not df_marketing.empty:
        st.subheader("Marketing URLs")
        st.dataframe(df_marketing)
        fname = f"{project_code} - Converted URLs.xlsx"
        st.download_button("📥 Download Converted URLs", make_excel_buffer(df_marketing), fname)
    else:
        st.info("No marketing URLs found.")

    # Product Inclusion List
    if not df_product.empty:
        st.header("Product Inclusion List")
        buffers = {}
        for lang in sorted(df_product['Language'].unique()):
            df_lang = df_product[df_product['Language'] == lang][['Product ID']]
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
        zip_buf = BytesIO()
        with zipfile.ZipFile(zip_buf, 'w') as zf:
            for lang, data in buffers.items():
                zf.writestr(f"Product Inclusion List_{project_code}_{lang}.xlsx", data)
        zip_buf.seek(0)
        st.download_button("Download All", zip_buf, f"Product Inclusion Lists_{project_code}.zip", mime="application/zip")
    else:
        st.info("No product entries found.")

if __name__ == "__main__":
    main()
