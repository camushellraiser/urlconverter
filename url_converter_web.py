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

def detect_header_marketing(df):
    for i, row in df.iterrows():
        row_str = row.astype(str).str.lower()
        if any("page title" in cell for cell in row_str) and any("url in " in cell for cell in row_str):
            return i
    return 0


def detect_header_language(df):
    for i, row in df.iterrows():
        cells = row.astype(str)
        if any(re.match(r'[a-z]{2}-[A-Z]{2}', c) for c in cells):
            return i
    return 0


def normalize_lang_column(colname):
    match = re.match(r'([a-z]{2}-[A-Z]{2})', str(colname))
    return match.group(1) if match else None


def make_excel_buffer(df):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    buf.seek(0)
    return buf


def make_product_excel_buffer(df_ids):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        # write ERP in first row
        ws = writer.book.create_sheet('Sheet1')
        writer.sheets['Sheet1'] = ws
        ws['A1'] = 'ERP'
        # write only IDs starting row 2 without header
        df_ids.to_excel(writer, index=False, header=False, startrow=1)
    buf.seek(0)
    return buf

# --- Processing All Sheets ---

def process_all_sheets(file):
    xls = pd.ExcelFile(file)
    marketing_records = []
    product_records = []

    for idx, sheet in enumerate(xls.sheet_names):
        preview = pd.read_excel(file, sheet_name=sheet, header=None)
        header_row = detect_header_marketing(preview) if idx == 0 else detect_header_language(preview)
        df = pd.read_excel(file, sheet_name=sheet, header=header_row)
        df.columns = [str(c).strip().replace("\n", " ") for c in df.columns]

        lang_cols = {col: normalize_lang_column(col) for col in df.columns if normalize_lang_column(col) in LANGUAGE_MAP}
        if not lang_cols:
            continue

        for _, row in df.iterrows():
            cells = list(row)
            pid = None
            for cell in cells:
                if isinstance(cell, str):
                    m = re.search(r'A\d{3,6}', cell)
                    if m:
                        pid = m.group(0)
                        break
                    if '/product/' in cell:
                        seg = cell.split('/product/',1)[1]
                        pid = re.split(r'[/\?#]', seg)[0]
                        break
            murl = next((c for c in cells if isinstance(c, str) and '/home/' in urlparse(c).path), None)

            if pid:
                for col, lang in lang_cols.items():
                    val = row.get(col)
                    if pd.notna(val) and str(val).strip().lower() in ['x','yes','✓','✔']:
                        product_records.append({'Product ID': pid, 'Language': lang})
            elif murl:
                try:
                    path = urlparse(murl).path
                    cleaned = "/home/" + re.sub(r'.*?/home/', '', path)
                    cleaned = re.sub(r'\.html($|[\?#])', r'\1', cleaned)
                except:
                    continue
                for col, lang in lang_cols.items():
                    val = row.get(col)
                    if pd.notna(val) and str(val).strip().lower() in ['x','yes','✓','✔']:
                        marketing_records.append({
                            'Original URL': murl,
                            'Language': lang,
                            'Localized Path': LANGUAGE_MAP[lang] + cleaned
                        })

    return pd.DataFrame(marketing_records), pd.DataFrame(product_records)

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

    buffers_all = {}

    # Marketing section
    if not df_marketing.empty:
        st.subheader("Marketing URLs")
        st.dataframe(df_marketing)
        marketing_buf = make_excel_buffer(df_marketing)
        fname_mark = f"{project_code} - Converted URLs.xlsx"
        st.download_button(
            label="📥 Download Converted URLs",
            data=marketing_buf,
            file_name=fname_mark
        )
        buffers_all[fname_mark] = marketing_buf.getvalue()
    else:
        st.warning("No marketing URLs found.")

    # Product Inclusion List section
    if not df_product.empty:
        st.header("Product Inclusion List")
        for lang in sorted(df_product['Language'].unique()):
            df_lang = df_product[df_product['Language'] == lang][['Product ID']]
            st.subheader(lang)
            st.table(df_lang)
            prod_buf = make_product_excel_buffer(df_lang)
            fname_prod = f"Product Inclusion List_{project_code}_{lang}.xlsx"
            st.download_button(
                label="📥 Download Inclusion List",
                data=prod_buf,
                file_name=fname_prod,
                key=f"dl_{lang}"
            )
            buffers_all[fname_prod] = prod_buf.getvalue()
    else:
        st.warning("No product entries found.")

    # Download all files
    if buffers_all:
        zip_buf = BytesIO()
        with zipfile.ZipFile(zip_buf, 'w') as zf:
            for fname, data in buffers_all.items():
                zf.writestr(fname, data)
        zip_buf.seek(0)
        zip_name = f"{project_code} - All Downloads.zip"
        st.download_button(
            label="📥 Download All",
            data=zip_buf,
            file_name=zip_name,
            mime="application/zip"
        )

if __name__ == "__main__":
    main()
