# 🌐 URL Converter Web App

This Streamlit app converts localized URLs from an Excel file into a formatted and styled output.

## Features
- Upload `.xlsx` files with localized URLs.
- Automatically detects languages and maps URLs.
- Removes `.html` suffixes from URLs.
- Styled Excel download with:
  - Wider columns
  - Wrapped text
  - Centered headers and language column

## How to Run on Streamlit Cloud
1. Upload the contents of this zip to a new GitHub repository.
2. Go to [Streamlit Cloud](https://share.streamlit.io/).
3. Connect your GitHub repository and deploy the `url_converter_web.py`.

## Local Development
```bash
pip install -r requirements.txt
streamlit run url_converter_web.py
```
