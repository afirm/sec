import pandas as pd
import re

def load_sanitized_data(file_path='res/raw.xlsx', output_path='res/sanitized.xlsx'):
    # Load Excel file
    df = pd.read_excel(file_path)

    # Define Persian character mapping
    persian_map = {
        'ي': 'ی',
        'ك': 'ک',
        '\u200c': ' ',
    }

    def sanitize(text):
        if pd.isna(text):
            return ""
        text = str(text)
        text = text.replace(', ', '&&&')
        for arabic_char, persian_char in persian_map.items():
            text = text.replace(arabic_char, persian_char)
        text = text.replace('&&&', 'AMP')
        text = re.sub(r'[^a-zA-Z0-9\u0600-\u06FF\s]', '', text)
        text = text.replace('AMP', '&&&')
        text = re.sub(r'\s+', ' ', text).strip()
        return text

    # Apply sanitization
    sanitized_df = df.map(sanitize)

    # # Save to Excel
    # sanitized_df.to_excel(output_path, index=False)

    return sanitized_df
