import pandas as pd
import re
import os

def sanitize_dataframe(df):
    """
    Sanitize all values in a DataFrame using Persian normalization rules.
    """
    persian_map = {
        'ي': 'ی',
        'ك': 'ک',
        '\u200c': ' ',  # ZWNJ to space
    }

    def sanitize(text):
        if pd.isna(text):
            return ""
        text = str(text)
        text = text.replace('pds , ','pds و ')
        text = text.replace('ISO 10002 , ISO 10004','ISO 10002 و ISO 10004')

        text = text.replace(', ', '&&&')
        for arabic_char, persian_char in persian_map.items():
            text = text.replace(arabic_char, persian_char)
        text = text.replace('&&&', 'ampersand')
        text = text.replace('،', '')
        text = re.sub(r'[^a-zA-Z0-9\u0600-\u06FF\s]', '', text)
        text = text.lower()  # Convert all English letters to lowercase here
        text = text.replace('ampersand', '&&&')
        text = re.sub(r'\s+', ' ', text).strip()
        return text

    return df.map(sanitize)


def load_sanitized_data(file_path):
    """
    Load and sanitize a single-sheet Excel file.
    Applies special rules for 'after.xlsx' and 'sales.xlsx'.
    """
    try:
        df = pd.read_excel(file_path)
        df = sanitize_dataframe(df)

        filename = os.path.basename(file_path).lower()

        if 'after' in filename:
            if 'نام خودرو' in df.columns:
                df['نام خودرو'] = df['نام خودرو'].apply(lambda x: x if x else 'عمومی')
        elif 'sales' in filename:
            if 'نام خودرو' not in df.columns:
                df['نام خودرو'] = 'عمومی'

        return df

    except Exception as e:
        print(f"Error loading {file_path}: {e}")
        return pd.DataFrame()


def load_all_sanitized_sheets(file_path):
    """
    Load and sanitize all worksheets in an Excel file.
    Returns a dictionary {sheet_name: sanitized DataFrame}.
    Applies same special handling as load_sanitized_data.
    """
    try:
        excel_file = pd.ExcelFile(file_path)
        sanitized_sheets = {}
        filename = os.path.basename(file_path).lower()

        for sheet_name in excel_file.sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            df = sanitize_dataframe(df)

            if 'after' in filename:
                if 'نام خودرو' in df.columns:
                    df['نام خودرو'] = df['نام خودرو'].apply(lambda x: x if x else 'عمومی')
            elif 'sales' in filename:
                if 'نام خودرو' not in df.columns:
                    df['نام خودرو'] = 'عمومی'

            sanitized_sheets[sheet_name] = df

        return sanitized_sheets
    except Exception as e:
        print(f"Error loading sheets from {file_path}: {e}")
        return {}
