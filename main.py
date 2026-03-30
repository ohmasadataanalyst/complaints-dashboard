import os
import json
import requests
import sys
import datetime
import pytz
import pandas as pd
import numpy as np
import gspread
import re
from datetime import datetime, timedelta
from google.oauth2.service_account import Credentials
from gspread_dataframe import set_with_dataframe

# --- 1. سحب المفاتيح من GitHub Secrets بدل userdata ---
API_KEY = os.getenv('ZENPUT_API_KEY')
GOOGLE_JSON_CREDENTIALS = os.getenv('GOOGLE_CREDENTIALS')

# --- CONFIGURATION ---
TEMPLATE_ID = 512247
TZ = pytz.timezone("Asia/Baghdad")
GOOGLE_SHEET_ID = "1avAzf7ROjVAy43_yDTfppUAhg6JdM191_wGeLOfICWA"

# --- AUTHENTICATION SETUP (GITHUB VERSION) ---
def get_gspread_client():
    if not GOOGLE_JSON_CREDENTIALS:
        sys.exit("❌ Error: GOOGLE_CREDENTIALS secret not found!")
    
    try:
        creds_dict = json.loads(GOOGLE_JSON_CREDENTIALS)
        creds = Credentials.from_service_account_info(
            creds_dict, 
            scopes=[
                "https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive"
            ]
        )
        return gspread.authorize(creds)
    except Exception as e:
        sys.exit(f"❌ Authentication failed: {e}")

# استخدام العميل الجديد
gc = get_gspread_client()
print("✅ Authentication successful via GitHub Secrets!")

# --- MAPPING CONFIGURATION (نفس الكود الأصلي بدون تغيير) ---
RAW_BRANCH_MAPPING_DATA = """
2197299 "LBRUH   B07"
2239240 FYJED  B32
2235670 "ANRUH   B31"
2190657 "SLAHS   B23"
2164026 "NDRUH   B15"
2164019 "SWRUH   B08"
2203271 "SARUH   B27"
2164017 "DARUH   B06"
2164032 "KRRUH   B21"
2164031 "SFJED   B24"
2164025 "RBRUH   B14"
2164016 RWRUH B05
2197297 "NSRUH   B04"
2164021 "SHRUH   B10"
2164013 "KHRUH   B02"
2155652 "NURUH   B01"
2164023 "TWRUH   B12"
2164020 "AZRUH   B09"
2199002 "RWAHS   B25"
2242934 "HIRJED      B33"
2164022 "NRRUH   B11"
2164030 "MURUH   B19"
2164014 "GHRUH   B03"
2211854 QARUH B30
2169459 "Lubda  Alaqeq Branch    LB01"
2254072 Garatiss QB03
2256386 Garatiss QB04
2232755 Garatis As Suwaidi - قراطيس السويدي   QB01
2258220 PSJED   B36
2185452 "OBJED   B22"
2243963 URRUH B34
2222802 "Lubda Alkhaleej Branch      LB02"
2199835 "HAJED   B26"
2210205 "MAJED   B28"
2250799 IRRUH B35
2164027 "BDRUH   B16"
2155654 "AQRUH   B13"
2197298 "TKRUH   B18"
2239240 "FAYJED      B32"
2250799 IRRUH35
2211854 "QADRUH      B30"
2243963 "URURUH      B34"
2239240 "FAYJED      B32"
2164017 Aldaraiah - الدرعية
2203271 Alsaadah branch - فرع السعادة
2155654 Al Aqeeq - العقيق
2164032 Alkharj - الخرج
2190657 Al Sulimaniyah Al Hofuf - السلمانية الهفوف
2211854 Al Qadisiyyah branch - فرع القادسية
2164013 Alkaleej - الخليج
2164027 Albadeah - البديعة
2171883 Twesste - تويستي TW01
2235805 Garatis Alnargis -  قراطيس النرجس  QB02
2164016 "RAWRUH      B05"
2164028 "QRRUH B17"
2257790 SHWMAK B37
2260889 UHDMM B38
2263062 HSRUH B39
"""

def create_branch_map_prioritized(raw_data):
    branch_map = {}
    lines = raw_data.strip().split('\n')
    code_pattern = re.compile(r'\b[A-Z]{1,3}[0-9]{1,2}\b')
    for line in lines:
        line = line.strip()
        if not line: continue
        parts = line.split(None, 1)
        if len(parts) == 2:
            code, branch_name = parts
            cleaned_name = ' '.join(branch_name.strip().strip('"').split())
            if code_pattern.search(cleaned_name):
                branch_map[code.strip()] = cleaned_name
    for line in lines:
        line = line.strip()
        if not line: continue
        parts = line.split(None, 1)
        if len(parts) == 2:
            code, branch_name = parts
            code = code.strip()
            if code not in branch_map:
                cleaned_name = ' '.join(branch_name.strip().strip('"').split())
                branch_map[code] = cleaned_name
    return branch_map

BRANCH_MAP = create_branch_map_prioritized(RAW_BRANCH_MAPPING_DATA)

def zenput_headers():
    return {"X-API-TOKEN": API_KEY, "Content-Type": "application/json"}

def fetch_submissions(template_id):
    print("🔍 Attempting to fetch 2025 & 2026 (Jan-Apr) submissions...")
    subs = []
    start = 0
    limit = 50
    max_records = 9000
    fetch_start_date = "2025-01-01"
    
    while len(subs) < max_records:
        try:
            params = {
                "form_template_id": template_id,
                "limit": limit,
                "start": start,
                "date_submitted_start": fetch_start_date
            }
            resp = requests.get("https://www.zenput.com/api/v3/submissions/", headers=zenput_headers(), params=params)
            if resp.status_code != 200:
                break
            batch = resp.json().get("data", [])
            if not batch: break
            
            batch_filtered = []
            for s in batch:
                submitted_date = s["smetadata"]["date_submitted_local"]
                is_2025 = submitted_date.startswith("2025-")
                is_2026_apr = submitted_date.startswith("2026-") and submitted_date <= "2026-04-30"
                if is_2025 or is_2026_apr:
                    batch_filtered.append(s)
            subs.extend(batch_filtered)
            start += limit
        except Exception:
            break
    return subs

def submissions_to_filtered_df(subs):
    if not subs: return pd.DataFrame()
    rows = []
    for original_id, s in enumerate(subs, 1):
        try:
            sm = s["smetadata"]
            answers = {ans["title"]: ans.get("value") for ans in s.get("answers", [])}
            raw_branch_code = str(answers.get("اختر الفرع", "")).strip()
            row = {
                "original_id": original_id,
                "اختر الفرع": BRANCH_MAP.get(raw_branch_code, raw_branch_code),
                "محتوى شكوى العميل": answers.get("محتوى شكوى العميل", ""),
                "نوع الشكوى": answers.get("نوع الشكوى"),
                "فى حاله كانت الشكوى جوده برجاء تحديد نوع الشكوى": answers.get("فى حاله كانت الشكوى جوده برجاء تحديد نوع الشكوى"),
                "الشكوى على اي منتج؟": answers.get("الشكوى على اي منتج؟"),
                "التاريخ": sm["date_submitted_local"],
                "مدير المنطقة المسؤول": answers.get("مدير المنطقة المسؤول", ""),
                "مدى الاجراء المتخذ": answers.get("مدى الاجراء المتخذ", ""),
                "مصدر الشكوى": answers.get("مصدر الشكوى", ""),
                "تم الطلب من خلال": answers.get("تم الطلب من خلال", ""),
            }
            rows.append(row)
        except Exception: continue

    df = pd.DataFrame(rows)
    def clean_value(value):
        if isinstance(value, list): return ', '.join(str(v).strip() for v in value if str(v).strip())
        return str(value).strip() if not pd.isna(value) else ""

    for col in df.columns:
        if col != 'original_id': df[col] = df[col].apply(clean_value)

    df['التاريخ'] = pd.to_datetime(df['التاريخ'], errors='coerce').dt.strftime('%Y-%m-%d')
    df['مدى الاجراء المتخذ'] = df['مدى الاجراء المتخذ'].apply(lambda x: x.split(',')[0].strip())
    df['مدى الاجراء المتخذ'].replace('', 'لم يتم المراجعة', inplace=True)

    cols_to_explode = ['نوع الشكوى', 'فى حاله كانت الشكوى جوده برجاء تحديد نوع الشكوى', 'الشكوى على اي منتج؟']
    for col in cols_to_explode:
        df[col] = df[col].str.split(',\s*')
        df = df.explode(col)

    quality_col = 'فى حاله كانت الشكوى جوده برجاء تحديد نوع الشكوى'
    df[quality_col] = df[quality_col].fillna('لا علاقة لها بالجودة').replace('', 'لا علاقة لها بالجودة')
    df.rename(columns={'original_id': 'INDEX'}, inplace=True)
    df = df[['INDEX'] + [c for c in df.columns if c != 'INDEX']]
    return df.reset_index(drop=True)

def write_to_google_sheet(df, gc_client):
    try:
        spreadsheet = gc_client.open_by_key(GOOGLE_SHEET_ID)
        worksheet = spreadsheet.get_worksheet(0)
        header = worksheet.row_values(1) or list(df.columns)
        worksheet.clear()
        worksheet.update('A1', [header])
        set_with_dataframe(worksheet, df.fillna(''), row=2, col=1, include_index=False, include_column_header=False)
        print(f"✅ Successfully updated {len(df)} rows.")
    except Exception as e:
        print(f"❌ Error: {e}")

# --- Main Flow ---
if __name__ == "__main__":
    print("🚀 Starting Sync...")
    try:
        submissions = fetch_submissions(TEMPLATE_ID)
        if submissions:
            df_final = submissions_to_filtered_df(submissions)
            if not df_final.empty:
                write_to_google_sheet(df_final, gc)
        print("🏁 Finished.")
    except Exception as e:
        print(f"❌ Critical Failure: {e}")
