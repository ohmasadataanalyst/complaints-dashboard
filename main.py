import os
import json
import requests
import sys
import pandas as pd
import gspread
import re
import gc  
from gspread_dataframe import set_with_dataframe, get_as_dataframe
from concurrent.futures import ThreadPoolExecutor, as_completed
from google.oauth2.service_account import Credentials

# --- 1. GITHUB SECRETS ---
API_KEY = os.getenv('ZENPUT_API_KEY')
GOOGLE_JSON_CREDENTIALS = os.getenv('GOOGLE_CREDENTIALS')

if not API_KEY:
    sys.exit("❌ Error: ZENPUT_API_KEY secret not found!")

# --- CONFIGURATION ---
TEMPLATE_ID = 512247
GOOGLE_SHEET_ID = "1avAzf7ROjVAy43_yDTfppUAhg6JdM191_wGeLOfICWA"
MAX_THREADS = 5  

# --- AUTHENTICATION SETUP FOR GITHUB ACTIONS ---
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

gsheet_client = get_gspread_client()
print("✅ Authentication successful via GitHub Secrets!")

# --- MAPPING CONFIGURATION ---
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

def fetch_and_parse_chunk(template_id, chunk_start, chunk_end):
    print(f"⏳ Requesting data for {chunk_start[:10]} to {chunk_end[:10]}...")
    parsed_rows = []
    start = 0
    limit = 50
    while True:
        try:
            params = {
                "form_template_id": template_id,
                "limit": limit,
                "start": start,
                "date_submitted_start": chunk_start,
                "date_submitted_end": chunk_end
            }
            resp = requests.get(
                "https://www.zenput.com/api/v3/submissions/",
                headers=zenput_headers(),
                params=params,
                timeout=30
            )

            if resp.status_code != 200:
                break

            batch = resp.json().get("data", [])
            if not batch:
                break 

            for s in batch:
                try:
                    sub_id = str(s.get("id"))
                    sm = s.get("smetadata", {})
                    answers_dict = {ans["title"]: ans.get("value") for ans in s.get("answers", [])}
                    raw_branch_code = str(answers_dict.get("اختر الفرع", "")).strip()

                    row = {
                        "Submission_ID": sub_id,
                        "اختر الفرع": BRANCH_MAP.get(raw_branch_code, raw_branch_code),
                        "محتوى شكوى العميل": answers_dict.get("محتوى شكوى العميل", ""),
                        "نوع الشكوى": answers_dict.get("نوع الشكوى"),
                        "فى حاله كانت الشكوى جوده برجاء تحديد نوع الشكوى": answers_dict.get("فى حاله كانت الشكوى جوده برجاء تحديد نوع الشكوى"),
                        "الشكوى على اي منتج؟": answers_dict.get("الشكوى على اي منتج؟"),
                        "التاريخ": sm.get("date_submitted_local", ""),
                        "مدير المنطقة المسؤول": answers_dict.get("مدير المنطقة المسؤول", ""),
                        "مدى الاجراء المتخذ": answers_dict.get("مدى الاجراء المتخذ", ""),
                        "مصدر الشكوى": answers_dict.get("مصدر الشكوى", ""),
                        "تم الطلب من خلال": answers_dict.get("تم الطلب من خلال", ""),
                        "قيمة التعويض": answers_dict.get("قيمة التعويض", ""), 
                    }
                    parsed_rows.append(row)
                except Exception:
                    continue
            
            del batch
            
            start += limit
            if start >= 9900:
                break

        except Exception as e:
            print(f"⚠️ Network timeout on {chunk_start[:10]}")
            break
            
    return parsed_rows

def fetch_submissions_optimized(template_id, start_date):
    print(f"⚡ Fetching submissions from {start_date} using MEMORY-SAFE MULTITHREADING...")
    
    start_dt = pd.to_datetime(start_date)
    now_dt = pd.Timestamp.now()
    
    chunks = []
    curr_dt = start_dt
    while curr_dt < now_dt:
        next_dt = curr_dt + pd.Timedelta(days=15)
        if next_dt > now_dt:
            next_dt = now_dt
        chunks.append((
            curr_dt.strftime('%Y-%m-%d %H:%M:%S'),
            next_dt.strftime('%Y-%m-%d %H:%M:%S')
        ))
        curr_dt = next_dt

    all_parsed_rows = []
    
    with ThreadPoolExecutor(max_workers=MAX_THREADS) as executor:
        future_to_chunk = {executor.submit(fetch_and_parse_chunk, template_id, c[0], c[1]): c for c in chunks}
        for future in as_completed(future_to_chunk):
            chunk = future_to_chunk[future]
            try:
                data = future.result()
                all_parsed_rows.extend(data)
                print(f"✅ Extracted {len(data)} records for: {chunk[0][:10]} to {chunk[1][:10]}")
                del data 
                gc.collect() 
            except Exception as e:
                print(f"❌ Chunk {chunk[0][:10]} failed: {e}")

    print(f"🎯 Total valid submissions fetched: {len(all_parsed_rows)}")
    return all_parsed_rows

def process_dataframe(rows_list):
    if not rows_list:
        return pd.DataFrame()

    df = pd.DataFrame(rows_list)
    
    df = df.drop_duplicates(subset=['Submission_ID'])

    def clean_value(value):
        if isinstance(value, list):
            return ', '.join(str(v).strip() for v in value if str(v).strip())
        elif pd.isna(value):
            return ""
        return str(value).strip()

    for col in df.columns:
        df[col] = df[col].apply(clean_value)

    df['التاريخ'] = pd.to_datetime(df['التاريخ'], errors='coerce').dt.strftime('%Y-%m-%d')
    
    df['مدى الاجراء المتخذ'] = df['مدى الاجراء المتخذ'].apply(lambda x: x.split(',')[0].strip())
    df['مدى الاجراء المتخذ'] = df['مدى الاجراء المتخذ'].replace('', 'لم يتم المراجعة')
    df['قيمة التعويض'] = pd.to_numeric(df['قيمة التعويض'], errors='coerce').fillna(0)

    cols_to_explode = [
        'نوع الشكوى',
        'فى حاله كانت الشكوى جوده برجاء تحديد نوع الشكوى',
        'الشكوى على اي منتج؟'
    ]
    for col in cols_to_explode:
        df[col] = df[col].str.split(r',\s*')
        df = df.explode(col)

    quality_col = 'فى حاله كانت الشكوى جوده برجاء تحديد نوع الشكوى'
    df[quality_col] = df[quality_col].fillna('لا علاقة لها بالجودة')
    df[quality_col] = df[quality_col].replace('', 'لا علاقة لها بالجودة')

    df['Submission_ID'] = df['Submission_ID'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()

    gc.collect()
    return df.reset_index(drop=True)

def update_google_sheet(new_df, existing_df, worksheet):
    print("ℹ️ Resolving edits and merging data...")
    
    if not existing_df.empty and 'Submission_ID' in existing_df.columns:
        existing_df['Submission_ID'] = existing_df['Submission_ID'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        
        columns_to_check = [col for col in existing_df.columns if col != 'INDEX']
        existing_df = existing_df.drop_duplicates(subset=columns_to_check)

        fetched_ids = set(new_df['Submission_ID'])
        existing_df = existing_df[~existing_df['Submission_ID'].isin(fetched_ids)]
        
        final_df = pd.concat([existing_df, new_df], ignore_index=True)
        del existing_df
        del new_df
        gc.collect()
    else:
        final_df = new_df

    final_df['TempDate'] = pd.to_datetime(final_df['التاريخ'], errors='coerce')
    final_df = final_df.sort_values(by=['TempDate', 'Submission_ID'])
    final_df.drop(columns=['TempDate'], inplace=True)

    if 'INDEX' in final_df.columns:
        final_df.drop(columns=['INDEX'], inplace=True)
    
    # Cumulative sum logic keeps INDEX exactly the same for rows sharing a Submission_ID
    submission_indices = (~final_df['Submission_ID'].duplicated()).cumsum()
    final_df.insert(0, 'INDEX', submission_indices)
    
    df_for_gsheet = final_df.fillna('')
    del final_df
    gc.collect()

    print("ℹ️ Writing updated dataset to Google Sheets...")
    worksheet.clear()
    set_with_dataframe(worksheet, df_for_gsheet, row=1, col=1, include_index=False, include_column_header=True)
    print(f"✅ Successfully wrote {len(df_for_gsheet)} total rows to Google Sheet.")

# --- Main Flow ---
if __name__ == "__main__":
    print("🚀 Starting High-Speed, Low-Memory Sync (GitHub Actions Mode)...")
    try:
        print("ℹ️ Opening Google Sheet to check existing data...")
        spreadsheet = gsheet_client.open_by_key(GOOGLE_SHEET_ID)
        worksheet = spreadsheet.get_worksheet(0)
        
        existing_df = get_as_dataframe(worksheet).dropna(how='all')
        gc.collect()
        
        if not existing_df.empty and 'التاريخ' in existing_df.columns and 'Submission_ID' in existing_df.columns:
            max_date_str = pd.to_datetime(existing_df['التاريخ'], errors='coerce').max().strftime('%Y-%m-%d')
            max_date = pd.to_datetime(max_date_str)
            thirty_days_ago = pd.Timestamp.now() - pd.Timedelta(days=30)
            
            start_date_obj = min(thirty_days_ago, max_date)
            start_date_for_api = start_date_obj.strftime('%Y-%m-%d')
            
            print(f"📅 Data found! Looking back to {start_date_for_api} to sync new records AND edits.")
        else:
            print("⚠️ Sheet is empty or missing structure. Fetching everything from 2025-01-01.")
            existing_df = pd.DataFrame() 
            start_date_for_api = "2025-01-01"

        raw_rows = fetch_submissions_optimized(TEMPLATE_ID, start_date_for_api)
        
        new_df = process_dataframe(raw_rows)
        del raw_rows
        gc.collect()
        
        if new_df.empty:
            print("ℹ️ No records found to process.")
        else:
            update_google_sheet(new_df, existing_df, worksheet)

        print("🏁 Finished.")
    except Exception as e:
        print(f"❌ Critical Failure: {e}")