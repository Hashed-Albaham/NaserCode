import pandas as pd
import requests
import os
from datetime import datetime

# 1. قراءة الإعدادات من متغيرات البيئة (GitHub Secrets) بدلاً من كتابتها مكشوفة
DROPBOX_URL = os.getenv("DROPBOX_URL")
BOT_TOKEN = os.getenv("BOT_TOKEN")
CHAT_ID = os.getenv("CHAT_ID")

# مسار الملف المحلي داخل بيئة تشغيل GitHub
FILE_PATH = "TLS-STC-Permit-Report.xlsx"

def send_telegram_message(message):
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/sendMessage"
    payload = {"chat_id": CHAT_ID, "text": message, "parse_mode": "Markdown"}
    try:
        response = requests.post(url, json=payload)
        return response.json()
    except Exception as e:
        print(f"Error sending message: {e}")
        return None

def download_file():
    try:
        response = requests.get(DROPBOX_URL)
        with open(FILE_PATH, 'wb') as f:
            f.write(response.content)
        return True
    except Exception as e:
        print(f"Error downloading file: {e}")
        return False

def is_actual_date(val):
    """Strict check if a value is a date object or a date-like string."""
    if pd.isnull(val):
        return False
    if isinstance(val, (datetime, pd.Timestamp)):
        return True
    val_str = str(val).strip()
    if any(char.isdigit() for char in val_str) and any(sep in val_str for sep in ['-', '/', '.']):
        try:
            pd.to_datetime(val_str)
            return True
        except:
            return False
    return False

def check_excel():
    if not download_file():
        return "❌ فشل تحميل الملف من Dropbox"

    xl = pd.ExcelFile(FILE_PATH)
    alerts_open = []      # For "الرخص المفتوحة"
    alerts_followup = []  # For "رخص منتهية أعمالها"
    
    # Keywords to look for in column headers
    rem_days_keywords = ['rem. days', 'remaining days', 'rem.days']
    close_date_keywords = ['date of close request', 'date of close', 'close request']
    site_name_keywords = ['site name', 'sitename', 'project name', 'site']
    status_keywords = ['status', 'permit status']
    
    for sheet in xl.sheet_names:
        try:
            df_raw = pd.read_excel(FILE_PATH, sheet_name=sheet, header=None)
            header_row = -1
            rem_days_idx = -1
            close_date_idx = -1
            site_name_idx = -1
            status_idx = -1
            
            for r_idx, row in df_raw.iterrows():
                row_str = [str(v).lower().strip() for v in row]
                found_rem = False
                for c_idx, val in enumerate(row_str):
                    if any(kw in val for kw in rem_days_keywords):
                        rem_days_idx = c_idx
                        found_rem = True
                    if any(kw in val for kw in close_date_keywords):
                        close_date_idx = c_idx
                    if any(kw in val for kw in site_name_keywords):
                        site_name_idx = c_idx
                    if any(kw == val or val in status_keywords for kw in status_keywords):
                        status_idx = c_idx
                
                if found_rem:
                    header_row = r_idx
                    break
            
            if header_row != -1:
                df = pd.read_excel(FILE_PATH, sheet_name=sheet, skiprows=header_row)
                rem_days_col = df.columns[rem_days_idx]
                close_date_col = df.columns[close_date_idx] if close_date_idx != -1 else None
                site_name_col = df.columns[site_name_idx] if site_name_idx != -1 else None
                status_col = df.columns[status_idx] if status_idx != -1 else None
                
                for index, row in df.iterrows():
                    try:
                        if status_col:
                            status_val = str(row[status_col]).strip().upper()
                            if status_val != "APPROVED":
                                continue
                        
                        days_val = pd.to_numeric(row[rem_days_col], errors='coerce')
                        if pd.notnull(days_val) and days_val < 2:
                            site_info = str(row[site_name_col]) if site_name_col else f"Row {index+1}"
                            if site_info.lower() == 'nan': site_info = f"Row {index+1}"
                            
                            # Reverted format: Site and Remaining Days only
                            alert_text = f"📍 الموقع: {site_info}\n⏳ المتبقي: {int(days_val)} يوم"
                            
                            if close_date_col:
                                close_val = row[close_date_col]
                                if pd.isnull(close_val) or str(close_val).strip() == "":
                                    alerts_open.append(alert_text)
                                elif is_actual_date(close_val):
                                    continue # Actual date means closed
                                else:
                                    # Text like 'للاغلاق' goes here
                                    alerts_followup.append(alert_text)
                            else:
                                alerts_open.append(alert_text)
                    except:
                        continue
        except Exception as e:
            print(f"Error processing sheet {sheet}: {e}")
            continue
                    
    # Send Message 1: رخص منتهية أعمالها يجب متابعتها مع المختبر
    if alerts_followup:
        header = f"📋 *رخص منتهية أعمالها يجب متابعتها مع المختبر ({len(alerts_followup)} رخصة)*\n\n"
        for i in range(0, len(alerts_followup), 20):
            chunk = alerts_followup[i:i+20]
            msg = header + "\n\n".join(chunk)
            send_telegram_message(msg)

    # Send Message 2: الرخص المفتوحة
    if alerts_open:
        header = f"🔓 *الرخص المفتوحة ({len(alerts_open)} رخصة)*\n\n"
        for i in range(0, len(alerts_open), 20):
            chunk = alerts_open[i:i+20]
            msg = header + "\n\n".join(chunk)
            send_telegram_message(msg)

    if not alerts_followup and not alerts_open:
        return "✅ لا توجد رخص معتمدة (APPROVED) تنتهي قريباً."
    return "✅ تم إرسال التنبيهات بنجاح."

if __name__ == "__main__":
    result = check_excel()
    print(result)
