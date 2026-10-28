import os
import re
import csv
import time
import random
import datetime
import base64
import requests
import dns.resolver
import smtplib
from urllib.parse import urlparse
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import streamlit as st

# File to persist global usage counter across app restarts (prevents quota bypass via page refreshes)
LIMIT_TRACKER_FILE = "daily_usage_tracker.txt"


def get_base64_of_bin_file(bin_file):
    """Converts a local binary file (like an image) into a Base64 string."""
    with open(bin_file, 'rb') as f:
        data = f.read()
    return base64.b64encode(data).decode()


def extract_domain(url):
    """Extracts a clean domain from an organisation's website URL."""
    if not url or str(url).strip() == '' or str(url).lower() == 'nan':
        return ""
    parsed = urlparse(str(url))
    domain = parsed.netloc if parsed.netloc else parsed.path
    domain = domain.split('/')[0]
    if domain.startswith('www.'):
        domain = domain[4:]
    return domain.strip().lower()


def generate_permutations(first_name, last_name, domain):
    """Generates standard corporate B2B email permutations."""
    f = str(first_name).strip().lower()
    l = str(last_name).strip().lower()
    d = str(domain).strip().lower()
    
    if not f or not l or not d or f == 'nan' or l == 'nan' or d == 'nan':
        return []
    
    f = ''.join(e for e in f if e.isalnum())
    l = ''.join(e for e in l if e.isalnum())
    
    if not f or not l:
        return []
        
    return [
        f"{f}.{l}@{d}",    # firstname.lastname@company.com
        f"{f[0]}{l}@{d}",  # flastname@company.com
        f"{f}@{d}"         # firstname@company.com
    ]


def verify_smtp_email(email, sender_email="verify@gmail.com"):
    """Pings target mail server via SMTP with MX record lookup."""
    try:
        domain = email.split('@')[1]
        records = dns.resolver.resolve(domain, 'MX')
        mx_record = str(records[0].exchange)
        
        server = smtplib.SMTP(timeout=10)
        server.connect(mx_record)
        server.helo('verification.local')
        server.mail(sender_email)
        code, message = server.rcpt(email)
        server.quit()
        
        if code == 250:
            return "Valid"
        elif 400 <= code < 500:
            return "Greylisted"
        else:
            return "Invalid"
    except Exception:
        return "Failed/Blocked"


def find_exposed_phone_number(first_name, last_name, company_name, domain=""):
    """
    Searches Google for personal direct phone lines using targeted queries
    and filters out generic company HQ/switchboard lines.
    """
    api_key = st.secrets.get("GOOGLE_API_KEY", "")
    cx = st.secrets.get("SEARCH_ENGINE_ID", "")
    
    if not api_key or not cx:
        return "Not Configured"

    queries = [
        f'"{first_name} {last_name}" "{company_name}" filetype:pdf (phone OR mobile)',
        f'site:github.com "{first_name} {last_name}" (phone OR mobile OR "cell")',
        f'"{first_name} {last_name}" "{company_name}" ("speaker" OR "press release") (phone OR contact)',
        f'"{first_name} {last_name}" "{company_name}" (phone OR mobile OR contact)'
    ]
    
    for query in queries:
        url = f"https://www.googleapis.com/customsearch/v1?q={query}&key={api_key}&cx={cx}"
        try:
            response = requests.get(url, timeout=5).json()
            
            # Rate limiting / Sleep timer to mimic human browsing speed and avoid API throttles
            time.sleep(random.uniform(2.0, 4.0))
            
            for item in response.get('items', []):
                snippet = item.get('snippet', '')
                snippet_lower = snippet.lower()
                link = item.get('link', '').lower()
                
                # Filter out general company contact/HQ pages
                if any(term in link for term in ['/contact', '/about', '/hq', '/support']):
                    continue
                
                full_name = f"{first_name} {last_name}".lower()
                company = str(company_name).lower()
                
                # Context safeguard: ensure snippet explicitly mentions target lead and company
                if full_name in snippet_lower and (company in snippet_lower or domain.lower() in snippet_lower):
                    matches = re.findall(r'(?:\+?\d{1,3}[-.\s]?)?\(?\d{2,4}\)?[-.\s]?\d{3,4}[-.\s]?\d{3,4}', snippet)
                    valid_numbers = [m.strip() for m in matches if len(re.sub(r'\D', '', m)) >= 10]
                    
                    if valid_numbers:
                        return valid_numbers[0]
        except Exception:
            # Also sleep briefly on exception to avoid rapid retry looping
            time.sleep(random.uniform(1.0, 2.0))
            continue
            
    return "Not Exposed"


def check_and_update_daily_limit(leads_to_process_count):
    """
    Checks if processing the requested number of leads exceeds the 100-lead 24-hour limit.
    Tracks usage globally using a persistent timestamped file.
    """
    now = datetime.datetime.now()
    daily_limit = 100
    
    current_count = 0
    last_reset_time = now
    
    # Read existing tracking data
    if os.path.exists(LIMIT_TRACKER_FILE):
        try:
            with open(LIMIT_TRACKER_FILE, "r") as f:
                lines = f.read().splitlines()
                if len(lines) >= 2:
                    last_reset_time = datetime.datetime.fromisoformat(lines[0])
                    current_count = int(lines[1])
        except Exception:
            pass
            
    # Check if 24 hours have passed since the last reset
    if now - last_reset_time > datetime.timedelta(hours=24):
        current_count = 0
        last_reset_time = now
        
    # Check if adding new leads exceeds the 100-lead cap
    if current_count + leads_to_process_count > daily_limit:
        remaining_allowed = max(0, daily_limit - current_count)
        return False, current_count, remaining_allowed
        
    return True, current_count, daily_limit - current_count


def increment_usage_counter(processed_count):
    """Updates the global usage counter file after processing successfully completes."""
    now = datetime.datetime.now()
    current_count = 0
    last_reset_time = now
    
    if os.path.exists(LIMIT_TRACKER_FILE):
        try:
            with open(LIMIT_TRACKER_FILE, "r") as f:
                lines = f.read().splitlines()
                if len(lines) >= 2:
                    last_reset_time = datetime.datetime.fromisoformat(lines[0])
                    current_count = int(lines[1])
        except Exception:
            pass
            
    # Reset if 24 hours have passed
    if now - last_reset_time > datetime.timedelta(hours=24):
        current_count = 0
        last_reset_time = now
        
    new_total = current_count + processed_count
    
    with open(LIMIT_TRACKER_FILE, "w") as f:
        f.write(f"{last_reset_time.isoformat()}\n{new_total}")


def create_formatted_excel(verified_leads, output_path):
    """Saves verified leads into a styled Excel sheet with green headers and custom borders."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Verified Leads"
    
    headers = list(verified_leads[0].keys())
    ws.append(headers)
    
    for lead in verified_leads:
        ws.append([lead[h] for h in headers])
        
    header_fill = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=11, color="000000")
    thin_border = Border(left=Side(style='thin', color='D3D3D3'),
                         right=Side(style='thin', color='D3D3D3'),
                         top=Side(style='thin', color='D3D3D3'),
                         bottom=Side(style='thin', color='D3D3D3'))
                         
    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
        
    for row_num in range(2, ws.max_row + 1):
        for col_num in range(1, len(headers) + 1):
            cell = ws.cell(row=row_num, column=col_num)
            cell.font = data_font
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="left", vertical="center")
            
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 5, 16)
        
    wb.save(output_path)