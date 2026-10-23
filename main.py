import os
import random
import time
import csv
import dns.resolver
import smtplib
from urllib.parse import urlparse
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def extract_domain(url):
    """Extracts clean domain from organisation website URL."""
    if not url or str(url).strip() == '' or str(url).lower() == 'nan':
        return ""
    parsed = urlparse(str(url))
    domain = parsed.netloc if parsed.netloc else parsed.path
    domain = domain.split('/')[0]
    if domain.startswith('www.'):
        domain = domain[4:]
    return domain.strip().lower()

def generate_permutations(first_name, last_name, domain):
    """Generates the 3 standard corporate email permutations."""
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
        f"{f[0]}{l}@{d}",  # firstinitial.lastname@company.com (flastname)
        f"{f}@{d}"         # firstname@company.com
    ]

def verify_smtp_email(email, sender_email="verify@gmail.com"):
    """Pings mail server via SMTP with MX record lookup."""
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

def process_leads(input_csv, output_xlsx):
    print(f"[*] Loading leads from {input_csv}...")
    if not os.path.exists(input_csv):
        print(f"[!] Error: {input_csv} not found in directory.")
        return
        
    verified_leads = []
    sl_no = 1
    
    # Read CSV using built-in csv module (No pandas required)
    with open(input_csv, mode='r', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for row in reader:
            first_name = row.get('First Name', '')
            last_name = row.get('Last Name', '')
            raw_website = row.get('Organisation Website', '')
            domain = extract_domain(raw_website)
            
            if not domain:
                continue
                
            permutations = generate_permutations(first_name, last_name, domain)
            found_email = None
            
            for email in permutations:
                print(f"[*] Testing: {email} ...", end=" ")
                status = verify_smtp_email(email)
                print(f"[{status}]")
                
                if status == "Valid":
                    found_email = email
                    break
                elif status == "Greylisted":
                    # Handle greylisting gracefully by waiting longer before next attempt
                    time.sleep(random.randint(5, 10))
                
                # Essential Safety Rule: Randomized delay between checks (3 to 8 seconds)
                time.sleep(random.randint(3, 8))
                
            if found_email:
                lead_data = {
                    'slno': sl_no,
                    'firstname': str(first_name).strip(),
                    'lastname': str(last_name).strip(),
                    'current title': str(row.get('Current Title', '')).strip(),
                    'organisation name': str(row.get('Organisation Name', '')).strip(),
                    'organization website': domain,
                    'location': str(row.get('Location', '')).strip(),
                    'industry': str(row.get('Industry', '')).strip(),
                    'seniority': str(row.get('Seniority', '')).strip(),
                    'email': found_email
                }
                verified_leads.append(lead_data)
                sl_no += 1
                
            # Throttling between leads to protect IP reputation
            time.sleep(random.randint(2, 5))

    if not verified_leads:
        print("[!] No verified emails found.")
        return

    # Save to Excel and apply professional formatting using openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Verified Leads"
    
    headers = list(verified_leads[0].keys())
    ws.append(headers)
    
    for lead in verified_leads:
        ws.append([lead[h] for h in headers])
        
    # Formatting styles
    header_fill = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid") # Dark Green
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF") # White Bold
    data_font = Font(name="Calibri", size=11, color="000000") # Black
    thin_border = Border(left=Side(style='thin', color='D3D3D3'),
                         right=Side(style='thin', color='D3D3D3'),
                         top=Side(style='thin', color='D3D3D3'),
                         bottom=Side(style='thin', color='D3D3D3'))
                         
    # Apply header formatting
    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
        
    # Apply data formatting
    for row_num in range(2, ws.max_row + 1):
        for col_num in range(1, len(headers) + 1):
            cell = ws.cell(row=row_num, column=col_num)
            cell.font = data_font
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="left", vertical="center")
            
    # Set proper column widths so content is fully visible without dragging
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 5, 16)
        
    wb.save(output_xlsx)
    print(f"\n[+] Success! {len(verified_leads)} verified leads saved and formatted in '{output_xlsx}'")

if __name__ == "__main__":
    input_file = "sample leads.csv"
    output_file = "verified_leads.xlsx"
    process_leads(input_file, output_file)