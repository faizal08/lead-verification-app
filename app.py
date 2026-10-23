import os
import random
import time
import csv
import datetime
import dns.resolver
import smtplib
from urllib.parse import urlparse
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import streamlit as st

# Page configuration
st.set_page_config(
    page_title="EIT Lead Verification App",
    page_icon="⚡",
    layout="wide"
)

# Custom CSS for custom styling (EIT Branding & Modern UI)
st.markdown("""
    <style>
    .main-header {
        background-color: #0d1b2a;
        padding: 20px;
        border-radius: 10px;
        color: white;
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin-bottom: 25px;
    }
    .header-title {
        font-size: 28px;
        font-weight: bold;
        color: #ffffff;
        margin: 0;
    }
    .header-sub {
        font-size: 14px;
        color: #e0e1dd;
    }
    .status-card {
        background-color: #1b263b;
        padding: 15px;
        border-radius: 8px;
        color: #e0e1dd;
        font-family: monospace;
        margin-top: 15px;
    }
    </style>
""", unsafe_allow_html=True)

# Helper Functions
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
    """Generates corporate email permutations."""
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

def create_formatted_excel(verified_leads, output_path):
    """Saves verified leads into a beautifully styled Excel sheet."""
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

# --- Header UI ---
col1, col2 = st.columns([1, 4])

with col1:
    # SVG styled logo matching EIT design
    st.markdown("""
        <div style="background:#0b192c; padding:10px; border-radius:8px; text-align:center;">
            <span style="background:#0052cc; color:white; font-weight:bold; font-size:24px; padding:2px 8px;">E</span><span style="background:#d92525; color:white; font-weight:bold; font-size:24px; padding:2px 8px;">I</span><span style="background:#0b192c; color:white; font-weight:bold; font-size:24px; padding:2px 8px;">T</span>
            <div style="color:#ffffff; font-size:10px; margin-top:4px;">Transforming Businesses Through IT</div>
        </div>
    """, unsafe_allow_html=True)

with col2:
    st.title("Lead Verification App")
    st.caption("B2B Email Permutation & SMTP Validation Platform")

st.divider()

# --- Drag and Drop File Upload ---
uploaded_file = st.file_uploader("Drag and drop your NavWise CSV file here", type=["csv"])

if uploaded_file is not None:
    # Parse CSV from memory
    content = uploaded_file.getvalue().decode('utf-8').splitlines()
    reader = list(csv.DictReader(content))
    total_leads = len(reader)
    
    st.success(f"File loaded successfully! Found **{total_leads}** leads to process.")
    
    if st.button("🚀 Start Verification Process", type="primary"):
        # UI Elements for progress tracking
        progress_bar = st.progress(0)
        status_text = st.empty()
        time_text = st.empty()
        log_box = st.empty()
        
        verified_leads = []
        sl_no = 1
        start_time = time.time()
        
        for idx, row in enumerate(reader):
            first_name = row.get('First Name', '')
            last_name = row.get('Last Name', '')
            raw_website = row.get('Organisation Website', '')
            domain = extract_domain(raw_website)
            
            # Estimate remaining time
            elapsed_time = time.time() - start_time
            if idx > 0:
                avg_time_per_lead = elapsed_time / idx
                remaining_leads = total_leads - idx
                est_seconds_left = int(avg_time_per_lead * remaining_leads)
                mins, secs = divmod(est_seconds_left, 60)
                time_text.info(f"⏳ **Estimated Time Remaining:** {mins} min {secs} sec")
            else:
                time_text.info("⏳ **Estimated Time Remaining:** Calculating...")

            if not domain:
                continue
                
            permutations = generate_permutations(first_name, last_name, domain)
            found_email = None
            
            for email in permutations:
                status_text.markdown(f"🔍 **Currently Testing:** `{email}`")
                
                status = verify_smtp_email(email)
                log_box.caption(f"Tested: `{email}` ➔ Result: **{status}**")
                
                if status == "Valid":
                    found_email = email
                    break
                elif status == "Greylisted":
                    time.sleep(random.randint(5, 10))
                
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
                
            # Update Progress Bar
            progress_bar.progress((idx + 1) / total_leads)
            time.sleep(random.randint(2, 5))

        status_text.empty()
        time_text.empty()
        progress_bar.progress(1.0)
        
        if verified_leads:
            # Generate filename with timestamp: verified_lead_{timestamp}.xlsx
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = f"verified_lead_{timestamp}.xlsx"
            
            create_formatted_excel(verified_leads, output_filename)
            
            st.balloons()
            st.success(f" Verification Complete! Found {len(verified_leads)} valid leads.")
            
            # Download Button
            with open(output_filename, "rb") as file:
                st.download_button(
                    label="📥 Download Verified Leads Excel",
                    data=file,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
        else:
            st.warning("No valid emails could be verified from the uploaded dataset.")