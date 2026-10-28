import os
import csv
import time
import random
import datetime
import streamlit as st

# Import backend & UI helper functions from utils.py
from utils import (
    get_base64_of_bin_file,
    extract_domain,
    generate_permutations,
    verify_smtp_email,
    find_exposed_phone_number,
    create_formatted_excel
)

# Page configuration
st.set_page_config(
    page_title="EIT Lead Verification App",
    page_icon="⚡",
    layout="wide"
)

# Set background wallpaper if file exists
bg_file = "email_verification_wallpaper.png"
if os.path.exists(bg_file):
    bin_str = get_base64_of_bin_file(bg_file)
    page_bg_img = f'''
    <style>
    .stApp {{
        background-image: url("data:image/png;base64,{bin_str}");
        background-size: cover;
        background-position: center;
        background-repeat: no-repeat;
        background-attachment: fixed;
    }}
    </style>
    '''
    st.markdown(page_bg_img, unsafe_allow_html=True)

# Custom High-End Modern CSS
st.markdown("""
    <style>
    /* Ensure top header allows clicks */
    header[data-testid="stHeader"] {
        background: transparent !important;
        z-index: 99998 !important;
    }

    /* Adjust main container padding */
    .block-container {
        padding-top: 1.5rem !important;
        padding-bottom: 3rem !important;
        max-width: 1400px !important;
    }

    /* Fixed & Clickable NavWise Link positioned safely to the left of the Deploy button */
    .navwise-top-link {
        position: fixed;
        top: 12px;
        right: 180px; /* Moved further left to create clear spacing from Deploy */
        background: linear-gradient(135deg, #0ea5e9, #0284c7);
        color: white !important;
        padding: 7px 16px;
        border-radius: 8px;
        font-size: 13px;
        font-weight: 700;
        text-decoration: none !important;
        box-shadow: 0 4px 15px rgba(14, 165, 233, 0.4);
        z-index: 999999 !important; /* Ensure it stays above Streamlit headers */
        pointer-events: auto !important; /* Enables mouse click interaction */
        display: inline-flex;
        align-items: center;
        gap: 6px;
        transition: all 0.2s ease-in-out;
        border: 1px solid rgba(255, 255, 255, 0.25);
    }
    .navwise-top-link:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(14, 165, 233, 0.6);
        background: linear-gradient(135deg, #0284c7, #0369a1);
        color: white !important;
    }

    /* Top-Left EIT Header Bar */
    .eit-top-header {
        position: relative;
        width: 100%;
        display: flex;
        align-items: center;
        justify-content: flex-start;
        gap: 15px;
        background: rgba(10, 15, 29, 0.92);
        backdrop-filter: blur(14px);
        -webkit-backdrop-filter: blur(14px);
        border: 1px solid rgba(255, 255, 255, 0.15);
        border-radius: 12px;
        padding: 12px 25px;
        margin-bottom: 25px;
        box-shadow: 0 4px 25px 0 rgba(0, 0, 0, 0.6);
        z-index: 9999;
    }
    .eit-logo-group {
        display: flex;
        align-items: center;
        gap: 5px;
    }
    .eit-letter {
        font-size: 18px;
        font-weight: 900;
        padding: 3px 9px;
        border-radius: 4px;
        color: white;
    }
    .eit-letter-e { background-color: #0052cc; }
    .eit-letter-i { background-color: #d92525; }
    .eit-letter-t { background-color: #0b192c; border: 1px solid rgba(255,255,255,0.2); }
    
    .eit-divider {
        width: 1px;
        height: 22px;
        background-color: rgba(255, 255, 255, 0.2);
        margin: 0 6px;
    }
    .eit-tagline {
        color: #e2e8f0;
        font-size: 14px;
        font-weight: 500;
        letter-spacing: 0.3px;
    }

    /* Engine Badge Styles */
    .engine-badge {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        background: rgba(15, 23, 42, 0.75);
        border: 1px solid rgba(56, 189, 248, 0.3);
        padding: 6px 14px;
        border-radius: 20px;
        color: #38bdf8;
        font-size: 13px;
        font-weight: 600;
        margin-bottom: 12px;
        box-shadow: 0 0 15px rgba(56, 189, 248, 0.15);
    }

    .engine-badge-sub {
        display: inline-flex;
        align-items: center;
        gap: 8px;
        background: rgba(15, 23, 42, 0.75);
        border: 1px solid rgba(56, 189, 248, 0.4);
        padding: 8px 18px;
        border-radius: 25px;
        color: #f8fafc;
        font-size: 14px;
        font-weight: 600;
        margin-top: 10px;
        box-shadow: 0 0 20px rgba(56, 189, 248, 0.2);
        text-shadow: 0 2px 8px rgba(0, 0, 0, 0.8);
    }

    /* Centered Hero Title Section */
    .hero-container {
        text-align: center;
        margin: 10px 0 30px 0;
    }
    .hero-title {
        font-size: 44px !important;
        font-weight: 900 !important;
        color: #ffffff !important;
        letter-spacing: -0.5px;
        margin-bottom: 8px !important;
        text-shadow: 0 4px 20px rgba(0, 0, 0, 0.8);
    }

    /* Transparent File Uploader Dropzone */
    [data-testid="stFileUploader"] {
        background: rgba(0, 0, 0, 0.0) !important;
        backdrop-filter: blur(4px) !important;
        -webkit-backdrop-filter: blur(4px) !important;
        border: 2px dashed rgba(56, 189, 248, 0.8) !important;
        border-radius: 35px !important;
        padding: 70px 40px !important;
        max-width: 950px !important;
        width: 80% !important;
        margin: 0 auto 35px auto !important;
        transition: all 0.3s ease !important;
        box-shadow: 0 20px 50px rgba(0, 0, 0, 0.4) !important;
    }
    
    [data-testid="stFileUploader"] section, 
    [data-testid="stFileUploader"] div, 
    [data-testid="stFileUploadDropzone"],
    [data-testid="stFileUploader"] > div {
        background: transparent !important;
        background-color: transparent !important;
        text-align: center !important;
    }

    [data-testid="stFileUploader"]:hover {
        border-color: #38bdf8 !important;
        background: rgba(56, 189, 248, 0.05) !important;
        box-shadow: 0 0 40px rgba(56, 189, 248, 0.3) !important;
    }
    
    [data-testid="stFileUploader"] label {
        color: #f8fafc !important;
        font-size: 23px !important;
        font-weight: 700 !important;
        text-align: center !important;
        display: block !important;
        width: 100% !important;
        margin-bottom: 20px !important;
        text-shadow: 0 2px 10px rgba(0, 0, 0, 0.8);
    }

    /* Action Button Styling */
    .stButton>button {
        border-radius: 12px !important;
        font-size: 16px !important;
        font-weight: 700 !important;
        padding: 12px 24px !important;
        background: linear-gradient(135deg, #2563eb, #1d4ed8) !important;
        border: 1px solid rgba(255, 255, 255, 0.2) !important;
        box-shadow: 0 4px 15px rgba(37, 99, 235, 0.4) !important;
        transition: all 0.2s ease-in-out !important;
    }
    .stButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(37, 99, 235, 0.6) !important;
    }
    </style>
""", unsafe_allow_html=True)

# --- Top NavWise Link (Spaced to the left of Deploy and Clickable) ---
st.markdown("""
    <a href="https://navwise.ai/navwise" target="_blank" class="navwise-top-link">
        🌐 NavWise Export
    </a>
""", unsafe_allow_html=True)

# --- Top-Left EIT Header Bar ---
st.markdown("""
    <div class="eit-top-header">
        <div class="eit-logo-group">
            <span class="eit-letter eit-letter-e">E</span>
            <span class="eit-letter eit-letter-i">I</span>
            <span class="eit-letter eit-letter-t">T</span>
        </div>
        <div class="eit-divider"></div>
        <div class="eit-tagline">Transforming Businesses Through IT</div>
    </div>
""", unsafe_allow_html=True)

# --- Centered Hero Section ---
st.markdown("""
    <div class="hero-container">
        <div class="engine-badge">
            ⚡ Standard B2B Lead Enrichment Engine
        </div>
        <div class="hero-title">Lead Verification App</div>
        <div class="engine-badge-sub">
            ⚡ Automated B2B Email Permutation & Real-Time SMTP Validation Engine
        </div>
    </div>
""", unsafe_allow_html=True)

# --- File Uploader Dropzone ---
uploaded_file = st.file_uploader("📂 Drag & Drop your NavWise CSV dataset here", type=["csv"])

if uploaded_file is not None:
    content = uploaded_file.getvalue().decode('utf-8').splitlines()
    reader = list(csv.DictReader(content))
    total_leads = len(reader)
    
    st.success(f"File loaded successfully! Found **{total_leads}** leads to process.")
    
    if st.button("🚀 Start Verification Process", type="primary", use_container_width=True):
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
            company_name = row.get('Organisation Name', '')
            
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
                status_text.markdown(f"📞 **Searching Phone for:** `{first_name} {last_name}`")
                phone_number = find_exposed_phone_number(first_name, last_name, company_name, domain)
                
                lead_data = {
                    'slno': sl_no,
                    'firstname': str(first_name).strip(),
                    'lastname': str(last_name).strip(),
                    'current title': str(row.get('Current Title', '')).strip(),
                    'organisation name': str(company_name).strip(),
                    'organization website': domain,
                    'location': str(row.get('Location', '')).strip(),
                    'industry': str(row.get('Industry', '')).strip(),
                    'email': found_email,
                    'phone': phone_number
                }
                verified_leads.append(lead_data)
                sl_no += 1
                
            progress_bar.progress((idx + 1) / total_leads)
            time.sleep(random.randint(2, 5))

        status_text.empty()
        time_text.empty()
        progress_bar.progress(1.0)
        
        if verified_leads:
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = f"verified_lead_{timestamp}.xlsx"
            
            create_formatted_excel(verified_leads, output_filename)
            
            st.balloons()
            st.success(f"🎉 Verification Complete! Successfully verified {len(verified_leads)} valid leads.")
            
            with open(output_filename, "rb") as file:
                st.download_button(
                    label="📥 Download Verified Leads Excel",
                    data=file,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )
        else:
            st.warning("No valid emails could be verified from the uploaded dataset.")