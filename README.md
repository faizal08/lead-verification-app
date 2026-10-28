# ⚡ EIT Lead Verification App

> **Transforming Businesses Through IT** — An automated B2B lead enrichment engine built with **Streamlit** that performs intelligent email permutations, real-time SMTP mail-server validation, and direct phone number discovery.

---

## 📸 App Preview

![EIT Lead Verification App Screenshot](lead_verification_app_screenshot.png)

## 🚀 Key Features

- **Smart Email Permutations:** Automatically generates standard corporate email variations (e.g., `firstname.lastname@company.com`, `flastname@company.com`, `firstname@company.com`) from first names, last names, and company domains.
- **Real-Time SMTP Verification:** Pings target mail servers via MX record lookups to verify mailbox validity (`Valid`, `Greylisted`, or `Invalid`) before export.
- **Direct Phone Number Discovery:** Integrates with the Google Custom Search API to safely scrape and filter out general company switchboards, targeting direct personal lines.
- **Automated Excel Styling:** Automatically formats and exports verified leads into a beautifully styled `.xlsx` report using `openpyxl` (with custom header fills, data fonts, and auto-adjusted column widths).
- **Modern Custom UI:** Features a high-end transparent glassmorphic UI, custom branding elements, background wallpaper support, and live progress/ETA trackers.

---

## 🛠️ Tech Stack & Libraries

- **Python** (3.8+)
- **Streamlit** (Interactive Frontend UI)
- **OpenPyXL** (Advanced Excel formatting & generation)
- **Dnspython** (DNS MX record resolution)
- **Requests** (Google Custom Search API integration)
- **Smtplib & Sockets** (SMTP handshake validation)

---

## ⚙️ Configuration & Secrets

To enable the phone number discovery feature via the Google Custom Search API, configure your Streamlit secrets file (`.streamlit/secrets.toml`) in your root directory:

```toml
GOOGLE_API_KEY = "your_google_api_key_here"
SEARCH_ENGINE_ID = "your_custom_search_engine_id_here"
```

---

## 📦 Installation & Local Setup

### 1. Clone the Repository

```bash
git clone https://github.com/faizal08/lead-verification-app.git
cd lead-verification-app
```

### 2. Install Dependencies

```bash
pip install streamlit openpyxl dnspython requests
```

### 3. Run the Streamlit App

```bash
python -m streamlit run app.py
```

---

## 💡 How to Use

1. Upload your target lead dataset generated using Navwise in CSV format (containing columns like `First Name`, `Last Name`, `Organisation Name`, and `Organisation Website`).
2. Click **🚀 Start Verification Process**.
3. Watch the real-time SMTP status logs and estimated time remaining.
4. Download your cleaned, verified, and beautifully formatted Excel report instantly!

---

## 📧 Contact

- **Developer:** [Faizal](https://github.com/faizal08)
- **Email:** [reachfaizal08@gmail.com](mailto:reachfaizal08@gmail.com)
