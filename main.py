import os
import re
import time
import subprocess
import logging
import pandas as pd
import requests
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from urllib.parse import urlparse, quote
from playwright.sync_api import sync_playwright
import coloredlogs
from datetime import datetime

GREEN = "\033[32m"
RESET = "\033[0m"

API_KEY = #
EXCEL_FILE = #
SMTP_SERVER = #
SMTP_PORT = #587
EMAIL_ADDRESS = #
EMAIL_PASSWORD = #

os.makedirs('logs', exist_ok=True)
timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
log_filename = f'logs/automation_{timestamp}.log'

coloredlogs.install(
    level='INFO',
    fmt='[%(asctime)s] [%(levelname)s] %(message)s',
    datefmt='%H:%M:%S'
)

file_handler = logging.FileHandler(log_filename)
file_handler.setFormatter(logging.Formatter('[%(asctime)s] [%(levelname)s] %(message)s'))
logging.getLogger().addHandler(file_handler)

logging.info(f"Logging started. Log file: {log_filename}")

EMAIL_REGEX = re.compile(r"^[^\s@]+@[^\s@]+\.[^\s@]+$")

def is_valid_email(email):
    if not isinstance(email, str):
        return False
    email = email.strip()
    if not email:
        return False
    return EMAIL_REGEX.match(email) is not None

def sanitize_url(raw):
    if not isinstance(raw, str):
        return ""
    raw = raw.strip()
    if not raw:
        return ""
    if not raw.startswith(("http://", "https://")):
        raw = "https://" + raw
    parsed = urlparse(raw)
    if not parsed.netloc:
        return ""
    return raw

def best_match_columns(df):
    df.columns = df.columns.astype(str).str.strip().str.lower()
    candidates = {
        "name": ["name", "full name", "contact name", "person", "lead name"],
        "website": ["website", "site", "url", "domain"],
        "email": ["email", "e-mail", "mail", "email id", "emailid", "work email"]
    }
    resolved = {}
    for key, opts in candidates.items():
        hit = next((c for c in df.columns if c in opts), None)
        if hit:
            resolved[key] = hit
    def fuzzy_find(target, deny=None):
        deny = deny or []
        for col in df.columns:
            if any(bad in col for bad in deny):
                continue
            if target in col:
                return col
        return None
    if "name" not in resolved:
        resolved["name"] = fuzzy_find("name")
    if "website" not in resolved:
        resolved["website"] = fuzzy_find("web") or fuzzy_find("site") or fuzzy_find("domain")
    if "email" not in resolved:
        resolved["email"] = fuzzy_find("email") or fuzzy_find("mail")
    missing = [k for k, v in resolved.items() if not v]
    if missing:
        raise ValueError(f"Could not find required columns: {', '.join(missing)}. Found columns: {list(df.columns)}")
    slim = df[[resolved["name"], resolved["website"], resolved["email"]]].copy()
    slim.columns = ["name", "website", "email"]
    return slim

def get_pagespeed_data(url, strategy="mobile"):
    url = sanitize_url(url)
    if not url:
        return None
    endpoint = "https://www.googleapis.com/pagespeedonline/v5/runPagespeed"
    params = {"url": url, "key": API_KEY, "strategy": strategy}
    try:
        resp = requests.get(endpoint, params=params, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        lighthouse = data["lighthouseResult"]["categories"]["performance"]
        audits = data["lighthouseResult"]["audits"]
        return {
            "score": int(lighthouse.get("score", 0) * 100),
            "fcp": audits.get("first-contentful-paint", {}).get("displayValue", "N/A"),
            "speed_index": audits.get("speed-index", {}).get("displayValue", "N/A"),
            "tti": audits.get("interactive", {}).get("displayValue", "N/A"),
            "tested_url": url
        }
    except Exception as e:
        logging.error(f"Error fetching PageSpeed data for {url}: {e}")
        return None

def install_playwright_browser():
    try:
        subprocess.run(["python", "-m", "playwright", "install", "chromium"], check=True, capture_output=True)
    except Exception as e:
        logging.error(f"Failed to install Playwright browser: {e}")

def generate_graphical_pagespeed_pdf(url, output_path, selector_timeout=90_000):
    try:
        encoded = quote(url, safe="")
        report_url = f"https://pagespeed.web.dev/report?url={encoded}"
        
        with sync_playwright() as p:
            browser = p.chromium.launch(
                headless=True, 
                args=[
                    "--no-sandbox", 
                    "--disable-setuid-sandbox",
                    "--disable-dev-shm-usage",
                    "--disable-web-security"
                ]
            )
            context = browser.new_context(
                viewport={'width': 1920, 'height': 1080},
                user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
            )
            page = context.new_page()
            
            page.goto(report_url, wait_until="networkidle", timeout=60000)
            
            try:
                page.wait_for_selector('div[data-testid="lh-score__gauge"]', timeout=30000)
                page.wait_for_selector('.lh-audit-group', timeout=30000)
                page.wait_for_selector('.lh-metrics-container', timeout=30000, state='visible')
                page.wait_for_timeout(5000)
            except Exception as e:
                logging.warning(f"Some elements didn't load completely, proceeding anyway: {e}")
            
            page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            page.wait_for_timeout(2000)
            page.evaluate("window.scrollTo(0, 0)")
            page.wait_for_timeout(1000)
            
            page.pdf(
                path=output_path,
                format="A4",
                print_background=True,
                prefer_css_page_size=True,
                margin={
                    'top': '0.5in',
                    'right': '0.5in',
                    'bottom': '0.5in',
                    'left': '0.5in'
                }
            )
            
            context.close()
            browser.close()
        return True
        
    except Exception as e:
        logging.error(f"First attempt failed for {url}: {e}")
        
        try:
            install_playwright_browser()
            
            with sync_playwright() as p:
                browser = p.chromium.launch(headless=True, args=["--no-sandbox"])
                context = browser.new_context(viewport={'width': 1920, 'height': 1080})
                page = context.new_page()
                
                page.goto(report_url, wait_until="domcontentloaded", timeout=60000)
                page.wait_for_timeout(15000)
                
                try:
                    page.wait_for_selector('.lh-root', timeout=20000)
                except:
                    pass
                
                page.pdf(
                    path=output_path,
                    format="A4",
                    print_background=True,
                    margin={'top': '0.5in', 'right': '0.5in', 'bottom': '0.5in', 'left': '0.5in'}
                )
                
                context.close()
                browser.close()
            return True
            
        except Exception as e2:
            logging.error(f"Graphical report generation failed for {url}: {e2}")
            return False

def send_email(name, email, website, report, attachment_paths):
    html_template = f'''
    <html>
      <body style="font-family: Arial, sans-serif; line-height: 1.6;">
        <p>Dear {name},</p>
        <p>Google's PageSpeed Insights score <strong>{report['score']}/100</strong> on mobile performance. </p>
        #
      </body>
    </html>
    '''

    msg = MIMEMultipart("alternative")
    msg["From"] = EMAIL_ADDRESS
    msg["To"] = email
    msg["Bcc"] = EMAIL_ADDRESS
    msg["Subject"] = f"A quick thought on the {website} website"
    msg.attach(MIMEText(html_template, "html"))

    for path in attachment_paths:
        if not path or not os.path.exists(path):
            continue
        with open(path, "rb") as f:
            part = MIMEApplication(f.read(), _subtype="pdf")
            part.add_header("Content-Disposition", "attachment", filename=os.path.basename(path))
            msg.attach(part)

    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            server.send_message(msg)
        return True
    except Exception as e:
        logging.error(f"SMTP error for {email}: {e}")
        return False


def main():
    try:
        df = pd.read_excel(EXCEL_FILE)
        df = best_match_columns(df)
    except Exception as e:
        logging.error(f"Failed to read or prepare Excel file: {e}")
        return
    df = df.dropna(subset=["email", "website"])
    df["email"] = df["email"].astype(str).str.strip()
    df["website"] = df["website"].astype(str).str.strip()
    df["name"] = df["name"].astype(str).str.strip()
    df = df[df["email"].apply(is_valid_email)]
    df = df.drop_duplicates(subset=["email"])
    if df.empty:
        logging.info("No valid leads found after cleaning.")
        return
    sent, failed, skipped = 0, 0, 0
    for idx, (_, row) in enumerate(df.iterrows(), 1):
        start_time = time.time()
        name, website, email = row["name"], row["website"], row["email"]
        if not website:
            logging.warning(f"Skipping {name} ({email}) due to missing website.")
            skipped += 1
            continue
        logging.info(f"{idx} Processing: {name} | {website} | {email}")
        report = get_pagespeed_data(website, strategy="mobile")
        if not report:
            logging.warning(f"Skipping due to missing PageSpeed data: {website}")
            skipped += 1
            continue
        safe_name = re.sub(r"[^A-Za-z0-9_\-]+", "_", name)[:40] or "lead"
        visual_pdf = f"{safe_name}_pagespeed_report.pdf"
        ok_visual = generate_graphical_pagespeed_pdf(report["tested_url"], visual_pdf)
        attachments = []
        if ok_visual and os.path.exists(visual_pdf):
            attachments.append(visual_pdf)
        if not attachments:
            logging.warning(f"No PDF generated for {name}, skipping email.")
            skipped += 1
            continue
        ok = send_email(name, email, website, report, attachments)
        elapsed = time.time() - start_time
        if ok:
            logging.info(f"{GREEN}Sent to {name} ({email}) in {elapsed:.2f} seconds.{RESET}")
            sent += 1
        else:
            logging.error(f"Failed to send to {name} ({email}) in {elapsed:.2f} seconds.")
            failed += 1
        for fpath in [visual_pdf]:
            if fpath and os.path.exists(fpath):
                try:
                    os.remove(fpath)
                except OSError:
                    pass
        time.sleep(0.8)
    logging.info(f"Completed. Sent: {sent} | Failed: {failed} | Skipped: {skipped}")

if __name__ == "__main__":
    main()
