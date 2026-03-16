import os, sys, time, logging, smtplib, csv, io
from email.message import EmailMessage
from email.utils import formatdate
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor
from simple_salesforce import Salesforce

# Selenium Imports
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# ================= CONFIGURATION =================
SF_USERNAME = os.getenv('SF_USERNAME')
SF_PASSWORD = os.getenv('SF_PASSWORD')
SF_TOKEN    = os.getenv('SF_TOKEN')

EMAIL_SENDER   = os.getenv('EMAIL_SENDER')
EMAIL_PASSWORD = os.getenv('EMAIL_PASSWORD')
EMAIL_RECEIVER = 'nikhil.chaudhary@loopwork.co' 

BASE_URL = 'https://loop-subscriptions.lightning.force.com/lightning/r/{obj}/{id}/view'

logging.basicConfig(level=logging.INFO, format='%(asctime)s | %(message)s')

# GLOBAL DRIVER PATH TO FIX EOF ERROR
GLOBAL_DRIVER_PATH = None

# ================= 🛠️ JAVASCRIPT LOGIC 🛠️ =================
JS_EXPAND_LOGIC = """
    (function() {
        function queryDeep(root) {
            let foundElements = [];
            let walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, null, false);
            let node;
            while (node = walker.nextNode()) {
                let txt = node.textContent.toLowerCase().trim();
                if ((txt.includes('reply') || txt.includes('replies')) && !txt.includes('collapse')) {
                    let parent = node.parentElement;
                    while (parent && parent.tagName !== 'BUTTON' && parent !== root) { parent = parent.parentElement; }
                    if (parent && parent.tagName === 'BUTTON' && parent.getAttribute('aria-pressed') === 'false') {
                        foundElements.push(parent);
                    }
                }
            }
            let all = root.querySelectorAll('*');
            for (let el of all) { if (el.shadowRoot) foundElements = foundElements.concat(queryDeep(el.shadowRoot)); }
            return foundElements;
        }
        queryDeep(document.body).forEach(btn => { try { btn.click(); } catch(e){} });
    })();
"""

JS_GET_DATES = """
    function getDates(root) {
        let res = [];
        let sels = ['.dueDate', '.slds-timeline__date', '.email-message-date'];
        sels.forEach(s => {
            root.querySelectorAll(s).forEach(el => {
                let txt = el.innerText ? el.innerText.trim() : "";
                if(txt.length > 0) res.push(txt);
            });
        });
        root.querySelectorAll('*').forEach(el => { if (el.shadowRoot) res = res.concat(getDates(el.shadowRoot)); });
        return res;
    }
    return getDates(document.body);
"""

# ================= HELPERS =================
def get_india_date_str():
    return (datetime.utcnow() + timedelta(hours=5, minutes=30)).strftime('%m/%d/%Y')

def get_india_full_timestamp():
    return (datetime.utcnow() + timedelta(hours=5, minutes=30)).strftime('%m/%d/%Y %I:%M %p (IST)')

def clean_date_to_mdy(text):
    if not text: return None
    t = text.split('|')[-1].strip().lower()
    now = datetime.now()
    if 'today' in t: return now.strftime('%m/%d/%Y')
    if 'yesterday' in t: return (now - timedelta(days=1)).strftime('%m/%d/%Y')
    t = t.replace('overdue', '').strip().title()
    for fmt in ('%d-%b-%Y', '%d-%b', '%d-%b-%y', '%m/%d/%Y'):
        try:
            dt = datetime.strptime(t, fmt)
            if fmt == '%d-%b': dt = dt.replace(year=now.year)
            return dt.strftime('%m/%d/%Y')
        except: continue
    return None

def create_html_body(title, data_rows, footer_note=""):
    rows_html = "".join([f"<tr><td style='padding:12px;border-bottom:1px solid #e0e0e0;font-weight:bold;width:50%;'>{l}</td><td style='padding:12px;border-bottom:1px solid #e0e0e0;'>{v}</td></tr>" for l, v in data_rows])
    return f"""<html><body style="font-family:'Segoe UI',Arial;background:#f9f9f9;padding:20px;"><div style="max-width:600px;margin:auto;background:#fff;padding:30px;border-radius:8px;box-shadow:0 2px 5px rgba(0,0,0,0.1);"><h2 style="color:#2c3e50;border-bottom:2px solid #3498db;padding-bottom:10px;">📊 {title}</h2><p style="font-size:14px;color:#7f8c8d;">{get_india_full_timestamp()}</p><table style="width:100%;border-collapse:collapse;">{rows_html}</table><p style="margin-top:25px;font-style:italic;color:#7f8c8d;font-size:13px;">{footer_note}</p><div style="margin-top:30px;border-top:1px solid #eee;padding-top:15px;font-size:12px;color:#999;text-align:center;">Automated by <b>Nikhil Chaudhary</b> ⚡</div></div></body></html>"""

def send_email_report(subject, html, parent_msg_id=None, csv_data=None):
    if not EMAIL_SENDER: return None
    msg = EmailMessage()
    msg['From'], msg['To'], msg['Subject'], msg['Date'] = EMAIL_SENDER, EMAIL_RECEIVER, subject, formatdate(localtime=True)
    msg.add_alternative(html, subtype='html')
    if csv_data: msg.add_attachment(csv_data.encode('utf-8'), maintype='text', subtype='csv', filename=f'App_Install_Report_{datetime.now().strftime("%m_%d_%Y")}.csv')
    if parent_msg_id: msg['In-Reply-To'] = msg['References'] = parent_msg_id
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(EMAIL_SENDER, EMAIL_PASSWORD)
        smtp.send_message(msg)
    return msg['Message-ID']

# ================= WORKER =================
def process_worker(lead_info, session_id):
    global GLOBAL_DRIVER_PATH
    lid = lead_info['Id']
    email = lead_info.get('Email', 'N/A')
    opts = Options()
    opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-gpu")
    
    driver = webdriver.Chrome(service=Service(GLOBAL_DRIVER_PATH), options=opts)
    
    report_data = {"Lead ID": lid, "Email": email, "Has Activity": "No", "Activity Count": 0, "First Activity Date": "N/A", "Latest Activity Date": "N/A"}
    try:
        driver.get(f"https://loop-subscriptions.lightning.force.com/secur/frontdoor.jsp?sid={session_id}")
        time.sleep(5)
        driver.get(BASE_URL.format(obj='Lead', id=lid))
        time.sleep(10)
        driver.execute_script(JS_EXPAND_LOGIC)
        time.sleep(5)
        raw_dates = driver.execute_script(JS_GET_DATES)
        cleaned = [clean_date_to_mdy(d) for d in raw_dates if clean_date_to_mdy(d)]
        if cleaned:
            date_objs = sorted([datetime.strptime(d, '%m/%d/%Y') for d in list(set(cleaned))])
            report_data.update({"Has Activity": "Yes", "Activity Count": len(date_objs), "First Activity Date": date_objs[0].strftime('%m/%d/%Y'), "Latest Activity Date": date_objs[-1].strftime('%m/%d/%Y')})
    except Exception as e: logging.error(f"Error on {lid}: {e}")
    finally:
        driver.quit()
        return report_data

# ================= MAIN =================
def main():
    global GLOBAL_DRIVER_PATH
    try:
        sf = Salesforce(username=SF_USERNAME, password=SF_PASSWORD, security_token=SF_TOKEN)
    except Exception as e:
        logging.error(f"SF Connection Failed: {e}"); sys.exit(1)

    # 🛠️ FIX: Pre-install driver once to avoid EOF Error
    logging.info("Initializing WebDriver...")
    GLOBAL_DRIVER_PATH = ChromeDriverManager().install()

    # 🛠️ QUERY FIX: Only App Install leads (Matches your 830 count)
    query = "SELECT Id, Email FROM Lead WHERE Sub_Source__c = 'App Install' AND CreatedDate >= LAST_N_DAYS:30 LIMIT 1000"
    recs = sf.query_all(query)['records']
    
    title = f"App Install Activity Report [{get_india_date_str()}]"
    start_info = [
        ("Total Leads (Last 30 Days)", len(recs)), 
        ("Filter Applied", "Sub_Source == 'App Install'"), 
        ("Parallel Workers", "5 Browsers"),
        ("Execution Mode", "Multi-threaded (English)")
    ]
    
    thread_id = send_email_report(title, create_html_body(title, start_info, "Extracting data for App Install leads only."))

    final_report = []
    with ThreadPoolExecutor(max_workers=5) as executor:
        futures = [executor.submit(process_worker, lead, sf.session_id) for lead in recs]
        for f in futures: final_report.append(f.result())

    output = io.StringIO()
    writer = csv.DictWriter(output, fieldnames=["Lead ID", "Email", "Has Activity", "Activity Count", "First Activity Date", "Latest Activity Date"])
    writer.writeheader()
    writer.writerows(final_report)

    summary_info = [("Total Processed", len(recs)), ("Data Format", "MM/DD/YYYY"), ("Status", "Completed")]
    send_email_report(title, create_html_body("✅ App Install Extraction Complete", summary_info, "CSV attached."), parent_msg_id=thread_id, csv_data=output.getvalue())

if __name__ == "__main__": main()
