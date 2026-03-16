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
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# ================= CONFIGURATION =================
SF_USERNAME = os.getenv('SF_USERNAME')
SF_PASSWORD = os.getenv('SF_PASSWORD')
SF_TOKEN    = os.getenv('SF_TOKEN')

# Email Config
EMAIL_SENDER   = os.getenv('EMAIL_SENDER')
EMAIL_PASSWORD = os.getenv('EMAIL_PASSWORD')
EMAIL_RECEIVER = os.getenv('EMAIL_RECEIVER')

BASE_URL = 'https://loop-subscriptions.lightning.force.com/lightning/r/{obj}/{id}/view'
MKT_API_COUNT = 'Count_of_Activities__c'
MKT_API_DATE  = 'Last_Activity_Date__c'

logging.basicConfig(level=logging.INFO, format='%(asctime)s | %(message)s')

# Global variable to store driver path to avoid re-downloading in threads
GLOBAL_DRIVER_PATH = None

# ================= 🛠️ JAVASCRIPT LOGIC (ORIGINAL) 🛠️ =================
JS_EXPAND_LOGIC = """
    (function() {
        function triggerClick(el) {
            if (!el) return;
            try {
                el.scrollIntoView({block: 'center'});
                el.click();
                let eventOpts = {bubbles: true, cancelable: true, view: window};
                el.dispatchEvent(new MouseEvent('mousedown', eventOpts));
                el.dispatchEvent(new MouseEvent('mouseup', eventOpts));
                el.dispatchEvent(new MouseEvent('click', eventOpts));
            } catch(e) { console.error(e); }
        }

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

        let targets = queryDeep(document.body);
        targets.forEach(btn => triggerClick(btn));
        
        let others = document.body.querySelectorAll('button');
        others.forEach(btn => {
            let t = (btn.innerText || "").toLowerCase();
            if(t.includes('show all') || t.includes('view more') || t.includes('email body')) { btn.click(); }
        });
    })();
"""

JS_GET_CUTOFF = """
    function getCutoff(root) {
        let markers = Array.from(root.querySelectorAll('.slds-timeline__date'));
        root.querySelectorAll('*').forEach(el => { if (el.shadowRoot) markers = markers.concat(getCutoff(el.shadowRoot)); });
        return markers;
    }
    let all = getCutoff(document.body);
    all.sort((a, b) => a.getBoundingClientRect().top - b.getBoundingClientRect().top);
    if(all.length > 0) return all[0].getBoundingClientRect().top + window.scrollY;
    return 0;
"""

JS_GET_DATES = """
    function getDates(root) {
        let res = [];
        let sels = ['.dueDate', '.slds-timeline__date', '.email-message-date'];
        sels.forEach(s => {
            root.querySelectorAll(s).forEach(el => {
                let txt = el.innerText ? el.innerText.trim() : "";
                if(txt.length > 0) {
                    res.push({ text: txt, y: el.getBoundingClientRect().top + window.scrollY });
                }
            });
        });
        root.querySelectorAll('*').forEach(el => { if (el.shadowRoot) res = res.concat(getDates(el.shadowRoot)); });
        return res;
    }
    return getDates(document.body);
"""

# ================= HELPER FUNCTIONS =================
def get_india_date_str():
    return (datetime.utcnow() + timedelta(hours=5, minutes=30)).strftime('%d-%b-%Y')

def get_india_full_timestamp():
    return (datetime.utcnow() + timedelta(hours=5, minutes=30)).strftime('%d-%b-%Y %I:%M %p (IST)')

def clean_activity_date(text):
    if not text: return None
    text = text.split('|')[-1].strip()
    text_lower = text.lower()
    now = datetime.now()
    if 'today' in text_lower: return now.strftime('%d-%b-%Y')
    elif 'yesterday' in text_lower: return (now - timedelta(days=1)).strftime('%d-%b-%Y')
    if 'overdue' in text_lower: text = text_lower.replace('overdue', '').strip().title()
    for fmt in ('%d-%b-%Y', '%d-%b'):
        try:
            dt = datetime.strptime(text, fmt)
            if fmt == '%d-%b': dt = dt.replace(year=now.year)
            return dt.strftime('%d-%b-%Y')
        except: continue
    return None

def convert_date_for_api(date_str):
    if not date_str: return None
    try: return datetime.strptime(date_str, '%d-%b-%Y').strftime('%Y-%m-%d')
    except: return None

def create_html_body(title, data_rows, footer_note=""):
    rows_html = "".join([
        f"<tr>"
        f"<td style='padding:12px;border-bottom:1px solid #e0e0e0;font-weight:bold;'>{row[0]}</td>"
        f"<td style='padding:12px;border-bottom:1px solid #e0e0e0;'>{row[1]}</td>"
        f"<td style='padding:12px;border-bottom:1px solid #e0e0e0;color:#7f8c8d;font-size:12px;'>{row[2]}</td>"
        f"</tr>" for row in data_rows
    ])
    return f"""
    <html>
    <body style="font-family: 'Segoe UI', Arial, sans-serif; color: #333; background-color: #f9f9f9; padding: 20px;">
        <div style="max-width: 650px; margin: auto; background: #fff; padding: 30px; border-radius: 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.1);">
            <h2 style="color: #2c3e50; margin-top: 0; border-bottom: 2px solid #3498db; padding-bottom: 10px;">📊 {title}</h2>
            <p style="font-size: 14px; color: #7f8c8d; margin-bottom: 20px;">{get_india_full_timestamp()}</p>
            <table style="width: 100%; border-collapse: collapse;">
                <tr style="background:#f2f2f2;">
                    <th style="padding:12px;text-align:left;border-bottom:2px solid #ddd;">Field</th>
                    <th style="padding:12px;text-align:left;border-bottom:2px solid #ddd;">Detail</th>
                    <th style="padding:12px;text-align:left;border-bottom:2px solid #ddd;">Reason</th>
                </tr>
                {rows_html}
            </table>
            <p style="margin-top: 25px; font-style: italic; color: #7f8c8d; font-size: 13px;">{footer_note}</p>
            <div style="margin-top: 30px; border-top: 1px solid #eee; padding-top: 15px; font-size: 12px; color: #999; text-align: center;">Automated by <b>Nikhil Chaudhary</b> ⚡</div>
        </div>
    </body>
    </html>"""

def send_email_report(subject, html, parent_msg_id=None, csv_data=None):
    if not EMAIL_SENDER: return None
    msg = EmailMessage()
    msg['From'], msg['To'], msg['Subject'], msg['Date'] = EMAIL_SENDER, EMAIL_RECEIVER, subject, formatdate(localtime=True)
    msg.add_alternative(html, subtype='html')
    if csv_data: msg.add_attachment(csv_data.encode('utf-8'), maintype='text', subtype='csv', filename='mkt_report_details.csv')
    if parent_msg_id: msg['In-Reply-To'] = msg['References'] = parent_msg_id
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(EMAIL_SENDER, EMAIL_PASSWORD)
        smtp.send_message(msg)
    return msg['Message-ID']

# ================= WORKER FOR THREADING =================
def process_lead_worker(lid, session_id):
    global GLOBAL_DRIVER_PATH
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    driver = webdriver.Chrome(service=Service(GLOBAL_DRIVER_PATH), options=options)
    try:
        driver.get(f"https://loop-subscriptions.lightning.force.com/secur/frontdoor.jsp?sid={session_id}")
        time.sleep(5)
        url = BASE_URL.format(obj='Lead', id=lid)
        driver.get(url)
        time.sleep(10)
        for _ in range(3):
            driver.execute_script(JS_EXPAND_LOGIC)
            time.sleep(3)
        cutoff_y = driver.execute_script(JS_GET_CUTOFF)
        raw_items = driver.execute_script(JS_GET_DATES)
        valid_dates = [clean_activity_date(i['text']) for i in raw_items if (cutoff_y == 0 or i['y'] >= (cutoff_y - 10)) and clean_activity_date(i['text'])]
        if not valid_dates: return lid, 0, None, None
        valid_dates.sort(key=lambda x: datetime.strptime(x, '%d-%b-%Y'), reverse=True)
        return lid, len(set(valid_dates)), valid_dates[0], None
    except Exception as e: return lid, 0, None, str(e)
    finally: driver.quit()

# ================= MAIN EXECUTION =================
def main():
    global GLOBAL_DRIVER_PATH
    try:
        sf = Salesforce(username=SF_USERNAME, password=SF_PASSWORD, security_token=SF_TOKEN)
    except Exception as e: logging.error(f"SF Connection Failed: {e}"); sys.exit(1)

    logging.info("Downloading Chrome Driver...")
    GLOBAL_DRIVER_PATH = ChromeDriverManager().install()

    start_dt = (datetime.now() - timedelta(days=30)).strftime('%Y-%m-%dT00:00:00Z')
    mkt_query = f"SELECT Id FROM Lead WHERE LeadSource = 'Marketing Inbound' AND Sub_Source__c != 'App Install' AND CreatedDate >= {start_dt} LIMIT 400"
    mkt_recs = [r['Id'] for r in sf.query_all(mkt_query)['records']]
    
    total_leads = len(mkt_recs)
    base_subject = f"Salesforce Daily Activity Report [{get_india_date_str()}]"
    
    # Starting Email Info
    start_info = [("Total Leads Found", total_leads, "N/A"), ("Parallel Workers", "4 Browsers", "N/A")]
    thread_id = send_email_report(base_subject, create_html_body(base_subject, start_info, "Processing non-app install tasks across 4 parallel browser instances..."))

    all_details = []
    with ThreadPoolExecutor(max_workers=4) as executor:
        futures = [executor.submit(process_lead_worker, lid, sf.session_id) for lid in mkt_recs]
        for f in futures: all_details.append(f.result())

    stats = {'updated': 0, 'skipped': 0, 'failed': 0}
    csv_rows = []

    for lid, count, last_date, err in all_details:
        if last_date and not err:
            try:
                payload = {MKT_API_COUNT: count, MKT_API_DATE: convert_date_for_api(last_date)}
                sf.Lead.update(lid, payload, headers={'Sforce-Auto-Assign': 'FALSE'})
                stats['updated'] += 1
                csv_rows.append([lid, count, last_date, "Success", "Synced to SF"])
            except Exception as ue:
                stats['failed'] += 1
                csv_rows.append([lid, 0, None, "Failed", str(ue)])
        elif err:
            stats['failed'] += 1
            csv_rows.append([lid, 0, None, "Failed", str(err)])
        else:
            stats['skipped'] += 1
            csv_rows.append([lid, 0, None, "Skipped", "No Activity Found"])

    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(['Lead ID', 'Activity Count', 'Last Date', 'Status', 'Reason'])
    writer.writerows(csv_rows)
    
    final_info = [
        ("Total Processed", total_leads, "N/A"),
        ("Updated", stats['updated'], "Successfully synced"),
        ("Skipped", stats['skipped'], "No Activity Found"),
        ("Failed", stats['failed'], "Execution Errors")
    ]
    
    send_email_report(base_subject, create_html_body("✅ Marketing Execution Complete", final_info, "Check CSV for individual details."), parent_msg_id=thread_id, csv_data=output.getvalue())

if __name__ == "__main__": main()
