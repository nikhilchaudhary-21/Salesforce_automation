import os, sys, time, logging, smtplib, csv, io
from email.message import EmailMessage
from email.utils import make_msgid, formatdate
from datetime import datetime, timedelta
from collections import Counter
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
SALES_API_DATE = 'Last_Activity_Date_V__c'

logging.basicConfig(level=logging.INFO, format='%(asctime)s | %(message)s')

GLOBAL_DRIVER_PATH = None

# ================= 🛠️ JAVASCRIPT LOGIC (TEXT WALKER) 🛠️ =================
JS_EXPAND_LOGIC = """
    (function() {
        console.log("🚀 Starting Universal Text Walker...");
        function triggerClick(el) {
            if (!el) return;
            try {
                el.scrollIntoView({block: 'center'});
                el.click();
                let eventOpts = {bubbles: true, cancelable: true, view: window};
                el.dispatchEvent(new MouseEvent('mousedown', eventOpts));
                el.dispatchEvent(new MouseEvent('mouseup', eventOpts));
                el.dispatchEvent(new MouseEvent('click', eventOpts));
                console.log("⚡ Clicked:", el.innerText);
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

# ================= 🎨 HTML TEMPLATE 🎨 =================
def create_html_body(title, data_rows, footer_note=""):
    rows_html = "".join([f"""
        <tr>
            <td style="padding: 12px; border-bottom: 1px solid #e0e0e0; font-weight: bold; color: #333; width: 40%;">{l}</td>
            <td style="padding: 12px; border-bottom: 1px solid #e0e0e0; color: #555;">{v}</td>
        </tr>""" for l, v in data_rows])
    
    return f"""
    <html>
    <body style="font-family: 'Segoe UI', Arial, sans-serif; color: #333; background-color: #f9f9f9; padding: 20px;">
        <div style="max-width: 600px; margin: auto; background: #fff; padding: 30px; border-radius: 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.1);">
            <h2 style="color: #2c3e50; margin-top: 0; border-bottom: 2px solid #3498db; padding-bottom: 10px;">📊 {title}</h2>
            <p style="font-size: 14px; color: #7f8c8d; margin-bottom: 20px;">{get_india_full_timestamp()}</p>
            <table style="width: 100%; border-collapse: collapse;">{rows_html}</table>
            <p style="margin-top: 25px; font-style: italic; color: #7f8c8d; font-size: 13px;">{footer_note}</p>
            <div style="margin-top: 30px; border-top: 1px solid #eee; padding-top: 15px; font-size: 12px; color: #999; text-align: center;">
                Automated by <b>Nikhil Chaudhary</b> ⚡
            </div>
        </div>
    </body>
    </html>"""

def send_email_report(subject, html, parent_msg_id=None, csv_data=None):
    if not EMAIL_SENDER or not EMAIL_RECEIVER: return None
    
    msg = EmailMessage()
    
    # Logic for multiple recipients
    recipients = [email.strip() for email in EMAIL_RECEIVER.split(',')]
    
    msg['From'] = EMAIL_SENDER
    msg['To'] = ", ".join(recipients)
    msg['Subject'] = subject
    msg['Date'] = formatdate(localtime=True)
    
    msg.add_alternative(html, subtype='html')
    
    if csv_data: 
        msg.add_attachment(csv_data.encode('utf-8'), maintype='text', subtype='csv', filename='sales_errors.csv')
    
    if parent_msg_id: 
        msg['In-Reply-To'] = msg['References'] = parent_msg_id
        
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(EMAIL_SENDER, EMAIL_PASSWORD)
        smtp.send_message(msg)
    return msg['Message-ID']

# ================= WORKER FOR THREADING =================
def process_account_worker(rec_id, session_id):
    global GLOBAL_DRIVER_PATH
    chrome_options = Options()
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")
    
    driver = webdriver.Chrome(service=Service(GLOBAL_DRIVER_PATH), options=chrome_options)
    
    try:
        driver.get(f"https://loop-subscriptions.lightning.force.com/secur/frontdoor.jsp?sid={session_id}")
        time.sleep(5)
        url = BASE_URL.format(obj='Account', id=rec_id)
        driver.get(url)
        time.sleep(10)
        for _ in range(3):
            driver.execute_script(JS_EXPAND_LOGIC)
            time.sleep(3)
        cutoff_y = driver.execute_script(JS_GET_CUTOFF)
        raw_items = driver.execute_script(JS_GET_DATES)
        valid_dates = [clean_activity_date(i['text']) for i in raw_items if (cutoff_y == 0 or i['y'] >= (cutoff_y - 10)) and clean_activity_date(i['text'])]
        if not valid_dates: return rec_id, None, None
        valid_dates.sort(key=lambda x: datetime.strptime(x, '%d-%b-%Y'), reverse=True)
        return rec_id, valid_dates[0], None
    except Exception as e:
        return rec_id, None, str(e)
    finally:
        driver.quit()

# ================= MAIN EXECUTION =================
def main():
    global GLOBAL_DRIVER_PATH
    try:
        sf = Salesforce(username=SF_USERNAME, password=SF_PASSWORD, security_token=SF_TOKEN)
    except Exception as e:
        logging.error(f"SF Connection Failed: {e}"); sys.exit(1)

    target_owners = "('Harshit Gupta', 'Abhishek Nayak', 'Deepesh Dubey', 'Prashant Jha', 'Mihir Hardiya')"
    sales_recs = sf.query_all(f"SELECT Id, Owner.Name FROM Account WHERE Owner.Name IN {target_owners}")['records']
    
    counts = Counter([r['Owner']['Name'] for r in sales_recs])
    breakdown = "".join([f"• {owner}: <b>{count}</b><br>" for owner, count in counts.items()])
    
    base_subject = f"Sales Account Activity Report [{get_india_date_str()}]"
    data = [
        ("Date", get_india_full_timestamp()),
        ("Sales Accounts Found", f"{len(sales_recs)} Accounts"),
        ("Sales Breakdown", breakdown)
    ]
    
    thread_id = send_email_report(base_subject, create_html_body(base_subject, data, "The automation script has started. Processing with 4 parallel browsers..."))

    logging.info("Downloading Chrome Driver...")
    GLOBAL_DRIVER_PATH = ChromeDriverManager().install()

    all_results = []
    with ThreadPoolExecutor(max_workers=4) as executor:
        futures = [executor.submit(process_account_worker, rec['Id'], sf.session_id) for rec in sales_recs]
        for f in futures:
            all_results.append(f.result())

    stats = {'updated': 0, 'failed': 0}
    failed_log = []

    for rid, last_date, err in all_results:
        if last_date:
            try:
                sf.Account.update(rid, {SALES_API_DATE: convert_date_for_api(last_date)})
                stats['updated'] += 1
            except Exception as ue:
                stats['failed'] += 1
                failed_log.append(['Account', rid, str(ue)])
        elif err:
            stats['failed'] += 1
            failed_log.append(['Account', rid, str(err)])

    csv_str = None
    if failed_log:
        output = io.StringIO()
        writer = csv.writer(output)
        writer.writerow(['Type', 'Record ID', 'Error'])
        writer.writerows(failed_log)
        csv_str = output.getvalue()

    final_html = create_html_body("✅ Sales Execution Complete", [
        ("Total Processed", len(sales_recs)),
        ("Successfully Updated", stats['updated']),
        ("Failed", stats['failed'])
    ], "Check the attached CSV if there are failures.")
    
    send_email_report(base_subject, final_html, parent_msg_id=thread_id, csv_data=csv_str)

if __name__ == "__main__":
    main()
