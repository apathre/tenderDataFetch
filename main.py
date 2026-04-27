import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
import time
import gspread
from google.oauth2.service_account import Credentials
from urllib.parse import urljoin
import re
import os
import json
import subprocess



# ────────────────────────────────────────────────
# CONFIG

SHEETS_CONFIG_FILE = "sheets.json"
WORKSHEET_NAME = "TendersData"

# Limits for testing (set to None for full run)
MAX_ORGANIZATIONS_TO_PROCESS = 5
MAX_TENDERS_PER_ORG = 10

BASE_URL = "https://eprocure.gov.in/eprocure/app"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# Fixed headers with Organisation and Tender Count added at the beginning
FIXED_HEADERS = [
    "Organisation",
    "Tender Count",
    "Organization Chain",
    "Tender Reference Number",
    "Tender ID",
    "Tender Type",
    "Tender Category",
    "Form of Contract",
    "Tender Fee in Rs",
    "Tender Fee Exemption Allowed",
    "EMD Amount in Rs",
    "EMD Exemption Allowed",
    "EMD Fee Type",
    "Title",
    "Work Description",
    "NDA/Pre Qualification",
    "Tender Value in Rs",
    "Contract Type",
    "Location",
    "Product Category",
    "Sub Category",
    "Bid Validity(Days)",
    "Period of Work(Days)",
    "Pre Bid Meeting Place",
    "Pre Bid Meeting Address",
    "Pre Bid Meeting Date",
    "Bid Opening Place",
    "Published Date",
    "Bid Opening Date",
    "Document Download / Sale Start Date",
    "Document Download / Sale End Date",
    "Clarification Start Date",
    "Clarification End Date",
    "Bid Submission Start Date",
    "Bid Submission End Date",
    "Tender Inviting Authority Name",
    "Tender Inviting Authority Address",
    "Tender Detail URL"
]

NUMERIC_FIELDS = [
    "Tender Fee in Rs",
    "EMD Amount in Rs",
    "Tender Value in Rs"
]

# ────────────────────────────────────────────────

# Load SHEET_URL from sheets.json
if not os.path.exists(SHEETS_CONFIG_FILE):
    print(f"Error: {SHEETS_CONFIG_FILE} not found.")
    exit(1)

with open(SHEETS_CONFIG_FILE, 'r') as f:
    config = json.load(f)
    SHEET_URL = config.get("SHEET_URL")
    if not SHEET_URL:
        print("Error: SHEET_URL missing in sheets.json")
        exit(1)

print(f"Using SHEET_URL: {SHEET_URL}")

# ────────────────────────────────────────────────
# Load credentials from GitHub secret (required in Actions)
# ────────────────────────────────────────────────

print("Checking for SERVICE_ACCOUNT_JSON secret...")

service_account_str = os.getenv("SERVICE_ACCOUNT_JSON")

if not service_account_str:
    print("SERVICE_ACCOUNT_JSON is empty or missing in environment")
    raise ValueError(
        "ERROR: SERVICE_ACCOUNT_JSON secret is missing or empty. "
        "Please add it in GitHub repo Settings → Secrets and variables → Actions."
    )

print(f"Secret found! Length: {len(service_account_str)} characters")

try:
    service_account_info = json.loads(service_account_str)
    print("JSON parsed successfully")
except json.JSONDecodeError as e:
    print(f"JSON parse error: {str(e)}")
    raise

try:
    creds = Credentials.from_service_account_info(service_account_info, scopes=SCOPES)
    print("Credentials object created successfully")
except Exception as e:
    print(f"Error creating credentials: {str(e)}")
    raise

# Authorize gspread
client = gspread.authorize(creds)
print("Google Sheets client authorized successfully!")

# Open spreadsheet
spreadsheet = client.open_by_url(SHEET_URL)

try:
    worksheet = spreadsheet.worksheet(WORKSHEET_NAME)
    print(f"Using existing worksheet: '{WORKSHEET_NAME}'")
except gspread.exceptions.WorksheetNotFound:
    worksheet = spreadsheet.add_worksheet(title=WORKSHEET_NAME, rows=2000, cols=len(FIXED_HEADERS))
    print(f"Created new worksheet: '{WORKSHEET_NAME}'")

# ────────────────────────────────────────────────
# Rest of your functions remain unchanged
# ────────────────────────────────────────────────
#---------------------------------------
#  Chrome detection code
def get_chrome_major():
    version = subprocess.check_output(
        ["google-chrome", "--version"]
    ).decode("utf-8")
    
    match = re.search(r"(\d+)\.", version)
    return int(match.group(1))

#-------------------------------------
def init_undetected_driver(headless=True):
    options = uc.ChromeOptions()
    if headless:
        options.add_argument("--headless=new")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--window-size=1366,768")
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36")
    chrome_major = get_chrome_major()
    driver = uc.Chrome(options=options, use_subprocess=True, version_main=chrome_major)
    return driver


def page_has_captcha(driver):
    try:
        text = driver.find_element(By.TAG_NAME, "body").text.lower()
        indicators = ["captcha", "verify", "robot", "recaptcha", "not a robot", "enter the code", "image below", "security check"]
        has = any(word in text for word in indicators)
        if has:
            print("→ CAPTCHA text detected (continuing anyway)")
        return has
    except:
        return False


def save_debug_html(driver, filename):
    path = os.path.join(os.getcwd(), filename)
    with open(path, "w", encoding="utf-8") as f:
        f.write(driver.page_source)
    print(f"Debug saved: {path}")


def extract_detail_data(driver, detail_url):
    soup = BeautifulSoup(driver.page_source, "html.parser")
    data = {header: "" for header in FIXED_HEADERS}

    data["Tender Detail URL"] = detail_url

    def clean_text(text):
        text = text.replace('\xa0', ' ').strip()
        text = re.sub(r'\s+', ' ', text)
        text = text.rstrip(':').strip()
        return text

    def clean_number(text):
        text = clean_text(text)
        text = re.sub(r'[^\d.]', '', text)
        return text if text else ""

    for table in soup.find_all("table", class_="tablebg"):
        for tr in table.find_all("tr"):
            tds = tr.find_all("td")
            i = 0
            while i < len(tds):
                cell = tds[i]
                if "td_caption" in cell.get("class", []):
                    caption = clean_text(cell.get_text())
                    field_value = ""
                    if i + 1 < len(tds) and "td_field" in tds[i+1].get("class", []):
                        field_value = clean_text(tds[i+1].get_text())
                        i += 1

                    caption_lower = caption.lower()
                    mapped = False
                    for header in FIXED_HEADERS[:-1]:
                        header_lower = header.lower().replace(" in rs", "").replace("(days)", "").replace(" in \u20b9", "").strip()
                        if header_lower in caption_lower or caption_lower in header_lower:
                            if header in NUMERIC_FIELDS:
                                data[header] = clean_number(field_value)
                            else:
                                data[header] = field_value
                            mapped = True
                            break

                    if not mapped:
                        if "organisation chain" in caption_lower:
                            data["Organization Chain"] = field_value
                        elif "tender reference number" in caption_lower:
                            data["Tender Reference Number"] = field_value
                        elif "tender id" in caption_lower:
                            data["Tender ID"] = field_value
                i += 1

    work_section = soup.find(string=lambda s: s and "Work Item Details" in s)
    if work_section:
        parent = work_section.find_parent("table")
        if parent:
            for tr in parent.find_all("tr"):
                tds = tr.find_all("td")
                if len(tds) >= 2:
                    key = clean_text(tds[0].get_text())
                    value = clean_text(" ".join(td.get_text() for td in tds[1:]))
                    if key == "Title":
                        data["Title"] = value
                    elif key == "Work Description":
                        data["Work Description"] = value
                    elif key == "NDA/Pre Qualification":
                        data["NDA/Pre Qualification"] = value
                    elif "Tender Value" in key:
                        data["Tender Value in Rs"] = clean_number(value)
                    elif "Product Category" in key:
                        data["Product Category"] = value
                    elif "Sub category" in key:
                        data["Sub Category"] = value
                    elif "Bid Validity" in key:
                        data["Bid Validity(Days)"] = value
                    elif "Period Of Work" in key:
                        data["Period of Work(Days)"] = value
                    elif "Location" in key:
                        data["Location"] = value
                    elif "Pre Bid Meeting Place" in key:
                        data["Pre Bid Meeting Place"] = value
                    elif "Pre Bid Meeting Address" in key:
                        data["Pre Bid Meeting Address"] = value
                    elif "Pre Bid Meeting Date" in key:
                        data["Pre Bid Meeting Date"] = value
                    elif "Bid Opening Place" in key:
                        data["Bid Opening Place"] = value

    critical_section = soup.find(string=lambda s: s and "Critical Dates" in s)
    if critical_section:
        parent = critical_section.find_parent("table")
        if parent:
            for tr in parent.find_all("tr"):
                tds = tr.find_all("td")
                if len(tds) >= 2:
                    key = clean_text(tds[0].get_text())
                    value = clean_text(tds[1].get_text())
                    if "Published Date" in key:
                        data["Published Date"] = value
                    elif "Bid Opening Date" in key:
                        data["Bid Opening Date"] = value
                    elif "Document Download / Sale Start Date" in key:
                        data["Document Download / Sale Start Date"] = value
                    elif "Document Download / Sale End Date" in key:
                        data["Document Download / Sale End Date"] = value
                    elif "Clarification Start Date" in key:
                        data["Clarification Start Date"] = value
                    elif "Clarification End Date" in key:
                        data["Clarification End Date"] = value
                    elif "Bid Submission Start Date" in key:
                        data["Bid Submission Start Date"] = value
                    elif "Bid Submission End Date" in key:
                        data["Bid Submission End Date"] = value

    authority_section = soup.find(string=lambda s: s and "Tender Inviting Authority" in s)
    if authority_section:
        parent = authority_section.find_parent("table")
        if parent:
            for tr in parent.find_all("tr"):
                tds = tr.find_all("td")
                if len(tds) >= 2:
                    key = clean_text(tds[0].get_text())
                    value = clean_text(tds[1].get_text())
                    if "Name" in key:
                        data["Tender Inviting Authority Name"] = value
                    elif "Address" in key:
                        data["Tender Inviting Authority Address"] = value

    filled = sum(1 for v in data.values() if v.strip())
    print(f"→ Extracted {filled} meaningful fields (out of {len(FIXED_HEADERS)-1} + URL)")

    return data


def scrape():
    driver = init_undetected_driver(headless=True)
    all_tenders = []
    
    try:
        print("→ Opening organizations list page...")
        driver.get(BASE_URL + "?page=FrontEndTendersByOrganisation&service=page")
        time.sleep(6)

        page_has_captcha(driver)
        save_debug_html(driver, "debug_organizations.html")

        soup = BeautifulSoup(driver.page_source, "html.parser")

        rows = soup.find_all("tr", attrs={"id": lambda x: x and "informal" in x.lower()})

        if len(rows) < 3:
            print("→ Primary selector found few rows → trying fallback")
            rows = []
            for tr in soup.select("table tr"):
                tds = tr.find_all("td")
                if len(tds) >= 3 and tds[1].get_text(strip=True):
                    rows.append(tr)

        print(f"→ Found {len(rows)} potential organization rows")

        processed_count = 0

        for row in rows:
            tds = row.find_all("td")
            if len(tds) < 3:
                continue

            try:
                org_name = tds[1].get_text(strip=True)
                count_cell = tds[2]
                tender_count = count_cell.get_text(strip=True).strip()
                link_tag = count_cell.find("a", href=True)

                if not link_tag:
                    continue

                org_url = urljoin(BASE_URL, link_tag["href"])

                print(f"  Processing → {org_name}  ({tender_count})")
                processed_count += 1

                if MAX_ORGANIZATIONS_TO_PROCESS and processed_count >= MAX_ORGANIZATIONS_TO_PROCESS:
                    print(f"→ Reached organization limit ({MAX_ORGANIZATIONS_TO_PROCESS})")
                    break

                driver.get(org_url)
                time.sleep(5)
                page_has_captcha(driver)

                tender_processed = 0
                while True:
                    soup = BeautifulSoup(driver.page_source, "html.parser")
                    tender_rows = soup.find_all("tr", attrs={"id": lambda x: x and "informal" in x.lower()})

                    if len(tender_rows) < 3:
                        tender_rows = []
                        for tr in soup.select("table tr"):
                            cells = tr.find_all("td")
                            if len(cells) >= 6 and cells[4].find("a"):
                                tender_rows.append(tr)

                    print(f"  → Found {len(tender_rows)} tender rows")

                    for trow in tender_rows:
                        cells = trow.find_all("td")
                        if len(cells) < 6:
                            continue

                        title_cell = cells[4]
                        title = title_cell.get_text(strip=True).strip()
                        detail_link = title_cell.find("a")
                        if not detail_link:
                            continue

                        detail_url = urljoin(BASE_URL, detail_link["href"])

                        print(f"     → {title[:70]}...")

                        driver.get(detail_url)
                        time.sleep(5)
                        page_has_captcha(driver)

                        detail_data = extract_detail_data(driver, detail_url)

                        tender = {
                            "Organisation": org_name,
                            "Tender Count": tender_count,
                            **detail_data
                        }

                        all_tenders.append(tender)

                        tender_processed += 1
                        if MAX_TENDERS_PER_ORG and tender_processed >= MAX_TENDERS_PER_ORG:
                            print(f"  → Reached per-org limit ({MAX_TENDERS_PER_ORG})")
                            break

                        driver.back()
                        time.sleep(2)

                    if MAX_TENDERS_PER_ORG and tender_processed >= MAX_TENDERS_PER_ORG:
                        break

                    try:
                        next_btn = driver.find_element(
                            By.XPATH,
                            "//a[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'next')]"
                        )
                        if "disabled" in (next_btn.get_attribute("class") or "") or not next_btn.is_enabled():
                            break
                        next_btn.click()
                        time.sleep(5)
                    except:
                        break

                time.sleep(3)

            except Exception as e:
                print(f"  Error: {str(e)}")
                continue

        print(f"\nCollected {len(all_tenders)} tenders")
        print(f"Processed {processed_count} organizations")

    except KeyboardInterrupt:
        print("→ Interrupted by user")
    except Exception as e:
        print("Error:", str(e))
    finally:
        try:
            driver.quit()
        except:
            pass

    return all_tenders


def append_new_tenders_to_sheet(tenders):
    if not tenders:
        print("No tenders collected.")
        return

    current_headers = worksheet.row_values(1)
    if not current_headers or current_headers != FIXED_HEADERS:
        worksheet.clear()
        worksheet.append_row(FIXED_HEADERS)
        print(f"→ Headers set to {len(FIXED_HEADERS)} columns")

    existing_data = worksheet.get_all_values()
    existing_ids = set()
    try:
        id_col = FIXED_HEADERS.index("Tender ID")
        existing_ids = {row[id_col] for row in existing_data[1:] if len(row) > id_col and row[id_col].strip()}
    except ValueError:
        print("→ Warning: Tender ID column not found for deduplication")

    new_rows = []
    for tender in tenders:
        tender_id = tender.get("Tender ID", "").strip()
        if not tender_id:
            print(f"→ Skipping row without Tender ID")
            continue
        if tender_id in existing_ids:
            print(f"→ Skipping duplicate: {tender_id}")
            continue

        row = [tender.get(h, "") for h in FIXED_HEADERS]
        new_rows.append(row)

    if new_rows:
        worksheet.append_rows(new_rows)
        print(f"→ Added {len(new_rows)} new rows to '{WORKSHEET_NAME}'")
    else:
        print("→ No new tenders (all duplicates or missing Tender ID)")


if __name__ == "__main__":
    print("Starting refined extraction...")
    collected_tenders = scrape()
    append_new_tenders_to_sheet(collected_tenders)
    print("Finished.")
