import imaplib
import email
import re
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials
import time

# ----------------- CONFIG -----------------
IMAP_SERVER = "mail.b2bgrowthexpo.com"
IMAP_USER = "exhibitor@b2bgrowthexpo.com"
IMAP_PASS = "~nv&[+oJ(286"

SENDER_FILTER = "The-Business-Revival-Series-2023@showoff.asp.events"

GOOGLE_SHEET_NAME = "Expo-Sales-Management"
SHEET_TAB = "exhibitors-1"

# Google service account credentials
SERVICE_ACCOUNT_FILE = "/etc/secrets/credentials.json"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
CREDS = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# ----------------- CONNECT SHEET -----------------
client = gspread.authorize(CREDS)
sheet = client.open(GOOGLE_SHEET_NAME).worksheet(SHEET_TAB)

# ----------------- CLEAN TEXT -----------------
def clean_text(text):
    if not text:
        return ""
    text = text.replace("&nbsp;", " ").replace("\xa0", " ").strip()
    text = re.sub(r"\s+", " ", text)
    return text

# ----------------- PARSE EMAIL BODY -----------------
def parse_details(body):
    # Replace HTML tags with newlines
    text = re.sub(r"<br\s*/?>", "\n", body, flags=re.I)
    text = re.sub(r"</?(p|div|tr|td|li)[^>]*>", "\n", text, flags=re.I)
    text = re.sub(r"<[^>]+>", "", text)

    # Clean HTML entities and whitespace
    text = re.sub(r"&[a-z]+;", " ", text)
    lines = [clean_text(l) for l in text.split("\n") if clean_text(l)]

    # ============================================================
    # 📌 FORMAT 3 – Floor Plan enquiry (NEW)
    # ============================================================
    if lines and "submitted enquiry for Floor Plan" in lines[0]:
        # Expected:
        # 1: Name
        # 2: Email
        # 3: Company
        # 4: Show
        # 5: Mobile
        # 6: Website
        name_parts = lines[1].split()
        first_name = name_parts[0]
        last_name = " ".join(name_parts[1:]) if len(name_parts) > 1 else ""

        return {
            "First Name": first_name,
            "Last Name": last_name,
            "Email": lines[2] if len(lines) > 2 else "",
            "Business Name": lines[3] if len(lines) > 3 else "",
            "Which event are you interested in": lines[4] if len(lines) > 4 else "",
            "Mobile Number": lines[5] if len(lines) > 5 else "",
            "LinkedIn Profile Link": "",
            "Business linkedln page or Website": lines[6] if len(lines) > 6 else "",
        }

    # ============================================================
    # 📌 FORMAT 2 – Media Pack enquiry
    # ============================================================
    if lines and "submitted enquiry for Media Pack" in lines[0]:
        name_parts = lines[1].split()
        first_name = name_parts[0]
        last_name = " ".join(name_parts[1:]) if len(name_parts) > 1 else ""

        return {
            "First Name": first_name,
            "Last Name": last_name,
            "Email": lines[2] if len(lines) > 2 else "",
            "Business Name": lines[3] if len(lines) > 3 else "",
            "Which event are you interested in": lines[4] if len(lines) > 4 else "",
            "Mobile Number": lines[5] if len(lines) > 5 else "",
            "LinkedIn Profile Link": "",
            "Business linkedln page or Website": "",
        }

    # ============================================================
    # 📌 FORMAT 1 – Normal booking enquiry
    # ============================================================
    if lines and "would like to book a stand" in lines[0].lower():
        lines = lines[1:]

    first_name, last_name = "", ""
    if len(lines) >= 1:
        name_parts = lines[0].split()
        first_name = name_parts[0]
        last_name = " ".join(name_parts[1:]) if len(name_parts) > 1 else ""

    return {
        "First Name": first_name,
        "Last Name": last_name,
        "Email": lines[1] if len(lines) > 1 else "",
        "Business Name": lines[2] if len(lines) > 2 else "",
        "Which event are you interested in": lines[3] if len(lines) > 3 else "",
        "Mobile Number": lines[4] if len(lines) > 4 else "",
        "LinkedIn Profile Link": "",
        "Business linkedln page or Website": "",
    }

# ----------------- DUPLICATE CHECK -----------------
def get_existing_emails():
    try:
        all_emails = sheet.col_values(7)  # Email column
        return set([e.lower().strip() for e in all_emails if e])
    except Exception as e:
        print(f"❌ Error fetching existing emails: {e}")
        return set()

# ----------------- FETCH EMAILS -----------------
def fetch_emails():
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(IMAP_USER, IMAP_PASS)
    mail.select("inbox")

    status, messages = mail.search(None, f'FROM "{SENDER_FILTER}"')
    email_ids = messages[0].split()
    leads = []

    for eid in email_ids[-50:]:  # last 50 only
        _, msg_data = mail.fetch(eid, "(RFC822)")
        raw_msg = msg_data[0][1]
        msg = email.message_from_bytes(raw_msg)

        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                ctype = part.get_content_type()
                cdispo = str(part.get("Content-Disposition"))
                payload = part.get_payload(decode=True)
                if not payload:
                    continue
                text = payload.decode(errors="ignore")

                if ctype == "text/plain" and "attachment" not in cdispo:
                    body = text
                    break
                elif ctype == "text/html" and not body:
                    body = text
        else:
            body = msg.get_payload(decode=True).decode(errors="ignore")

        if body.strip():
            print("📩 Raw body preview:\n", body[:200], "\n---")
            details = parse_details(body)
            print("🔎 Parsed details:", details)
            leads.append(details)

    mail.logout()
    return leads

# ----------------- PROCESS EMAILS -----------------
def process_emails(leads):
    existing_emails = get_existing_emails()
    new_rows = []

    for details in leads:
        email_value = details["Email"].lower().strip()
        if not email_value:
            print("⚠️ No email found, skipping.")
            continue
        if email_value in existing_emails:
            print(f"⏩ Duplicate skipped: {email_value}")
            continue

        # ✅ EXACTLY 34 columns (A → AH)
        row = [
            datetime.now().strftime("%Y-%m-%d"),     # A Lead Date
            "Businessrevivalseries",                 # B Lead Source
            details.get("First Name", ""),           # C First_Name
            details.get("Last Name", ""),            # D Last Name
            details.get("Business Name", ""),        # E Company Name
            details.get("Mobile Number", ""),        # F Mobile
            email_value,                             # G Email
            details.get("Which event are you interested in", ""),  # H Show

            "",  # I Next Followup
            "",  # J Email-Count
            "",  # K Call Attempt
            "",  # L Linkedin Msg
            "",  # M WhatsApp msg count
            "",  # N Comments
            "",  # O Pitch Deck URL

            "Exhibitors_opportunity",  # P Interested for

            "",  # Q Follow-Up Count
            "",  # R Last Follow-Up Date
            "",  # S Reply Status

            "",  # T LINKEDIN-HEADLINE
            "",  # U LINKEDIN-REPLY
            "",  # V LINKEDIN-URL
            "",  # W Stand Size
            "",  # X Amount
            "",  # Y CRM Update
            "",  # Z CRM Lead ID
            "",  # AA Eventbrite Update
            "",  # AB Exhibitor MIS
            "",  # AC Welcome Email
            "",  # AD Welcome Msg
            "",  # AE Canva Update
            "",  # AF Website Update
            "",  # AG Social Media Post
            ""   # AH Payment Status
        ]

        new_rows.append(row)
        existing_emails.add(email_value)

    if new_rows:
        for row in reversed(new_rows):
            sheet.insert_row(row, 2, value_input_option="USER_ENTERED", inherit_from_before=False)
        print(f"✅ Added {len(new_rows)} new leads")
    else:
        print("ℹ️ No new leads to add")

# ----------------- RUN LOOP -----------------
if __name__ == "__main__":
    while True:
        print(f"⏱ Running email fetch at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        try:
            leads = fetch_emails()
            process_emails(leads)
        except Exception as e:
            print(f"❌ Error in run loop: {e}")
        print("💤 Sleeping for 1 hour...\n")
        time.sleep(3600)
