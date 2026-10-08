"""
Bulk email sender for Many receipients announcement.
Reads recipient list from Excel and sends personalized emails via SMTP,
with batching + delays to avoid "deferred" / rate-limit errors from mass sending.

Includes:
- Email format validation (skips malformed addresses before sending)
- A circuit breaker that pauses (not stops) the run if too many failures
  happen in a short window, then resumes automatically after a cooldown
- Resume support: already-SENT recipients are skipped on any re-run
- Incremental logging: every row is written to disk immediately
- Bounce checking via IMAP: reads your own mailbox for bounce/NDR
  notifications and reconciles them against the log, and remembers
  permanently-dead addresses so future runs never retry them

IMPORTANT ABOUT "SENT" STATUS:
When this script logs a recipient as SENT, that only means your outgoing
mail server ACCEPTED the message for delivery. It does NOT mean the message
actually reached the recipient's inbox. If the recipient address doesn't
exist, the failure is reported back later as a bounce email delivered to
YOUR OWN mailbox - Python has no way to see this at the moment sendmail()
is called. That's why bounce checking (below) exists: it is the only way
to catch these after-the-fact failures.

BEFORE RUNNING:
1. Fill in SMTP_HOST, SMTP_PORT, SENDER_EMAIL, SENDER_PASSWORD below.
   NOTE: double-check SMTP_HOST / IMAP_HOST with your mail host - these
   should almost certainly be the same hostname.
2. Make sure your Excel file has columns named "surname" and "email"
   (edit COL_NAME / COL_EMAIL below if your headers differ).
3. Place the logo file (logo_name.jpg) in the same folder as this script.
4. Install dependencies:  pip install pandas openpyxl
"""

import smtplib
import imaplib
import email as email_lib
import ssl
import time
import csv
import re
from collections import deque
from pathlib import Path
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

import pandas as pd

# ----------------------------------------------------------------------
# CONFIGURATION - EDIT THESE VALUES
# ----------------------------------------------------------------------

base_dir = Path(__file__).parent

EXCEL_FILE = base_dir / "your_excel_file.xlsx"   # <-- path to your Excel file
COL_NAME = "surname"                      # <-- column with recipient surname
COL_EMAIL = "email"                       # <-- column with email address

# CONFIRM with your host: SMTP_HOST and IMAP_HOST are almost always the
# whatever your host (hostname.com) actually confirms.
SMTP_HOST = "mail.yourdomain_name.com"
SMTP_PORT = 465                # 465 = SSL, 587 = STARTTLS (adjust to match)
USE_SSL = True                 # True for port 465, False for 587 (STARTTLS)

SENDER_EMAIL = "email@yourdomainname.com"
SENDER_PASSWORD = "password_of_email@yourdomain_name.com"   # <-- fill in (better: read from env var)

SENDER_NAME = "The Name That Will in the Receipient's Inbox" #<-- Name of Company or person sending the email
SUBJECT = "Subject of the Mail"

# Throttling settings (tune these if you still see deferrals)
BATCH_SIZE = 20                 # emails per SMTP connection before reconnecting
DELAY_BETWEEN_EMAILS = 3        # seconds to wait after each email
DELAY_BETWEEN_BATCHES = 30      # seconds to wait after each batch (reconnect happens here)
MAX_RETRIES = 2                 # retry attempts per email on failure

# Circuit breaker: pause the run if too many failures happen in a short
# window (this is what protects you from tripping your host's hourly
# "max defers and failures" lock).
CIRCUIT_BREAKER_MAX_FAILURES = 3      # failures allowed...
CIRCUIT_BREAKER_WINDOW_SECONDS = 600  # ...within this many seconds (10 min)

# Instead of stopping for good, wait this long then resume automatically.
CIRCUIT_BREAKER_COOLDOWN_SECONDS = 65 * 60

# Safety cap: stop for good (not just pause) after this many cooldown
# cycles, so a genuinely broken list can't loop unattended indefinitely.
MAX_COOLDOWN_CYCLES = 5

# Basic email format check (not a guarantee the mailbox exists, but catches
# typos and malformed addresses before they ever reach your SMTP server)
EMAIL_REGEX = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

# --- Bounce checking (IMAP) ---
ENABLE_BOUNCE_CHECKING = True
IMAP_HOST = "Your SMTP Host"   # should match SMTP_HOST's domain - confirm with host
IMAP_PORT = 993                  # 993 = IMAP over SSL (standard)
BOUNCE_CHECK_EVERY_N = BATCH_SIZE

LOG_FILE = base_dir / "send_log.csv"

# Permanent record of addresses confirmed dead by a bounce, across all runs.
DEAD_ADDRESSES_FILE = base_dir / "dead_addresses.csv"

# The STC crest logo, placed next to this script
LOGO_FILE = base_dir / "logo_name.jpg"
LOGO_CID = "logo_image"

# Read the logo once at startup (not per-email - keeps mass sending fast)
with open(LOGO_FILE, "rb") as f:
    LOGO_BYTES = f.read()

# ----------------------------------------------------------------------
# EMAIL BODY TEMPLATE
# ----------------------------------------------------------------------

REGISTRATION_LINK = "Any link you want to put in the email body"


def build_message(name: str, to_email: str) -> MIMEMultipart:
    body_html = f"""\
<html>
  <body style="font-family: Arial, sans-serif; font-size: 14px; color: #000000; line-height: 1.5;">
    <p>Dear {name},</p>

    <p><strong>Heading of the Mail</strong></p>

    <p><a href="{REGISTRATION_LINK}" style="color:#0000EE;">Kindly follow this link to register for the Convention</a></p>

    <p>First paragraph of the email body.</p>

    <p>The convention carries the theme <em>"Restoring Excellence; Rebuilding the Base
    of Our Alma Mater."</em> The theme reflects the association's commitment to renewing
    the school that shaped generations of its members, with a particular focus on the
    physical and academic foundations that future students will depend on.</p>

    <p><strong>A Landmark Infrastructure Drive</strong></p>

    <p>The most significant item on the programme is the launch of a One Billion Naira
    infrastructure upgrade for the college. The initiative shows the old boys' resolve
    to give back to the institution in a lasting way, and the convention will serve as
    the platform for rallying support across branches and graduating sets.</p>

    <p>Alongside the launch, delegates will witness the commissioning of Old Boys
    intervention projects. These are projects already carried out by members and
    branches for the benefit of the college.</p>

    <p><strong>Highlights of the Convention</strong></p>

    <p>The three-day gathering combines serious business with celebration:</p>
    <ul>
      <li>Launch for infrastructure upgrade in the college</li>
      <li>Commissioning of Old Boys intervention projects</li>
      <li>Novelty football match, a chance for old boys to relive their school-days rivalry</li>
      <li>Banquet dinner, an evening of fellowship, food, and reunion</li>
      <li>Special awards to deserving recipients, recognising outstanding contributions and service</li>
    </ul>

    <p><strong>A Call to All Old Boys</strong></p>

    <p>Organisers are urging every old boy, along with all branches and graduating sets,
    to take part. The association's message is that rebuilding the school is a shared
    responsibility, and the convention is the moment for every member to identify with
    STOBA's programmes and add their voice, time, and resources.</p>

    <p>Whether you graduated decades ago or only recently, the convention offers a
    chance to reconnect with classmates, meet fellow alumni, and help shape the future
    of the school.</p>

    <p><strong>Event Details</strong></p>
    <table cellpadding="4" cellspacing="0" style="border-collapse: collapse;">
      <tr><td><strong>Event</strong></td><td>Event Name</td></tr>
      <tr><td><strong>Date</strong></td><td>October 23rd to 25th, 2026</td></tr>
      <tr><td><strong>Venue</strong></td><td>The College Hall, St. Teresa's College</td></tr>
      <tr><td><strong>Theme</strong></td><td>Restoring Excellence; Rebuilding the Base of Our Alma Mater</td></tr>
    </table>

    <p><strong>Enquiries</strong><br>
    For more information, contact the organisers.</p>

    <p><em>Slogan of Association</em></p>

    <p><a href="{REGISTRATION_LINK}" style="color:#0000EE;">Register here for the Convention</a></p>

    <p>
      Kind Regards<br>
      <img src="cid:{LOGO_CID}" alt="Logo" style="height:50px; margin-top:8px;"><br>
      <strong>Sender Name</strong><br>
      St. Teresa's College<br>
      Address of the Sender <br><br>
      Email: Email of the Sender<br>
      Tel: + Phone numbers of the Sender<br>
      <a href="https://www.domainname.com">www.domainname.com</a>
    </p>
  </body>
</html>
"""

    msg = MIMEMultipart("related")
    msg["From"] = f"{SENDER_NAME} <{SENDER_EMAIL}>"
    msg["To"] = to_email
    msg["Subject"] = SUBJECT
    msg.attach(MIMEText(body_html, "html"))

    logo_part = MIMEImage(LOGO_BYTES, name="logo_name.jpg")
    logo_part.add_header("Content-ID", f"<{LOGO_CID}>")
    logo_part.add_header("Content-Disposition", "inline", filename="logo_name.jpg")
    msg.attach(logo_part)

    return msg


# ----------------------------------------------------------------------
# SENDING LOGIC
# ----------------------------------------------------------------------

def connect_smtp():
    if USE_SSL:
        context = ssl.create_default_context()
        server = smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=context, timeout=30)
    else:
        server = smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=30)
        server.starttls(context=ssl.create_default_context())
    server.login(SENDER_EMAIL, SENDER_PASSWORD)
    return server


def send_one(server, name, email_addr):
    msg = build_message(name, email_addr)
    server.sendmail(SENDER_EMAIL, email_addr, msg.as_string())


def is_valid_email(addr: str) -> bool:
    return bool(EMAIL_REGEX.match(addr))


def load_already_sent(log_path: Path) -> set:
    """Read any existing log and return the set of emails already marked SENT,
    so a restart (manual or after a crash) never re-emails someone who
    already got their message."""
    already_sent = set()
    if log_path.exists():
        with open(log_path, "r", newline="", encoding="utf-8") as f:
            for row in csv.DictReader(f):
                if row.get("STATUS") == "SENT":
                    already_sent.add(row.get("EMAIL", "").strip())
    return already_sent


def load_dead_addresses(path: Path) -> set:
    """Addresses previously confirmed dead by a bounce, across ALL past runs.
    These are skipped instantly - never retried again."""
    dead = set()
    if path.exists():
        with open(path, "r", newline="", encoding="utf-8") as f:
            for row in csv.DictReader(f):
                addr = row.get("EMAIL", "").strip()
                if addr:
                    dead.add(addr)
    return dead


def record_dead_address(path: Path, email_addr: str, reason: str):
    """Append one confirmed-dead address to the permanent list."""
    is_new = not path.exists()
    with open(path, "a", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["EMAIL", "REASON", "DATE"])
        if is_new:
            writer.writeheader()
        writer.writerow({
            "EMAIL": email_addr,
            "REASON": reason,
            "DATE": time.strftime("%Y-%m-%d %H:%M:%S"),
        })


# Matches "Final-Recipient: rfc822;someone@example.com" inside bounce reports.
_FINAL_RECIPIENT_RE = re.compile(r"Final-Recipient:\s*rfc822;\s*([^\s;]+)", re.IGNORECASE)
# "5.x.x" = permanent failure; "4.x.x" = temporary/deferred (should NOT be marked dead)
_STATUS_RE = re.compile(r"Status:\s*(\d)\.\d+\.\d+", re.IGNORECASE)


def _looks_like_bounce(msg) -> bool:
    from_header = (msg.get("From") or "").lower()
    subject = (msg.get("Subject") or "").lower()
    content_type = msg.get_content_type()
    return (
        "mailer-daemon" in from_header
        or "postmaster" in from_header
        or "delivery status" in subject
        or "undelivered" in subject
        or "returned to sender" in subject
        or content_type == "multipart/report"
    )


def check_for_bounces(dead_addresses: set) -> dict:
    """Log into the sender's own mailbox via IMAP, scan for bounce/NDR
    messages, and return {email: reason} for any PERMANENT (5.x.x) failures
    found. Marks processed messages as \\Seen so they aren't re-parsed next
    time. Newly confirmed-dead addresses are also written to DEAD_ADDRESSES_FILE
    and added to the in-memory `dead_addresses` set immediately.
    """
    newly_dead = {}

    if not ENABLE_BOUNCE_CHECKING:
        return newly_dead

    try:
        imap = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT, timeout=30)
        imap.login(SENDER_EMAIL, SENDER_PASSWORD)
        imap.select("INBOX")

        status, data = imap.search(None, "UNSEEN")
        if status != "OK":
            imap.logout()
            return newly_dead

        message_ids = data[0].split()
        for msg_id in message_ids:
            status, msg_data = imap.fetch(msg_id, "(RFC822)")
            if status != "OK" or not msg_data or msg_data[0] is None:
                continue

            raw_bytes = msg_data[0][1]
            msg = email_lib.message_from_bytes(raw_bytes)

            if not _looks_like_bounce(msg):
                continue

            full_text_parts = []
            if msg.is_multipart():
                for part in msg.walk():
                    ctype = part.get_content_type()
                    if ctype in ("text/plain", "message/delivery-status", "message/rfc822"):
                        try:
                            payload = part.get_payload(decode=True)
                            if payload:
                                full_text_parts.append(payload.decode("utf-8", errors="replace"))
                        except Exception:
                            pass
            else:
                try:
                    payload = msg.get_payload(decode=True)
                    if payload:
                        full_text_parts.append(payload.decode("utf-8", errors="replace"))
                except Exception:
                    pass

            full_text = "\n".join(full_text_parts)

            recipient_match = _FINAL_RECIPIENT_RE.search(full_text)
            status_match = _STATUS_RE.search(full_text)

            if recipient_match and status_match:
                failed_addr = recipient_match.group(1).strip()
                status_class = status_match.group(1)

                if status_class == "5" and failed_addr not in dead_addresses:
                    reason = f"Bounce: status {status_match.group(0)}"
                    newly_dead[failed_addr] = reason
                    dead_addresses.add(failed_addr)
                    record_dead_address(DEAD_ADDRESSES_FILE, failed_addr, reason)

            imap.store(msg_id, "+FLAGS", "\\Seen")

        imap.logout()

    except Exception as e:
        print(f"  (bounce check skipped - could not read mailbox: {e})")

    return newly_dead


def main():
    df = pd.read_excel(EXCEL_FILE, dtype=str)
    df.columns = df.columns.str.strip()   # guards against stray spaces in headers
    df = df.dropna(subset=[COL_EMAIL])
    df[COL_EMAIL] = df[COL_EMAIL].str.strip()

    total = len(df)
    print(f"Loaded {total} recipients from {EXCEL_FILE.name}")

    already_sent = load_already_sent(LOG_FILE)
    if already_sent:
        print(f"Found {len(already_sent)} already-SENT recipients in existing log - these will be skipped.")

    dead_addresses = load_dead_addresses(DEAD_ADDRESSES_FILE)
    if dead_addresses:
        print(f"Found {len(dead_addresses)} known-dead addresses from past bounces - these will be skipped.")

    if ENABLE_BOUNCE_CHECKING:
        print("Checking for bounce notifications from previous runs...")
        newly_dead = check_for_bounces(dead_addresses)
        if newly_dead:
            print(f"  Found {len(newly_dead)} new permanent bounce(s) - added to dead-address list.")

    print("NOTE: 'SENT' below means your mail server ACCEPTED the message, "
          "not that it was confirmed delivered. Real failures (e.g. 'no such "
          "user') often only show up later as a bounce - this script checks "
          "for those periodically and will reclassify them as BOUNCED.\n")

    server = connect_smtp()
    sent_in_batch = 0
    cooldown_cycles_used = 0
    emails_since_bounce_check = 0

    failure_times = deque()

    log_is_new = not LOG_FILE.exists()
    log_file_handle = open(LOG_FILE, "a", newline="", encoding="utf-8")
    writer = csv.DictWriter(log_file_handle, fieldnames=["NAME", "EMAIL", "STATUS", "ERROR"])
    if log_is_new:
        writer.writeheader()
        log_file_handle.flush()

    sent_count = 0
    failed_count = 0
    skipped_count = 0

    def log_row(name, email_addr, status, error):
        writer.writerow({"NAME": name, "EMAIL": email_addr, "STATUS": status, "ERROR": error})
        log_file_handle.flush()

    try:
        for idx, row in df.iterrows():
            name = str(row.get(COL_NAME, "")).strip() or "Recipient's Title"
            email_addr = str(row[COL_EMAIL]).strip()

            if email_addr in already_sent:
                continue

            if email_addr in dead_addresses:
                print(f"[{idx + 1}/{total}] {email_addr} -> SKIPPED (known dead address)")
                log_row(name, email_addr, "SKIPPED_KNOWN_DEAD", "Confirmed dead by a previous bounce")
                skipped_count += 1
                continue

            if not is_valid_email(email_addr):
                print(f"[{idx + 1}/{total}] {email_addr} -> SKIPPED (invalid format)")
                log_row(name, email_addr, "SKIPPED_INVALID_FORMAT", "Failed regex format check")
                skipped_count += 1
                continue

            success = False
            last_error = ""

            for attempt in range(1, MAX_RETRIES + 2):
                try:
                    send_one(server, name, email_addr)
                    success = True
                    break
                except (smtplib.SMTPServerDisconnected, smtplib.SMTPConnectError, OSError) as e:
                    last_error = str(e)
                    print(f"  Connection issue on {email_addr} (attempt {attempt}): {e}. Reconnecting...")
                    try:
                        server.quit()
                    except Exception:
                        pass
                    time.sleep(5)
                    server = connect_smtp()
                except smtplib.SMTPException as e:
                    last_error = str(e)
                    print(f"  SMTP error on {email_addr} (attempt {attempt}): {e}")
                    time.sleep(3)

            status = "SENT" if success else "FAILED"
            print(f"[{idx + 1}/{total}] {email_addr} -> {status}")
            log_row(name, email_addr, status, last_error)

            if success:
                sent_count += 1
            else:
                failed_count += 1
                now = time.time()
                failure_times.append(now)
                while failure_times and now - failure_times[0] > CIRCUIT_BREAKER_WINDOW_SECONDS:
                    failure_times.popleft()

                if len(failure_times) >= CIRCUIT_BREAKER_MAX_FAILURES:
                    cooldown_cycles_used += 1

                    if cooldown_cycles_used > MAX_COOLDOWN_CYCLES:
                        print("\n" + "=" * 70)
                        print(f"CIRCUIT BREAKER: hit the cooldown cap ({MAX_COOLDOWN_CYCLES} cycles).")
                        print("Stopping for good - this list likely has a deeper problem")
                        print("(e.g. many invalid/nonexistent addresses) that waiting won't fix.")
                        print("Check send_log.csv for FAILED rows before re-running.")
                        print("=" * 70)
                        break

                    print("\n" + "=" * 70)
                    print(f"CIRCUIT BREAKER TRIPPED: {len(failure_times)} failures within "
                          f"{CIRCUIT_BREAKER_WINDOW_SECONDS // 60} minutes.")
                    print(f"Pausing for {CIRCUIT_BREAKER_COOLDOWN_SECONDS // 60} minutes "
                          f"(cooldown cycle {cooldown_cycles_used}/{MAX_COOLDOWN_CYCLES}), "
                          "then resuming automatically.")
                    print("=" * 70)

                    try:
                        server.quit()
                    except Exception:
                        pass
                    time.sleep(CIRCUIT_BREAKER_COOLDOWN_SECONDS)
                    server = connect_smtp()
                    sent_in_batch = 0
                    failure_times.clear()
                    print("Resuming after cooldown...")

            sent_in_batch += 1
            emails_since_bounce_check += 1
            time.sleep(DELAY_BETWEEN_EMAILS)

            if ENABLE_BOUNCE_CHECKING and emails_since_bounce_check >= BOUNCE_CHECK_EVERY_N:
                print("  Checking for new bounce notifications...")
                newly_dead = check_for_bounces(dead_addresses)
                if newly_dead:
                    for addr in newly_dead:
                        print(f"    Confirmed dead (bounced): {addr}")
                    print(f"  {len(newly_dead)} address(es) added to the dead list - won't be retried again.")
                emails_since_bounce_check = 0

            if sent_in_batch >= BATCH_SIZE:
                print(f"--- Batch of {BATCH_SIZE} done. Reconnecting and pausing {DELAY_BETWEEN_BATCHES}s ---")
                try:
                    server.quit()
                except Exception:
                    pass
                time.sleep(DELAY_BETWEEN_BATCHES)
                server = connect_smtp()
                sent_in_batch = 0

    except KeyboardInterrupt:
        print("\nInterrupted by user (Ctrl+C). Progress up to this point is saved in the log.")

    finally:
        try:
            server.quit()
        except Exception:
            pass
        log_file_handle.close()

    if ENABLE_BOUNCE_CHECKING:
        print("\nFinal bounce check...")
        final_dead = check_for_bounces(dead_addresses)
        if final_dead:
            print(f"  {len(final_dead)} more address(es) confirmed dead - see {DEAD_ADDRESSES_FILE.name}")

    print(f"\nDone. Sent (accepted by server): {sent_count}, Failed (immediate): {failed_count}, "
          f"Skipped: {skipped_count}")
    print(f"Full log written to {LOG_FILE}")
    print(f"Confirmed-dead addresses written to {DEAD_ADDRESSES_FILE}")
    print("Note: check dead_addresses.csv and send_log.csv together for the true final picture - "
          "'Sent' above may still include some addresses that bounce later, if the bounce arrives "
          "after this script has already finished.")


if __name__ == "__main__":
    main()
