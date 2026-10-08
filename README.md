# Mass-Email-To-Many-Recipients
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
