#!/usr/bin/env python3
"""
NetJet three step cold email sequence for Sduduzo Cele / Charisma Technology.

Builds on netjet_outreach.py (same spreadsheet, same name and email
helpers) and adds a threaded three email sequence with reply, STOP and
bounce detection over IMAP, a daily send cap, a SAST send window, and a
rerunnable state file.

Run it on your own computer. Needs Python 3.8+ and openpyxl:

    pip install openpyxl

STEPS, IN ORDER

  1. See the list's columns and how many prospects survive cleaning:

       python netjet_sequence.py list

  2. Dry run. Prints the first 3 personalised emails of each step. Sends
     nothing, asks for no password:

       python netjet_sequence.py dry-run

  3. Test. Sends Email 1, 2 and 3, threaded, to TEST_RECIPIENT only:

       python netjet_sequence.py test

  4. Live. Run once per weekday. Checks the inbox for replies, STOPs and
     bounces, shows what is due, waits for you to type CONFIRM, sends, then
     prints the daily report:

       python netjet_sequence.py run

  Report only (inbox check plus summary, nothing sent):

       python netjet_sequence.py check

  Never email someone again:

       python netjet_sequence.py dnc someone@company.co.za

Your mailbox password is asked for with getpass at the start of any run
that needs it, held in memory for that run only, and never written anywhere.
"""

import argparse
import csv
import email
import getpass
import imaplib
import os
import random
import re
import smtplib
import ssl
import sys
import time
from datetime import datetime, timedelta, timezone
from email.message import EmailMessage
from email.utils import formatdate, make_msgid, parseaddr, parsedate_to_datetime

import openpyxl

from netjet_outreach import EMAIL_RE, greeting_name, pick_email

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

FROM_NAME = "Sduduzo Cele"
FROM_EMAIL = "sduduzo@charismatech.co.za"

SMTP_HOST = "mail.charismatech.co.za"
SMTP_PORT = 465
IMAP_HOST = "mail.charismatech.co.za"
IMAP_PORT = 993

TEST_RECIPIENT = "sduduzocele43@gmail.com"

PILOT_OFFER = False

# Blocks every Email 2 (and so every Email 3) until set to False. Email 1s
# keep going out. Held prospects wait and go out on the first run after.
HOLD_STEP_2 = True

DAILY_LIMIT = 40                 # all three steps together
MIN_DELAY_SECONDS = 90
MAX_DELAY_SECONDS = 240
TEST_DELAY_SECONDS = 15          # between the three test emails only
SEND_START_HOUR = 8              # 08:00 SAST
SEND_END_HOUR = 16               # 16:00 SAST
STEP_GAP_DAYS = {2: 3, 3: 5}     # Email 2 three days after 1, Email 3 five days after 2

# Turns "WBHO Construction (Pty) Ltd" into "WBHO Construction" so the subject
# reads "Hours on WBHO Construction sites". Set False to use names as listed.
CLEAN_COMPANY_NAMES = True

# Skip the rows marked Yes in "Previously Emailed (Job Search)".
SKIP_PREVIOUSLY_EMAILED = True

INPUT_PATH = "NetJet_Prospecting_Master.xlsx"
SHEET_NAME = "Email Leads"
LOG_PATH = "sequence_log.csv"
DNC_PATH = "do_not_contact.csv"
OLD_SUPPRESSION_PATH = "netjet_suppress.txt"   # STOP list from netjet_outreach.py
OLD_SEND_LOG_PATH = "netjet_send_log.csv"      # send log from netjet_outreach.py

LOG_FIELDS = ["email", "company", "first_name", "step_sent", "last_sent_at", "message_id", "status"]
DNC_FIELDS = ["email", "reason", "added_at"]

SAST = timezone(timedelta(hours=2), "SAST")    # South Africa has no daylight saving

DASH_RE = re.compile("[\\-‐‑‒–—―−]")
STOP_RE = re.compile(r"\b(stop|unsubscribe|remove)\b", re.I)
OUR_STOP_LINE = "reply STOP and I will not contact you again"

FREE_MAIL_DOMAINS = {
    "gmail.com", "googlemail.com", "yahoo.com", "yahoo.co.za", "outlook.com",
    "hotmail.com", "hotmail.co.za", "live.com", "live.co.za", "icloud.com",
    "me.com", "aol.com", "webmail.co.za", "mweb.co.za", "telkomsa.net",
    "vodamail.co.za", "iafrica.com", "absamail.co.za", "lantic.net",
    "cybersmart.co.za", "afrihost.co.za", "xsinet.co.za",
}

# ---------------------------------------------------------------------------
# Email copy. Wording exactly as supplied; only placeholders are filled.
# ---------------------------------------------------------------------------

SUBJECT = "Hours on {company} sites"

EMAIL_1 = """Hi {first_name},

Quick question. How does {company} currently track hours worked on site?

Many contractors still rely on paper timesheets or WhatsApp photos, and they only find out a job lost money once it is finished.

NetJet fixes that. Your team captures time on their phones, supervisors approve timesheets in one place, and you can see each job's profitability while the job is still running. Quotes, purchase orders and invoices sit in the same system.

It is built and supported by Charisma Technology in Pinetown, so you deal with a local team, not an overseas help desk.

Would a 15 minute look be worth your time this week or next?

Kind regards,
Sduduzo Cele
Sales Representative, Charisma Technology
Cell: +27 71 709 8160
WhatsApp: +27 678 729 242
netjet.co.za

If this is not relevant, reply STOP and I will not contact you again."""

EMAIL_2 = """Hi {first_name},

Following up on my earlier note.

One thing I did not mention. If {company} already runs Sage, NetJet syncs with it directly, so approved hours and invoices flow through without anyone retyping them.

Pricing starts at R149 per user per month for a single module, or R349 per user for the full system.

Is site time tracking something you are looking at right now, or should I check back later in the year?

Kind regards,
Sduduzo

If this is not relevant, reply STOP and I will not contact you again."""

EMAIL_3 = """Hi {first_name},

I have not heard back, so I will assume the timing is not right and leave it here.

{pilot_paragraph}

Either way, if hours, quotes or job costs ever become a headache, my details are below.

Kind regards,
Sduduzo Cele
Cell: +27 71 709 8160
WhatsApp: +27 678 729 242

If this is not relevant, reply STOP and I will not contact you again."""

PILOT_PARAGRAPH = (
    "If it helps, we can start NetJet on one site or one team first, so you "
    "can see the numbers before committing the whole company."
)

BODIES = {1: EMAIL_1, 2: EMAIL_2, 3: EMAIL_3}


# ---------------------------------------------------------------------------
# Rendering
# ---------------------------------------------------------------------------

def no_dashes(text):
    return " ".join(DASH_RE.sub(" ", text).split())


def clean_company(raw):
    s = " ".join(str(raw).split())
    if CLEAN_COMPANY_NAMES:
        m = re.search(r"\bt/a\b\s*(.+)$", s, re.I)
        if m:
            s = m.group(1)
        s = re.sub(r"\(.*?\)", " ", s)
        s = re.sub(r"\b(pty|ltd|limited|cc|inc)\b\.?", " ", s, flags=re.I)
        s = re.sub(r"\bc\.\s?c\.?(?=\s|$)", " ", s, flags=re.I)
        s = " ".join(s.split()).strip(" ,.&/")
    return no_dashes(s)


def render(step, company, first_name):
    body = BODIES[step]
    if not first_name:
        body = body.replace("Hi {first_name},", "Good day,", 1)
    if step == 3:
        if PILOT_OFFER:
            body = body.replace("{pilot_paragraph}", PILOT_PARAGRAPH)
        else:
            body = body.replace("{pilot_paragraph}\n\n", "")
    body = body.replace("{first_name}", first_name or "").replace("{company}", company)
    subject = SUBJECT.format(company=company)
    if step > 1:
        subject = "Re: " + subject
    if DASH_RE.search(subject) or DASH_RE.search(body):
        raise ValueError(f"Dash found in rendered email for {company}; refusing to send.")
    return subject, body


def build_email(step, company, first_name, to_addr, previous_id):
    subject, body = render(step, company, first_name)
    msg = EmailMessage()
    msg["From"] = f"{FROM_NAME} <{FROM_EMAIL}>"
    msg["To"] = to_addr
    msg["Subject"] = subject
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid(domain="charismatech.co.za")
    if step > 1 and previous_id:
        msg["In-Reply-To"] = previous_id
        msg["References"] = previous_id
    msg.set_content(body)   # plain text, 7bit when ASCII
    return msg


# ---------------------------------------------------------------------------
# Prospect list
# ---------------------------------------------------------------------------

def load_prospects(path=INPUT_PATH, sheet=SHEET_NAME):
    """Returns (prospects, report). Each prospect: email, company, first_name, phone."""
    if not os.path.exists(path):
        sys.exit(f"Prospect list not found: {path}")
    wb = openpyxl.load_workbook(path, data_only=True)
    if sheet not in wb.sheetnames:
        sys.exit(f"Sheet '{sheet}' not found. Sheets: {wb.sheetnames}")
    ws = wb[sheet]
    headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    idx = {h: i for i, h in enumerate(headers) if h}
    for col in ("Company", "Email"):
        if col not in idx:
            sys.exit(f"Column '{col}' missing from '{sheet}'. Columns: {headers}")

    def cell(row, name):
        return row[idx[name]] if name in idx else None

    old_tool_sent = read_old_tool_sent()
    report = {"rows": 0, "blank": 0, "invalid": 0, "duplicate": 0, "previously_emailed": 0,
              "old_tool": 0}
    seen, prospects = set(), []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not cell(row, "Company"):
            continue
        report["rows"] += 1
        raw_email = cell(row, "Email")
        if raw_email is None or not str(raw_email).strip():
            report["blank"] += 1
            continue
        first = greeting_name(cell(row, "Contact Name"))
        addr = pick_email(raw_email, first)
        if not addr or not EMAIL_RE.fullmatch(addr):
            report["invalid"] += 1
            continue
        addr = addr.lower()
        if addr in seen:
            report["duplicate"] += 1
            continue
        seen.add(addr)
        prev = str(cell(row, "Previously Emailed (Job Search)") or "").strip().lower() == "yes"
        if prev and SKIP_PREVIOUSLY_EMAILED:
            report["previously_emailed"] += 1
            continue
        if addr in old_tool_sent:
            report["old_tool"] += 1
            continue
        prospects.append({
            "email": addr,
            "company": clean_company(cell(row, "Company")),
            "first_name": no_dashes(first) if first else "",
            "phone": str(cell(row, "Phone") or "").split(";")[0].strip(),
        })
    report["valid"] = len(prospects)
    return prospects, report, headers


# ---------------------------------------------------------------------------
# State files
# ---------------------------------------------------------------------------

def read_old_tool_sent():
    """Addresses netjet_outreach.py already emailed live. They get no Email 1."""
    if not os.path.exists(OLD_SEND_LOG_PATH):
        return set()
    with open(OLD_SEND_LOG_PATH, encoding="utf-8", newline="") as f:
        return {r["email"].strip().lower() for r in csv.DictReader(f)
                if r.get("status") == "sent" and r.get("mode") == "live"}


def read_log():
    if not os.path.exists(LOG_PATH):
        return {}
    with open(LOG_PATH, encoding="utf-8", newline="") as f:
        return {r["email"].lower(): r for r in csv.DictReader(f)}


def write_log(state):
    tmp = LOG_PATH + ".tmp"
    with open(tmp, "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=LOG_FIELDS)
        w.writeheader()
        for r in state.values():
            w.writerow({k: r.get(k, "") for k in LOG_FIELDS})
    os.replace(tmp, LOG_PATH)   # atomic, so a crash never leaves a half written log


def read_dnc():
    dnc = set()
    if os.path.exists(DNC_PATH):
        with open(DNC_PATH, encoding="utf-8", newline="") as f:
            dnc |= {r["email"].strip().lower() for r in csv.DictReader(f) if r.get("email")}
    if os.path.exists(OLD_SUPPRESSION_PATH):
        with open(OLD_SUPPRESSION_PATH, encoding="utf-8") as f:
            dnc |= {line.strip().lower() for line in f if line.strip()}
    return dnc


def add_dnc(addr, reason):
    addr = addr.strip().lower()
    if addr in read_dnc():
        return
    new_file = not os.path.exists(DNC_PATH)
    with open(DNC_PATH, "a", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=DNC_FIELDS)
        if new_file:
            w.writeheader()
        w.writerow({"email": addr, "reason": reason, "added_at": now_sast().isoformat(timespec="seconds")})


def sync_state(prospects):
    """Adds new prospects from the list to the log as step 0. Never touches existing rows."""
    state = read_log()
    dnc = read_dnc()
    added = 0
    for p in prospects:
        if p["email"] in state or p["email"] in dnc:
            continue
        state[p["email"]] = {
            "email": p["email"], "company": p["company"], "first_name": p["first_name"],
            "step_sent": "0", "last_sent_at": "", "message_id": "", "status": "active",
        }
        added += 1
    # Anyone added to DNC by hand since the last run leaves the sequence too.
    for addr, r in state.items():
        if addr in dnc and r["status"] == "active":
            r["status"] = "stopped"
    write_log(state)
    return state, added


# ---------------------------------------------------------------------------
# Time
# ---------------------------------------------------------------------------

def now_sast():
    return datetime.now(SAST)


def in_send_window(t=None):
    t = t or now_sast()
    return t.weekday() < 5 and SEND_START_HOUR <= t.hour < SEND_END_HOUR


def parse_ts(s):
    return datetime.fromisoformat(s).astimezone(SAST) if s else None


def is_due(r, today):
    if r["status"] != "active":
        return False
    step = int(r["step_sent"])
    if step == 0:
        return True
    if step >= 3 or (step == 1 and HOLD_STEP_2):
        return False
    last = parse_ts(r["last_sent_at"])
    return last is not None and today >= last.date() + timedelta(days=STEP_GAP_DAYS[step + 1])


def sent_today(state, today):
    counts = {1: 0, 2: 0, 3: 0}
    for r in state.values():
        last = parse_ts(r["last_sent_at"])
        if last and last.date() == today and int(r["step_sent"]) in counts:
            counts[int(r["step_sent"])] += 1
    return counts


# ---------------------------------------------------------------------------
# Mail connections
# ---------------------------------------------------------------------------

def smtp_connect(password):
    smtp = smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=ssl.create_default_context())
    smtp.login(FROM_EMAIL, password)
    return smtp


def imap_connect(password):
    imap = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT, ssl_context=ssl.create_default_context())
    imap.login(FROM_EMAIL, password)
    return imap


def find_sent_folder(imap):
    typ, data = imap.list()
    names = []
    for raw in data or []:
        line = raw.decode(errors="replace") if isinstance(raw, bytes) else str(raw)
        m = re.match(r'\((?P<flags>[^)]*)\)\s+(?:"[^"]*"|NIL)\s+(?P<name>.+)$', line)
        if not m:
            continue
        name = m.group("name").strip()
        if "\\sent" in m.group("flags").lower():
            return name
        names.append(name.strip('"'))
    for candidate in ("INBOX.Sent", "Sent", "Sent Items", "Sent Messages", "INBOX/Sent"):
        if candidate in names:
            return f'"{candidate}"'
    return '"INBOX.Sent"'


def append_to_sent(password, msg):
    """Copies a sent message into the IMAP Sent folder so it shows in webmail."""
    try:
        imap = imap_connect(password)
        try:
            folder = find_sent_folder(imap)
            imap.append(folder, "\\Seen", imaplib.Time2Internaldate(time.time()), msg.as_bytes())
        finally:
            imap.logout()
    except Exception as exc:
        print(f"    (could not copy to Sent folder: {exc})")


class Sender:
    """One SMTP session that reconnects once if the server drops it."""

    def __init__(self, password):
        self.password = password
        self.smtp = smtp_connect(password)

    def send(self, msg):
        try:
            self.smtp.send_message(msg)
        except smtplib.SMTPServerDisconnected:
            self.smtp = smtp_connect(self.password)
            self.smtp.send_message(msg)
        append_to_sent(self.password, msg)

    def close(self):
        try:
            self.smtp.quit()
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Inbox check: replies, STOPs, bounces
# ---------------------------------------------------------------------------

def message_text(msg):
    parts = []
    for part in msg.walk():
        ctype = part.get_content_type()
        if ctype == "text/plain":
            payload = part.get_payload(decode=True) or b""
            parts.append(payload.decode(part.get_content_charset() or "utf-8", errors="replace"))
        elif ctype in ("message/delivery-status", "text/rfc822-headers", "message/rfc822"):
            parts.append(part.as_string())
        elif ctype == "text/html" and not any(p.get_content_type() == "text/plain" for p in msg.walk()):
            payload = part.get_payload(decode=True) or b""
            html = payload.decode(part.get_content_charset() or "utf-8", errors="replace")
            parts.append(re.sub(r"<[^>]+>", " ", html))
    return "\n".join(parts)


def new_text_only(body):
    """The part of a reply the person typed, without our quoted email."""
    out = []
    for line in body.splitlines():
        s = line.strip()
        if s.startswith(">"):
            continue
        if re.match(r"^(On .{0,200}wrote:?|-+\s*Original Message\s*-+|From:\s|Sent from my |_{5,})", s, re.I):
            break
        out.append(line)
    return "\n".join(out).replace(OUR_STOP_LINE, "")


def is_auto_reply(msg):
    auto = (msg.get("Auto-Submitted") or "").lower()
    if auto and auto != "no":
        return True
    if msg.get("X-Autoreply") or msg.get("X-Autorespond"):
        return True
    if (msg.get("Precedence") or "").lower() in ("auto_reply", "bulk", "junk"):
        return True
    subj = (msg.get("Subject") or "").lower()
    return bool(re.search(r"auto(matic)?\s*reply|out of (the )?office|autoreply|on leave|away from", subj))


def is_bounce(msg):
    sender = (msg.get("From") or "").lower()
    subj = (msg.get("Subject") or "").lower()
    return bool(
        re.search(r"mailer-daemon|postmaster|mail delivery", sender)
        or re.search(r"undeliver|delivery status notification|delivery failure|failure notice|"
                     r"returned mail|mail delivery failed|could not be delivered|delivery has failed", subj)
    )


def check_inbox(password, state):
    """Updates state in place. Returns dict of lists for the report."""
    findings = {"replied": [], "stopped": [], "bounced": [], "auto_replies": []}
    watched = {a: r for a, r in state.items()
               if int(r["step_sent"]) >= 1 and r["status"] in ("active", "completed")}
    if not watched:
        return findings

    # Rough date of each prospect's first email, with a week of slack for
    # weekends and the daily cap. Mail older than that (job search replies,
    # for example) is ignored.
    first_sent_estimate = {}
    for a, r in watched.items():
        back = sum(STEP_GAP_DAYS[s] for s in range(2, int(r["step_sent"]) + 1)) + 7
        first_sent_estimate[a] = parse_ts(r["last_sent_at"]) - timedelta(days=back)
    since = min(first_sent_estimate.values())

    by_domain = {}
    for a in watched:
        d = a.split("@")[1]
        if d not in FREE_MAIL_DOMAINS and d != FROM_EMAIL.split("@")[1]:
            by_domain.setdefault(d, []).append(a)

    imap = imap_connect(password)
    try:
        imap.select("INBOX", readonly=True)    # readonly: nothing is marked as read
        typ, data = imap.search(None, "SINCE", since.strftime("%d-%b-%Y"))
        ids = data[0].split() if typ == "OK" and data and data[0] else []
        for num in ids:
            typ, hdr = imap.fetch(num, "(BODY.PEEK[HEADER])")
            if typ != "OK" or not hdr or not isinstance(hdr[0], tuple):
                continue
            head = email.message_from_bytes(hdr[0][1])
            sender = parseaddr(head.get("From", ""))[1].lower()
            if sender == FROM_EMAIL:
                continue

            bounce = is_bounce(head)
            matches = []
            if not bounce:
                if sender in watched:
                    matches = [sender]
                elif "@" in sender and sender.split("@")[1] in by_domain:
                    matches = by_domain[sender.split("@")[1]]
                if not matches:
                    continue

            typ, full = imap.fetch(num, "(BODY.PEEK[])")
            if typ != "OK" or not full or not isinstance(full[0], tuple):
                continue
            msg = email.message_from_bytes(full[0][1])
            text = message_text(msg)

            if bounce:
                lower = text.lower()
                for a, r in watched.items():
                    if a in lower and r["status"] != "bounced":
                        r["status"] = "bounced"
                        add_dnc(a, "bounced")
                        findings["bounced"].append(r)
                continue

            try:
                received = parsedate_to_datetime(msg.get("Date")).astimezone(SAST)
            except Exception:
                received = None

            for a in matches:
                r = state[a]
                if r["status"] in ("replied", "stopped", "bounced"):
                    continue
                if received and received < first_sent_estimate[a]:
                    continue
                if is_auto_reply(msg):
                    findings["auto_replies"].append((r, msg.get("Subject", "")))
                    continue
                if STOP_RE.search(new_text_only(text)):
                    r["status"] = "stopped"
                    add_dnc(a, "asked to stop")
                    findings["stopped"].append(r)
                else:
                    r["status"] = "replied"
                    add_dnc(a, "replied")
                    findings["replied"].append((r, sender))
    finally:
        try:
            imap.logout()
        except Exception:
            pass
    write_log(state)
    return findings


# ---------------------------------------------------------------------------
# Report
# ---------------------------------------------------------------------------

def print_report(state, findings, phones):
    today = now_sast().date()
    sent = sent_today(state, today)
    print("\n" + "=" * 60)
    print(f"DAILY REPORT  {today.isoformat()}")
    print("=" * 60)
    print(f"Sent today:  Email 1: {sent[1]}   Email 2: {sent[2]}   Email 3: {sent[3]}   "
          f"(total {sum(sent.values())} of {DAILY_LIMIT})")

    if findings is None:
        print("\nInbox not checked this run.")
    else:
        print(f"\nReplies detected: {len(findings['replied'])}  (call these)")
        for r, sender in findings["replied"]:
            who = f"{r['first_name']}, " if r["first_name"] else ""
            via = f" (from {sender})" if sender != r["email"] else ""
            print(f"  {r['company']}  {who}{r['email']}{via}  {phones.get(r['email'], '')}")
        print(f"\nStops: {len(findings['stopped'])}")
        for r in findings["stopped"]:
            print(f"  {r['company']}  {r['email']}")
        print(f"\nBounces: {len(findings['bounced'])}")
        for r in findings["bounced"]:
            print(f"  {r['company']}  {r['email']}")
        if findings["auto_replies"]:
            print(f"\nAuto replies ignored (still in sequence): {len(findings['auto_replies'])}")
            for r, subj in findings["auto_replies"]:
                print(f"  {r['company']}  {r['email']}  \"{subj}\"")

    active = [r for r in state.values() if r["status"] == "active"]
    print("\nProspects left in each step:")
    print(f"  Waiting for Email 1: {sum(1 for r in active if r['step_sent'] == '0')}")
    waiting_2 = sum(1 for r in active if r['step_sent'] == '1')
    print(f"  Waiting for Email 2: {waiting_2}" + ("  (ON HOLD: HOLD_STEP_2 = True)" if HOLD_STEP_2 else ""))
    print(f"  Waiting for Email 3: {sum(1 for r in active if r['step_sent'] == '2')}")
    for status in ("completed", "replied", "stopped", "bounced"):
        print(f"  {status.capitalize()}: {sum(1 for r in state.values() if r['status'] == status)}")


# ---------------------------------------------------------------------------
# Commands
# ---------------------------------------------------------------------------

def cmd_list(args):
    prospects, rep, headers = load_prospects(args.input)
    print(f"File: {args.input}, sheet '{SHEET_NAME}'")
    print("Columns:", ", ".join(str(h) for h in headers if h))
    print("\nMapping: company <- Company, first_name <- Contact Name (first name only), email <- Email")
    print(f"\nRows with a company: {rep['rows']}")
    print(f"  Blank email:          {rep['blank']}")
    print(f"  Invalid email:        {rep['invalid']}")
    print(f"  Duplicate email:      {rep['duplicate']}")
    if SKIP_PREVIOUSLY_EMAILED:
        print(f"  Previously emailed (job search), skipped: {rep['previously_emailed']}")
    if os.path.exists(OLD_SEND_LOG_PATH):
        print(f"  Already emailed by netjet_outreach.py ({OLD_SEND_LOG_PATH}): {rep['old_tool']}")
    dnc = read_dnc()
    on_dnc = sum(1 for p in prospects if p["email"] in dnc)
    if on_dnc:
        print(f"  On do_not_contact list: {on_dnc}")
    print(f"Valid prospects: {rep['valid'] - on_dnc}")
    print(f"  with a first name: {sum(1 for p in prospects if p['first_name'])}, "
          f"\"Good day\" greeting: {sum(1 for p in prospects if not p['first_name'])}")


def cmd_dry_run(args):
    prospects, rep, _ = load_prospects(args.input)
    sample = prospects[: args.count]
    print(f"DRY RUN. Nothing is sent. {rep['valid']} valid prospects; showing {len(sample)} per step.")
    print(f"PILOT_OFFER = {PILOT_OFFER}   HOLD_STEP_2 = {HOLD_STEP_2}\n")
    for step in (1, 2, 3):
        for p in sample:
            subject, body = render(step, p["company"], p["first_name"])
            print("=" * 72)
            print(f"EMAIL {step}  To: {p['email']}")
            print(f"Subject: {subject}")
            if step > 1:
                print("In-Reply-To / References: <Message-ID of Email {}>".format(step - 1))
            print("-" * 72)
            print(body)
            print()


def cmd_test(args):
    prospects, _, _ = load_prospects(args.input)
    if args.sample:
        matches = [p for p in prospects if p["email"] == args.sample.lower()]
        if not matches:
            sys.exit(f"{args.sample} is not in the cleaned list.")
        p = matches[0]
    else:
        p = prospects[0]
    if HOLD_STEP_2:
        print("Note: HOLD_STEP_2 only blocks prospects. The test still sends Email 2 to you.")
    print(f"TEST: sending Email 1, 2 and 3 to {TEST_RECIPIENT} only, personalised as "
          f"{p['company']} / {p['first_name'] or 'Good day'}. No prospect is emailed.")
    password = getpass.getpass(f"Password for {FROM_EMAIL}: ")
    sender = Sender(password)
    previous = None
    try:
        for step in (1, 2, 3):
            msg = build_email(step, p["company"], p["first_name"], TEST_RECIPIENT, previous)
            sender.send(msg)
            previous = msg["Message-ID"]
            print(f"  Sent Email {step}: {msg['Subject']}  {previous}")
            if step < 3:
                time.sleep(TEST_DELAY_SECONDS)
    finally:
        sender.close()
    print(f"\nDone. Check {TEST_RECIPIENT}: all three should sit in one thread. "
          "The log was not touched.")


def phone_map(prospects):
    return {p["email"]: p["phone"] for p in prospects}


def cmd_check(args):
    prospects, _, _ = load_prospects(args.input)
    state, added = sync_state(prospects)
    password = getpass.getpass(f"Password for {FROM_EMAIL}: ")
    findings = check_inbox(password, state)
    print_report(state, findings, phone_map(prospects))


def cmd_run(args):
    prospects, _, _ = load_prospects(args.input)
    state, added = sync_state(prospects)
    if added:
        print(f"{added} new prospect(s) added to {LOG_PATH}.")

    password = getpass.getpass(f"Password for {FROM_EMAIL}: ")
    print("Checking inbox for replies, STOPs and bounces...")
    findings = check_inbox(password, state)

    now = now_sast()
    today = now.date()
    already = sum(sent_today(state, today).values())
    room = max(0, DAILY_LIMIT - already)

    due = [r for r in state.values() if is_due(r, today)]
    # Follow ups first so threads stay on schedule, then new Email 1s in list order.
    followups = sorted((r for r in due if r["step_sent"] != "0"), key=lambda r: r["last_sent_at"])
    fresh = [r for r in due if r["step_sent"] == "0"]
    queue = (followups + fresh)[:room]
    if HOLD_STEP_2:
        held = sum(1 for r in state.values() if r["status"] == "active" and r["step_sent"] == "1"
                   and parse_ts(r["last_sent_at"]).date() + timedelta(days=STEP_GAP_DAYS[2]) <= today)
        print(f"\nEmail 2 is ON HOLD (HOLD_STEP_2 = True). {held} prospect(s) would otherwise be due.")

    if not in_send_window(now):
        print(f"\nOutside the send window (Mon to Fri, {SEND_START_HOUR:02d}:00 to "
              f"{SEND_END_HOUR:02d}:00 SAST). Now {now:%a %H:%M} SAST. Nothing sent.")
        print_report(state, findings, phone_map(prospects))
        return
    if not queue:
        reason = "daily limit reached" if room == 0 and due else "nothing due"
        print(f"\nNo emails to send today ({reason}).")
        print_report(state, findings, phone_map(prospects))
        return

    counts = {s: sum(1 for r in queue if int(r["step_sent"]) + 1 == s) for s in (1, 2, 3)}
    print(f"\nDue now: {len(queue)} email(s)  (Email 1: {counts[1]}, Email 2: {counts[2]}, "
          f"Email 3: {counts[3]}). Already sent today: {already}.")
    print(f"Delay between emails: {MIN_DELAY_SECONDS} to {MAX_DELAY_SECONDS} seconds.")
    if input("Type CONFIRM to email real prospects: ").strip() != "CONFIRM":
        print("Not confirmed. Nothing sent.")
        print_report(state, findings, phone_map(prospects))
        return

    sender = Sender(password)
    dnc = read_dnc()
    try:
        for i, r in enumerate(queue):
            if not in_send_window():
                print("Send window closed. Stopping; the rest go out on the next run.")
                break
            if r["email"] in dnc or r["status"] != "active":
                continue
            step = int(r["step_sent"]) + 1
            msg = build_email(step, r["company"], r["first_name"], r["email"], r["message_id"] or None)
            try:
                sender.send(msg)
            except Exception as exc:
                print(f"[{i + 1}/{len(queue)}] FAILED Email {step} to {r['company']} <{r['email']}>: {exc}")
                continue
            r["step_sent"] = str(step)
            r["last_sent_at"] = now_sast().isoformat(timespec="seconds")
            r["message_id"] = msg["Message-ID"]
            if step == 3:
                r["status"] = "completed"
            write_log(state)    # saved after every send, so a rerun never double sends
            print(f"[{i + 1}/{len(queue)}] Sent Email {step} to {r['company']} <{r['email']}>")
            if i < len(queue) - 1:
                time.sleep(random.uniform(MIN_DELAY_SECONDS, MAX_DELAY_SECONDS))
    except KeyboardInterrupt:
        print("\nStopped by you. Everything sent so far is saved.")
    finally:
        sender.close()

    print_report(state, findings, phone_map(prospects))


def cmd_dnc(args):
    state = read_log()
    for addr in args.emails:
        add_dnc(addr, "added by hand")
        r = state.get(addr.strip().lower())
        if r and r["status"] == "active":
            r["status"] = "stopped"
    if state:
        write_log(state)
    print(f"Added {len(args.emails)} address(es) to {DNC_PATH}.")


def main():
    parser = argparse.ArgumentParser(description="NetJet three step email sequence")
    parser.add_argument("--input", default=INPUT_PATH)
    sub = parser.add_subparsers(dest="command", required=True)

    sub.add_parser("list", help="Show columns and the cleaned prospect count.").set_defaults(func=cmd_list)

    p = sub.add_parser("dry-run", help="Print sample emails for each step. Sends nothing.")
    p.add_argument("--count", type=int, default=3)
    p.set_defaults(func=cmd_dry_run)

    p = sub.add_parser("test", help=f"Send all three emails, threaded, to {TEST_RECIPIENT} only.")
    p.add_argument("--sample", help="Prospect email to personalise the test with (default: first in list).")
    p.set_defaults(func=cmd_test)

    sub.add_parser("check", help="Inbox check and report. Sends nothing.").set_defaults(func=cmd_check)
    sub.add_parser("run", help="Daily live run. Asks for CONFIRM before sending.").set_defaults(func=cmd_run)

    p = sub.add_parser("dnc", help="Add addresses to the do_not_contact list.")
    p.add_argument("emails", nargs="+")
    p.set_defaults(func=cmd_dnc)

    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
