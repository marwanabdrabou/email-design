#!/usr/bin/env python3
"""
send_email.py - send the Site Bot launch email (email.html) through Outlook.

All local images used in email.html (logo, phone, laptop, QR code) are EMBEDDED
inside the message (CID attachments), so you do NOT need to host them online and
they show up in Outlook without the "download pictures" prompt.

Three ways to send
------------------
1) Outlook desktop app (Windows)  -> uses the account already signed in to Outlook
       pip install pywin32
       python send_email.py --method outlook --to client@company.com

2) SMTP (Office 365 / any OS)     -> needs your mailbox login
       set SMTP_USER=you@company.com
       set SMTP_PASSWORD=your-password-or-app-password
       python send_email.py --method smtp --to client@company.com

3) Gmail                          -> sends from your Gmail address
       (one-time) Google Account > Security > turn on 2-Step Verification,
       then create an "App password" (16 characters) at myaccount.google.com/apppasswords
       set SMTP_USER=you@gmail.com
       set SMTP_PASSWORD=abcdefghijklmnop        (the App password, NOT your normal password)
       python send_email.py --method gmail --to client@company.com --send
       (Mac/Linux: use  export SMTP_USER=...  instead of  set)

Safe by default: without --send nothing is delivered.
  - outlook: opens the draft on screen so you can check it
  - smtp / gmail : writes preview.eml (double-click to open in Outlook)

Examples
--------
  python send_email.py --method outlook --to a@x.com b@y.com            # draft only
  python send_email.py --method outlook --to-file clients.txt --send    # send one mail per client
  python send_email.py --method smtp --to a@x.com --send
  python send_email.py --method gmail --to-file clients.txt --send      # from Gmail
"""
import argparse
import mimetypes
import os
import re
import smtplib
import sys
import time
from email.message import EmailMessage
from email.utils import formatdate, make_msgid
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_HTML = BASE_DIR / "email.html"
DEFAULT_SUBJECT = "Meet Site Bot - Now Live"

PLAIN_TEXT = (
    "Meet Site Bot - now live!\n\n"
    "Stop searching. Start asking.\n"
    "Site Bot finds any site, location or network configuration just by asking.\n\n"
    "This email is best viewed in an HTML-capable mail client.\n\n"
    "(c) 2026 Telecom Egypt, Inc"
)

# ------------------------------------------------------------------ helpers
def load_recipients(args):
    recipients = ["marwan.gamal@te.eg"]
    if args.to_file:
        for line in Path(args.to_file).read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line and not line.startswith("#"):
                recipients.append(line)
    # remove duplicates, keep order
    seen, unique = set(), []
    for r in recipients:
        if r.lower() not in seen:
            seen.add(r.lower())
            unique.append(r)
    return unique


def prepare_html(html_path: Path):
    """Replace every local image src with a cid: reference.
    Returns (new_html, [(file_path, cid), ...])."""
    html = html_path.read_text(encoding="utf-8")
    images, cid_by_path = [], {}

    def repl(match):
        src = match.group(1)
        if re.match(r"^(https?:|cid:|data:|mailto:|//)", src, re.I):
            return match.group(0)  # leave web images alone
        file_path = (html_path.parent / src).resolve()
        if not file_path.exists():
            sys.exit(f"Image not found: {file_path}  (referenced in {html_path.name})")
        if file_path not in cid_by_path:
            cid = f"img{len(cid_by_path) + 1}_{file_path.stem}".replace(" ", "_")
            cid_by_path[file_path] = cid
            images.append((file_path, cid))
        return f'src="cid:{cid_by_path[file_path]}"'

    new_html = re.sub(r'src="([^"]+)"', repl, html)
    return new_html, images


# ------------------------------------------------------------------ method 1: Outlook desktop
# def send_with_outlook(recipients, subject, html, images, send, delay):
#     try:
#         import win32com.client  # pip install pywin32
#     except ImportError:
#         sys.exit("pywin32 is missing (Windows only). Run:  pip install pywin32")

#     outlook = win32com.client.Dispatch("Outlook.Application")
#     PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
#     PR_ATTACHMENT_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"

#     targets = recipients if send else recipients[:1] or [""]
#     for i, to in enumerate(targets, 1):
#         mail = outlook.CreateItem(0)  # 0 = olMailItem
#         mail.To = to
#         mail.Subject = subject
#         mail.HTMLBody = html
#         for path, cid in images:
#             att = mail.Attachments.Add(str(path))
#             pa = att.PropertyAccessor
#             pa.SetProperty(PR_ATTACH_CONTENT_ID, cid)
#             pa.SetProperty(PR_ATTACHMENT_HIDDEN, True)  # keep out of the paperclip list
#         if send:
#             mail.Send()
#             print(f"[{i}/{len(targets)}] sent -> {to}")
#             time.sleep(delay)
#         else:
#             mail.Display()
#             print("Draft opened in Outlook. Check it, then re-run with --send to deliver.")


# ------------------------------------------------------------------ method 2: SMTP (Office 365)
def build_message(sender, to, subject, html, images):
    msg = EmailMessage()
    msg["From"] = sender
    msg["To"] = to
    msg["Subject"] = subject
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid()
    msg.set_content(PLAIN_TEXT)
    msg.add_alternative(html, subtype="html")
    html_part = msg.get_payload()[1]
    for path, cid in images:
        ctype = mimetypes.guess_type(path.name)[0] or "image/png"
        maintype, subtype = ctype.split("/")
        html_part.add_related(
            path.read_bytes(), maintype, subtype, cid=f"<{cid}>", filename=path.name,
            disposition="inline",
        )
    return msg


def send_with_smtp(recipients, subject, html, images, send, delay, args):
    user = os.environ.get("SMTP_USER", "marwangamal507@gmail.com")
    password = os.environ.get("SMTP_PASSWORD", "lxyd wlpr ensy xrsy")
    sender = args.sender or user or "sender@example.com"

    if not send:
        msg = build_message(sender, ", ".join(recipients[:1]) or "client@example.com",
                            subject, html, images)
        out = BASE_DIR / "preview.eml"
        out.write_bytes(bytes(msg))
        print(f"Preview written to {out}  (double-click it to open in Outlook).")
        print("Nothing was sent. Re-run with --send to deliver.")
        return

    if not user or not password:
        sys.exit("Set SMTP_USER and SMTP_PASSWORD environment variables first.")
    with smtplib.SMTP(args.smtp_host, args.smtp_port, timeout=60) as server:
        server.starttls()
        server.login(user, password)
        for i, to in enumerate(recipients, 1):
            msg = build_message(sender, to, subject, html, images)
            server.send_message(msg)
            print(f"[{i}/{len(recipients)}] sent -> {to}")
            time.sleep(delay)


# ------------------------------------------------------------------ main
def main():
    p = argparse.ArgumentParser(description="Send the Site Bot launch email via Outlook.")
    p.add_argument("--method", choices=["outlook", "smtp", "gmail"], default="outlook")
    p.add_argument("--to", nargs="*", help="one or more recipient emails")
    p.add_argument("--to-file", help="text file with one email per line")
    p.add_argument("--subject", default=DEFAULT_SUBJECT)
    p.add_argument("--html", default=str(DEFAULT_HTML), help="path to the HTML file")
    p.add_argument("--sender", help="From address (smtp only, defaults to SMTP_USER)")
    p.add_argument("--smtp-host", default="smtp.office365.com")
    p.add_argument("--smtp-port", type=int, default=587)
    p.add_argument("--delay", type=float, default=2.0, help="seconds between emails")
    p.add_argument("--send", action="store_true", help="really send (otherwise preview only)")
    args = p.parse_args()

    recipients = load_recipients(args)
    if args.send and not recipients:
        sys.exit("No recipients. Use --to or --to-file.")

    html, images = prepare_html(Path(args.html))
    print(f"Embedded {len(images)} image(s): " + ", ".join(p.name for p, _ in images))

    
    if args.method == "gmail":
        args.smtp_host, args.smtp_port = "smtp.gmail.com", 587
    send_with_smtp(recipients, args.subject, html, images, args.send, args.delay, args)


if __name__ == "__main__":
    main()