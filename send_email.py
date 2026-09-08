"""
send_email.py — Send an email via Outlook using the default signature.
Cross-platform: Windows uses COM automation; macOS uses AppleScript.

Usage:
  python send_email.py --to addr --subject "text" --body "text"
  python send_email.py --to addr --subject "text" --body "<p>HTML</p>" --html
"""

import argparse
import os
import platform
import subprocess
import sys
import tempfile


def _escape_applescript(text):
    return text.replace("\\", "\\\\").replace('"', '\\"')


def send_email_mac(to, subject, html_body):
    safe_to      = _escape_applescript(to)
    safe_subject = _escape_applescript(subject)
    safe_html    = _escape_applescript(html_body)

    script = f'''tell application "Microsoft Outlook"
    set theMessage to make new outgoing message with properties {{subject:"{safe_subject}", html content:"{safe_html}"}}
    make new to recipient at theMessage with properties {{email address:{{address:"{safe_to}"}}}}
    send theMessage
end tell'''

    with tempfile.NamedTemporaryFile(suffix=".applescript", mode="w", delete=False) as f:
        f.write(script)
        tmp = f.name
    try:
        subprocess.run(["osascript", tmp], check=True)
        print(f"[OK] Email sent to {to}")
    finally:
        os.unlink(tmp)


def send_email_windows(to, subject, html_body):
    try:
        import win32com.client
    except ImportError:
        print("Missing dependency. Run:  pip install pywin32")
        sys.exit(1)

    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = to
    mail.Subject = subject
    mail.Display(False)
    existing = mail.HTMLBody
    mail.HTMLBody = f"<div>{html_body}</div><br>{existing}"
    mail.Send()
    print(f"[OK] Email sent to {to}")


def send_email(to, subject, body, is_html=False):
    if is_html:
        html_body = body
    else:
        html_body = f"<div>{body.replace(chr(10), '<br>')}</div>"

    if platform.system() == "Darwin":
        send_email_mac(to, subject, html_body)
    else:
        send_email_windows(to, subject, html_body)


def main():
    parser = argparse.ArgumentParser(
        description="Send email via Outlook with default signature. Works on Windows and macOS."
    )
    parser.add_argument("--to",      required=True, help="Recipient email address")
    parser.add_argument("--subject", required=True, help="Email subject")
    parser.add_argument("--body",    required=True, help="Email body text or HTML")
    parser.add_argument("--html",    action="store_true", help="Treat body as raw HTML")
    args = parser.parse_args()

    send_email(args.to, args.subject, args.body, is_html=args.html)


if __name__ == "__main__":
    main()
