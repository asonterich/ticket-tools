"""
new_hire_setup.py — Send virtual desktop setup instructions email via Outlook.
Cross-platform: Windows uses COM automation; macOS uses AppleScript.

Usage:
  python new_hire_setup.py TICKET "Full Name"                            (auto-scrapes work email from ticket)
  python new_hire_setup.py TICKET "Full Name" user@your-company.com     (manual email override)
"""

import argparse
import os
import platform
import re
import subprocess
import sys
import tempfile

try:
    import requests
except ImportError:
    print("Missing dependency. Run:  pip install requests")
    sys.exit(1)

JIRA_BASE = "https://your-jira-instance.example.com"

# Custom field ID for the JSM client/reporter user field — update to match your Jira schema
FIELD_CLIENT_USER = "customfield_XXXXX"


def get_pat():
    pat = os.environ.get("JIRA_PAT")
    if pat:
        return pat
    if platform.system() != "Darwin":
        try:
            result = subprocess.run(
                ["powershell.exe", "-Command",
                 "[System.Environment]::GetEnvironmentVariable('JIRA_PAT', 'User')"],
                capture_output=True, text=True
            )
            pat = result.stdout.strip()
            if pat:
                return pat
        except Exception:
            pass
    print("\nERROR: JIRA_PAT not found.")
    if platform.system() == "Darwin":
        print("  Add to your shell profile:  export JIRA_PAT=\"your_token_here\"")
        print("  Then restart your terminal or run:  source ~/.zshrc")
    else:
        print("  Run:  setx JIRA_PAT \"your_token_here\"  then restart terminal")
    sys.exit(1)


def get_headers():
    return {
        "Authorization": f"Bearer {get_pat()}",
        "Content-Type": "application/json",
        "Accept": "application/json",
    }


def get_issue(ticket, headers):
    r = requests.get(f"{JIRA_BASE}/rest/api/2/issue/{ticket}", headers=headers, verify=True)
    if r.status_code == 401:
        print("ERROR: Authentication failed. Check your JIRA_PAT token.")
        sys.exit(1)
    if r.status_code == 404:
        print(f"ERROR: Ticket {ticket} not found.")
        sys.exit(1)
    r.raise_for_status()
    return r.json()


def scrape_work_email(ticket, headers):
    """Pull work email from FIELD_CLIENT_USER (JSM client field) or scrape from ticket fields."""
    import json
    issue = get_issue(ticket, headers)
    fields = issue.get("fields", {})

    client_field = (fields.get("FIELD_CLIENT_USER") or {})
    email = client_field.get("emailAddress")
    if email:
        return email

    # Fallback: scrape any work email from ticket fields — update domain to match your org
    blob = json.dumps(fields)
    matches = re.findall(r"[\w.\-+]+@your-company\.com", blob, re.I)
    if matches:
        return matches[0]

    return None


def build_html(name, work_email):
    return f"""
<p style="font-family:'Segoe UI',-apple-system,BlinkMacSystemFont,Roboto,'Helvetica Neue',sans-serif;font-size:15px;color:rgb(0,0,0);margin-top:1em;margin-bottom:1em;">Hello {name},</p>

<p style="font-family:'Segoe UI',-apple-system,BlinkMacSystemFont,Roboto,'Helvetica Neue',sans-serif;font-size:15px;color:rgb(0,0,0);margin-top:1em;margin-bottom:1em;">Here are your virtual desktop setup instructions for your upcoming orientation. Your login information will be sent to your mobile phone and personal email address. (Please check both sources of communication to receive your complete login credentials)</p>

<ol style="font-family:'Segoe UI',-apple-system,BlinkMacSystemFont,Roboto,'Helvetica Neue',sans-serif;font-size:15px;color:rgb(0,0,0);">
<li style="margin-bottom:6px;">Install and open Windows App (formerly known as Microsoft Remote Desktop) from your app store on your computer:<br><a href="https://www.microsoft.com/store/productId/9WZDNCRFJ3PS?ocid=pdpshare">https://www.microsoft.com/store/productId/9WZDNCRFJ3PS?ocid=pdpshare</a></li>
<li style="margin-bottom:6px;">Once install is complete, launch Windows App and click <b>Add</b>, then <b>Work or School Account</b>.</li>
<li style="margin-bottom:6px;">Login using your work email address and the temporary password being provided to you via your phone/personal email address.</li>
<li style="margin-bottom:6px;">You will see a <b>More Information is Required</b>&nbsp;screen. Continue to click <b>Next</b>&nbsp;and you will be prompted to install the MS Authenticator App on your mobile device.</li>
<li style="margin-bottom:6px;">Once MS Authenticator App has been installed, click the <b>+</b>&nbsp;sign at the top right corner and select <b>Scan QR Code</b>.</li>
<li style="margin-bottom:6px;">This will add the Authenticator to your device. You may proceed to Authenticate.</li>
<li style="margin-bottom:6px;">Once you see your Virtual Desktop available, click <b>Connect</b>.</li>
</ol>

<p style="font-family:'Segoe UI',-apple-system,BlinkMacSystemFont,Roboto,'Helvetica Neue',sans-serif;font-size:15px;color:rgb(0,0,0);margin-top:1em;margin-bottom:1em;">Once you have reached the Virtual Desktop, please ensure you are able to access both <b>Outlook</b>&nbsp;(email) and <b>Microsoft Teams</b>. These applications can be found in your Start Menu under <b>All Apps</b>&nbsp;or you may search for them at the bottom of your desktop.</p>

<p style="font-family:'Segoe UI',-apple-system,BlinkMacSystemFont,Roboto,'Helvetica Neue',sans-serif;font-size:15px;color:rgb(0,0,0);margin-top:1em;margin-bottom:1em;"><b>**IMPORTANT</b></p>

<p style="font-family:'Segoe UI',-apple-system,BlinkMacSystemFont,Roboto,'Helvetica Neue',sans-serif;font-size:15px;color:rgb(0,0,0);margin-top:1em;margin-bottom:1em;">The password provided to you is only a temporary password. Once you have logged in to your virtual desktop it is important to change your password to something unique. Below is the criteria and key commands to follow if you are on a MacBook or Windows laptop.</p>

<ul style="font-family:'Segoe UI',-apple-system,BlinkMacSystemFont,Roboto,'Helvetica Neue',sans-serif;font-size:15px;color:rgb(0,0,0);">
<li style="margin-bottom:6px;"><b>On a Windows Laptop:</b>&nbsp;Ctrl + Alt + End &gt; Change Password</li>
<li style="margin-bottom:6px;"><b>On a MacBook Laptop:</b>&nbsp;Control + Option + Delete &gt; Change Password</li>
</ul>

<p style="font-family:'Segoe UI',-apple-system,BlinkMacSystemFont,Roboto,'Helvetica Neue',sans-serif;font-size:15px;color:rgb(0,0,0);margin-top:1em;margin-bottom:1em;">Password criteria must meet the following:</p>

<ul style="font-family:'Segoe UI',-apple-system,BlinkMacSystemFont,Roboto,'Helvetica Neue',sans-serif;font-size:15px;color:rgb(0,0,0);">
<li>15 characters in length</li>
<li>Must contain upper &amp; lowercase letters</li>
<li>Number &amp; special characters (optional)</li>
</ul>

<p style="font-family:'Segoe UI',-apple-system,BlinkMacSystemFont,Roboto,'Helvetica Neue',sans-serif;font-size:15px;color:rgb(0,0,0);margin-top:1em;margin-bottom:1em;">*Do not use easily guessable words or number sequences such as (1234 or 2222)</p>

<p style="font-family:'Segoe UI',-apple-system,BlinkMacSystemFont,Roboto,'Helvetica Neue',sans-serif;font-size:15px;color:rgb(0,0,0);margin-top:1em;margin-bottom:1em;">If you encounter any difficulties completing the setup, we will follow up with you during your scheduled IT orientation, or you may contact the IT Help Desk during business hours for assistance.</p>

<p style="font-family:'Segoe UI',-apple-system,BlinkMacSystemFont,Roboto,'Helvetica Neue',sans-serif;font-size:15px;color:rgb(0,0,0);background-color:rgb(255,255,0);padding:8px;margin-top:1em;margin-bottom:1em;display:inline-block;"><b>Work Email: {work_email}</b><br><b>Check your phone for password.</b></p>
"""


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
    mail.HTMLBody = html_body + existing
    mail.Send()
    print(f"[OK] Email sent to {to}")


def send_email(to, subject, html_body):
    if platform.system() == "Darwin":
        send_email_mac(to, subject, html_body)
    else:
        send_email_windows(to, subject, html_body)


def main():
    parser = argparse.ArgumentParser(
        description="Send virtual desktop setup instructions via Outlook. Works on Windows and macOS.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            'Examples:\n'
            '  python new_hire_setup.py TICKET-123 "John Smith"                         (auto-scrapes email)\n'
            '  python new_hire_setup.py TICKET-123 "John Smith" jsmith@your-company.com  (manual override)'
        )
    )
    parser.add_argument("ticket", help="Ticket number, e.g. TICKET-123")
    parser.add_argument("name",   help="New hire's first name or full name")
    parser.add_argument("email",  nargs="?", default=None,
                        help="New hire's work email (optional — auto-scraped from ticket if omitted)")
    args = parser.parse_args()

    ticket = args.ticket.upper()

    if args.email:
        work_email = args.email
        print(f"[--] Using provided email: {work_email}")
    else:
        print(f"[..] Scraping ticket {ticket} for work email...")
        headers = get_headers()
        work_email = scrape_work_email(ticket, headers)
        if not work_email:
            print("ERROR: Could not find a work email on this ticket. Please provide it manually.")
            sys.exit(1)
        print(f"[OK] Found email: {work_email}")

    subject    = f"Virtual Desktop Setup Instructions - {ticket}"
    html_body  = build_html(args.name, work_email)
    send_email(work_email, subject, html_body)
    print(f"     Subject: {subject}")


if __name__ == "__main__":
    main()
