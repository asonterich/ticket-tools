"""
equipment_return.py — Post equipment return comment to Jira and open Outlook draft.
Cross-platform: Windows uses COM/VBScript; macOS uses AppleScript.

Usage:
  python equipment_return.py TICKET "Full Name"                          (auto-scrapes email, defaults to office_a)
  python equipment_return.py TICKET "Full Name" --site office_b          (site-specific onsite template)
  python equipment_return.py TICKET "Full Name" --remote                 (remote template)
  python equipment_return.py TICKET "Full Name" --email-only             (skip Jira comment)
  python equipment_return.py TICKET "Full Name" email@example.com        (override email manually)

Sites: office_a, office_b, office_c, office_d
"""

import argparse
import os
import platform
import re
import sys
import subprocess
import json
import tempfile

try:
    import requests
except ImportError:
    print("Missing dependency. Run:  pip install requests")
    sys.exit(1)

JIRA_BASE  = "https://your-jira-instance.example.com"
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# Custom field ID for the JSM client/reporter user field — update to match your Jira schema
FIELD_CLIENT_USER = "customfield_XXXXX"

JIRA_COMMENT_ONSITE = (
    "Contacted client via email to coordinate return of company equipment (onsite). "
    "Requested follow-up to arrange drop-off at IT."
)

JIRA_COMMENT_REMOTE = (
    "Contacted client via email to coordinate return of company equipment (remote). "
    "Requested follow-up to arrange FedEx shipment of applicable devices."
)

EMAIL_SUBJECT = "Equipment Return - {ticket}"

# Update these with your office locations and IT staff contact info
SITE_DROPOFF = {
    "office_a": "any IT staff member at your main office location",
    "office_b": "any IT staff member at your second office location",
    "office_c": "any IT staff member at your third office location",
    "office_d": "any IT staff member at your fourth office location",
}

EMAIL_BODY_ONSITE = """\
Hello {name},

I hope this message finds you well. As part of your separation from the organization, Information Technology requires that all company-issued equipment be returned to our inventory. Returning your equipment ensures that you will not be held responsible for any devices following your departure.

Items that may require return include:

- Laptop
- Mobile device(s)
- 2FA authenticator (security key or token)

When you are ready, equipment can be dropped off with {dropoff}.

Please don't hesitate to reach out if you have any questions. We wish you all the best in your future endeavors.

Thank you,
"""

EMAIL_BODY_REMOTE = """\
Hello {name},

I hope this message finds you well. As part of your separation from the organization, Information Technology requires that all company-issued equipment be returned to our inventory. Returning your equipment ensures that you will not be held responsible for any devices following your departure.

Items that may require return include:

- Laptop
- Mobile device(s)
- 2FA authenticator (security key or token)
- External monitor (if applicable)
- Docking station (if applicable)

To facilitate the return, we can arrange FedEx shipping at no cost to you. Please let us know which items you will be returning so we can send appropriately sized packaging materials.

Please don't hesitate to reach out if you have any questions. We wish you all the best in your future endeavors.

Thank you,
"""

# Update ORG_DOMAINS to match your organization's email domain(s)
ORG_DOMAINS = {"your-company.com"}
EMAIL_RE = re.compile(r"[\w.\-+]+@[\w.\-]+\.[a-z]{2,}", re.I)


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


def scrape_personal_email(ticket, headers):
    """Walk the ticket then its parent looking for a non-org email address."""
    issue = get_issue(ticket, headers)
    fields = issue.get("fields", {})

    parent_key = (fields.get("parent") or {}).get("key")
    candidates = [fields]
    if parent_key:
        parent_issue = get_issue(parent_key, headers)
        candidates.append(parent_issue.get("fields", {}))

    for f in candidates:
        blob = json.dumps(f)
        for email in EMAIL_RE.findall(blob):
            domain = email.split("@", 1)[1].lower()
            if not any(domain.endswith(d) for d in ORG_DOMAINS):
                return email

    return None


def post_comment(ticket, comment_text, headers):
    url = f"{JIRA_BASE}/rest/api/2/issue/{ticket}/comment"
    r = requests.post(url, headers=headers, json={"body": comment_text}, verify=True)
    if r.status_code == 401:
        print("ERROR: Authentication failed. Check your JIRA_PAT token.")
        sys.exit(1)
    if r.status_code == 404:
        print(f"ERROR: Ticket {ticket} not found.")
        sys.exit(1)
    r.raise_for_status()


def body_to_html(text):
    lines = text.strip().splitlines()
    html_lines = ["<div style=\"font-family:Calibri,sans-serif;font-size:11pt;\">"]
    in_list = False
    for line in lines:
        line = line.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        if line.startswith("- "):
            if not in_list:
                html_lines.append("<ul style='margin:4px 0;padding-left:20px;'>")
                in_list = True
            html_lines.append(f"<li>{line[2:]}</li>")
        else:
            if in_list:
                html_lines.append("</ul>")
                in_list = False
            if line == "":
                html_lines.append("<br>")
            else:
                html_lines.append(f"<p style='margin:0'>{line}</p>")
    if in_list:
        html_lines.append("</ul>")
    html_lines.append("</div><br>")
    return "".join(html_lines)


def _escape_applescript(text):
    return text.replace("\\", "\\\\").replace('"', '\\"')


def open_outlook_draft_mac(to_email, subject, html_body):
    safe_to      = _escape_applescript(to_email)
    safe_subject = _escape_applescript(subject)
    safe_html    = _escape_applescript(html_body)

    script = f'''tell application "Microsoft Outlook"
    activate
    set theMessage to make new outgoing message with properties {{subject:"{safe_subject}", html content:"{safe_html}"}}
    make new to recipient at theMessage with properties {{email address:{{address:"{safe_to}"}}}}
    open theMessage
end tell'''

    with tempfile.NamedTemporaryFile(suffix=".applescript", mode="w", delete=False) as f:
        f.write(script)
        tmp = f.name
    try:
        subprocess.Popen(["osascript", tmp])
    except Exception as e:
        print(f"WARNING: Could not open Outlook draft via AppleScript: {e}")
    finally:
        # Small delay before cleanup so osascript has time to read the file
        import time; time.sleep(1)
        os.unlink(tmp)


def open_outlook_draft_windows(to_email, subject, body):
    html_body = body_to_html(body).replace('"', '""')
    safe_subject = subject.replace('"', '""')

    vbs = f"""Set oApp = CreateObject("Outlook.Application")
Set oMail = oApp.CreateItem(0)
oMail.To = "{to_email}"
oMail.Subject = "{safe_subject}"
oMail.Display
oMail.HTMLBody = "{html_body}" & oMail.HTMLBody
"""
    vbs_path = os.path.join(SCRIPT_DIR, "_draft_temp.vbs")
    with open(vbs_path, "w") as f:
        f.write(vbs)
    subprocess.Popen(["wscript.exe", vbs_path])


def open_outlook_draft(to_email, subject, body):
    if platform.system() == "Darwin":
        open_outlook_draft_mac(to_email, subject, body_to_html(body))
    else:
        open_outlook_draft_windows(to_email, subject, body)


def main():
    parser = argparse.ArgumentParser(
        description="Post Jira comment and open Outlook draft for equipment return. Works on Windows and macOS.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            'Examples:\n'
            '  python equipment_return.py TICKET-123 "John Smith"\n'
            '  python equipment_return.py TICKET-123 "John Smith" --site office_b\n'
            '  python equipment_return.py TICKET-123 "John Smith" --remote\n'
            '  python equipment_return.py TICKET-123 "John Smith" john@gmail.com  (manual email override)\n'
            'Sites: office_a, office_b, office_c, office_d'
        )
    )
    parser.add_argument("ticket", help="Ticket key, e.g. TICKET-123")
    parser.add_argument("name",   nargs="?", default=None,
                        help="Client's full name (auto-scraped from ticket summary if omitted)")
    parser.add_argument("email",  nargs="?", default=None,
                        help="Client's personal email (optional — auto-scraped from parent ticket if omitted)")
    parser.add_argument("--site", default="office_a",
                        choices=["office_a", "office_b", "office_c", "office_d"],
                        help="Your office site — determines drop-off location in email (default: office_a)")
    parser.add_argument("--remote", action="store_true",
                        help="Use remote template (includes monitor/dock, FedEx instructions)")
    parser.add_argument("--email-only", action="store_true",
                        help="Open Outlook draft without posting a Jira comment")
    args = parser.parse_args()

    ticket = args.ticket.upper()
    headers = get_headers()

    issue = get_issue(ticket, headers)

    if args.name:
        client_name = args.name
    else:
        client_field_name = ((issue.get("fields") or {}).get("FIELD_CLIENT_USER") or {}).get("displayName")
        if client_field_name:
            client_name = client_field_name
            print(f"[..] Name from ticket client field: {client_name}")
        else:
            summary = (issue.get("fields") or {}).get("summary", "")
            m = re.search(r"(?i)\bfrom\s+([A-Z][a-z]+(?:\s+[A-Z][a-z]+)+)", summary)
            if m:
                client_name = m.group(1)
            else:
                summary = re.sub(r"(?i)(equipment\s+return\s*[-–:]\s*|[-–:]\s*equipment\s+return)", "", summary).strip()
                client_name = summary if summary else "there"
            print(f"[..] Name scraped from ticket: {client_name}")

    client_field = (issue.get("fields") or {}).get("FIELD_CLIENT_USER") or {}
    work_email = client_field.get("emailAddress") or f"{(issue.get('fields') or {}).get('reporter', {}).get('name', '')}@your-company.com"

    if args.email:
        personal_email = args.email
        print(f"[--] Using provided personal email: {personal_email}")
    else:
        print(f"[..] Scraping parent ticket for personal email...")
        personal_email = scrape_personal_email(ticket, headers)
        if not personal_email:
            print("WARNING: No personal email found. Drafting to work email only.")
        else:
            print(f"[OK] Found personal email: {personal_email}")

    to_email = f"{work_email}; {personal_email}" if personal_email else work_email
    subject  = EMAIL_SUBJECT.format(ticket=ticket)

    if args.remote:
        jira_comment = JIRA_COMMENT_REMOTE
        email_body   = EMAIL_BODY_REMOTE.format(name=client_name)
        mode_label   = "remote"
    else:
        dropoff      = SITE_DROPOFF.get(args.site, SITE_DROPOFF["office_a"])
        jira_comment = JIRA_COMMENT_ONSITE
        email_body   = EMAIL_BODY_ONSITE.format(name=client_name, dropoff=dropoff)
        mode_label   = f"onsite ({args.site})"

    if not args.email_only:
        post_comment(ticket, jira_comment, headers)
        print(f"[OK] Comment posted on {ticket} ({mode_label})")
        print(f"     View: {JIRA_BASE}/browse/{ticket}")

    open_outlook_draft(to_email, subject, email_body)
    print(f"[OK] Outlook draft opened -> {to_email} ({mode_label} template)")


if __name__ == "__main__":
    main()
