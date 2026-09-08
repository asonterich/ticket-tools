"""
apple_repair_subtask.py — Create an Apple repair invoice subtask linked to a parent IT ticket.

Usage:
  python apple_repair_subtask.py TICKET-123 --repair-id RA12345678 --serial ABC123XYZ --cost "$250.00" --description "Logic board replacement"
  python apple_repair_subtask.py TICKET-123 --repair-id RA12345678 --serial ABC123XYZ --cost "$250.00" --description "..." --notes "See attached email"
  python apple_repair_subtask.py TICKET-123 --ready     (post ready-for-pickup comment on existing subtask)

The script:
  1. Creates a subtask on the parent repair ticket
  2. Sets summary and description with all Apple repair details
  3. Assigns to the finance/payment team contact for invoice coordination
  4. Prints the new subtask URL

Setup:
  setx JIRA_PAT "your_token_here"   (then restart terminal)
"""

import argparse
import os
import sys
import json

try:
    import requests
except ImportError:
    print("Missing dependency. Run:  pip install requests")
    sys.exit(1)

JIRA_BASE        = "https://your-jira-instance.example.com"
DEFAULT_ASSIGNEE = "finance.team.username"    # Jira username of the person handling repair invoice payments
ASSIGNMENT_GROUP = "YOUR-FINANCE-TEAM-GROUP" # Assignment group for the finance/payment coordination team

# Assignment group custom field — update to match your Jira schema
FIELD_ASSIGNMENT_GROUP = "customfield_XXXXX"

SUBTASK_ISSUETYPE_NAME = "Sub-task"


def get_pat():
    pat = os.environ.get("JIRA_PAT")
    if pat:
        return pat
    try:
        import subprocess
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
    print("  Run:  setx JIRA_PAT \"your_token_here\"  then restart terminal")
    sys.exit(1)


def get_headers():
    return {
        "Authorization": f"Bearer {get_pat()}",
        "Content-Type": "application/json",
        "Accept": "application/json",
    }


def get_project_key(parent_ticket, headers):
    return parent_ticket.split("-")[0].upper()


def get_parent_summary(parent_ticket, headers):
    url = f"{JIRA_BASE}/rest/api/2/issue/{parent_ticket}"
    r = requests.get(url, headers=headers, verify=True)
    if r.status_code == 404:
        print(f"ERROR: Ticket {parent_ticket} not found.")
        sys.exit(1)
    r.raise_for_status()
    return r.json()["fields"].get("summary", "")


def get_subtask_issue_type_id(project_key, headers):
    """Get the Sub-task issue type ID for the project."""
    url = f"{JIRA_BASE}/rest/api/2/project/{project_key}"
    r = requests.get(url, headers=headers, verify=True)
    r.raise_for_status()
    issue_types = r.json().get("issueTypes", [])
    for it in issue_types:
        if it.get("subtask") or SUBTASK_ISSUETYPE_NAME.lower() in it["name"].lower():
            return it["id"]
    # Fallback: query instance-level issue types
    url2 = f"{JIRA_BASE}/rest/api/2/issuetype"
    r2 = requests.get(url2, headers=headers, verify=True)
    r2.raise_for_status()
    for it in r2.json():
        if it.get("subtask"):
            return it["id"]
    print("ERROR: Could not find Sub-task issue type. Check project configuration.")
    sys.exit(1)


def build_description(repair_id, serial, cost, description, notes):
    lines = [
        "Apple Repair — Invoice Payment",
        "",
        f"Apple Repair ID:   {repair_id}",
        f"Device Serial #:   {serial}",
        f"Estimated Cost:    {cost}",
        "",
        "Repair Description:",
        description,
    ]
    if notes:
        lines += ["", "Additional Notes:", notes]
    lines += [
        "",
        "---",
        "Action Required:",
        "When the device is ready for pickup, please arrange payment for the repair using the appropriate purchasing method.",
        "",
        "- AI-Assisted Workflow | Ticket Tools",
    ]
    return "\n".join(lines)


def create_subtask(parent_ticket, summary, description, assignee, headers):
    project_key = get_project_key(parent_ticket, headers)
    issue_type_id = get_subtask_issue_type_id(project_key, headers)

    payload = {
        "fields": {
            "project":              {"key": project_key},
            "parent":               {"key": parent_ticket},
            "summary":              summary,
            "description":          description,
            "issuetype":            {"id": issue_type_id},
            "assignee":             {"name": assignee},
            FIELD_ASSIGNMENT_GROUP: {"name": ASSIGNMENT_GROUP},
        }
    }

    url = f"{JIRA_BASE}/rest/api/2/issue"
    r = requests.post(url, headers=headers, json=payload, verify=True)

    if r.status_code == 400:
        # Retry without assignment group if it caused a field error
        del payload["fields"][FIELD_ASSIGNMENT_GROUP]
        r = requests.post(url, headers=headers, json=payload, verify=True)

    if r.status_code not in (200, 201):
        print(f"ERROR: Failed to create subtask — {r.status_code}: {r.text}")
        sys.exit(1)

    return r.json()["key"]


def post_comment(ticket, comment_text, headers):
    url = f"{JIRA_BASE}/rest/api/2/issue/{ticket}/comment"
    r = requests.post(url, headers=headers, json={"body": comment_text}, verify=True)
    r.raise_for_status()


def main():
    parser = argparse.ArgumentParser(
        description="Create an Apple repair invoice subtask assigned to the finance/payment team.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""Examples:
  python apple_repair_subtask.py TICKET-123 --repair-id RA12345678 --serial ABC123XYZ --cost "$250" --description "Logic board replacement"
  python apple_repair_subtask.py TICKET-123 --repair-id RA12345678 --serial ABC123XYZ --cost "$250" --description "..." --notes "Repair email attached"
  python apple_repair_subtask.py TICKET-456 --assignee "other.username"  (override default assignee)
        """
    )
    parser.add_argument("ticket",       help="Parent repair ticket key, e.g. TICKET-123")
    parser.add_argument("--repair-id",   required=True, metavar="ID",    help="Apple Repair Case ID (e.g. RA12345678)")
    parser.add_argument("--serial",      required=True, metavar="SN",    help="Device serial number")
    parser.add_argument("--cost",        required=True, metavar="COST",  help='Estimated repair cost, e.g. "$250.00"')
    parser.add_argument("--description", required=True, metavar="DESC",  help="Repair description from Apple")
    parser.add_argument("--notes",       default="",   metavar="NOTES",  help="Any additional notes to include")
    parser.add_argument("--assignee",    default=DEFAULT_ASSIGNEE, metavar="USER",
                        help=f"Jira username to assign subtask to (default: {DEFAULT_ASSIGNEE})")

    args = parser.parse_args()
    headers = get_headers()

    parent = args.ticket.upper()
    get_parent_summary(parent, headers)  # validates ticket exists

    subtask_summary = f"Apple Repair Invoice — {args.repair_id}"
    description = build_description(
        repair_id=args.repair_id,
        serial=args.serial,
        cost=args.cost,
        description=args.description,
        notes=args.notes,
    )

    print(f"Creating subtask on {parent}: \"{subtask_summary}\"...")
    subtask_key = create_subtask(parent, subtask_summary, description, args.assignee, headers)

    print(f"[OK] Subtask created: {subtask_key}")
    print(f"     Assigned to: {args.assignee} ({ASSIGNMENT_GROUP})")
    print(f"     Apple Repair ID: {args.repair_id}")
    print(f"\nView subtask: {JIRA_BASE}/browse/{subtask_key}")
    print(f"View parent:  {JIRA_BASE}/browse/{parent}")


if __name__ == "__main__":
    main()
