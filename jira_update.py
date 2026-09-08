"""
jira_update.py — Transition a Jira ticket and post a comment.

Usage:
  python jira_update.py TICKET-123 --template mobile_inprogress
  python jira_update.py TICKET-123 --template mobile_resolved --var asset_tag=AT-00123
  python jira_update.py TICKET-123 --template mobile_scheduled --var date="June 25"
  python jira_update.py TICKET-123 --status "in progress" --comment "custom note"
  python jira_update.py TICKET-123 --list-transitions
  python jira_update.py --list-templates

Setup:
  setx JIRA_PAT "your_token_here"   (then restart terminal)
"""

import argparse
import os
import sys
import json
from datetime import datetime

try:
    import requests
except ImportError:
    print("Missing dependency. Run:  pip install requests")
    sys.exit(1)

JIRA_BASE     = "https://your-jira-instance.example.com"
TEMPLATE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "jira_templates.json")

# Custom field IDs — these are instance-specific. Update to match your Jira schema.
# Find yours at: your-jira-instance/rest/api/2/field
FIELD_ASSIGNMENT_GROUP  = "customfield_XXXXX"  # Assignment group field
FIELD_RESOLUTION_CAT    = "customfield_XXXXX"  # Resolution categorization (cascading select)
FIELD_PENDING_REASON    = "customfield_XXXXX"  # Pending reason field
FIELD_CLIENT_USER       = "customfield_XXXXX"  # JSM client/reporter user field


def get_pat():
    # Try current process environment first
    pat = os.environ.get("JIRA_PAT")
    if pat:
        return pat
    # Fall back to Windows user environment (survives sessions where setx was used)
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
    print("  1. Go to your Jira instance > Profile > Configure API")
    print("  2. Create a Bearer token")
    print("  3. Run:  setx JIRA_PAT \"your_token_here\"  then restart terminal")
    sys.exit(1)


def get_headers():
    pat = get_pat()
    return {
        "Authorization": f"Bearer {pat}",
        "Content-Type": "application/json",
        "Accept": "application/json",
    }


def load_templates():
    if not os.path.exists(TEMPLATE_FILE):
        print(f"ERROR: Template file not found at {TEMPLATE_FILE}")
        sys.exit(1)
    with open(TEMPLATE_FILE) as f:
        return json.load(f)


def get_transitions(ticket, headers):
    url = f"{JIRA_BASE}/rest/api/2/issue/{ticket}/transitions"
    r = requests.get(url, headers=headers, verify=True)
    if r.status_code == 401:
        print("ERROR: Authentication failed. Check your JIRA_PAT token.")
        sys.exit(1)
    if r.status_code == 404:
        print(f"ERROR: Ticket {ticket} not found.")
        sys.exit(1)
    r.raise_for_status()
    return r.json().get("transitions", [])


def find_transition(transitions, status_input):
    needle = status_input.lower().strip()
    matches = [t for t in transitions if needle in t["name"].lower()]
    if len(matches) == 1:
        return matches[0]["id"], matches[0]["name"]
    if len(matches) > 1:
        names = ", ".join(t["name"] for t in matches)
        print(f"ERROR: '{status_input}' matches multiple transitions: {names}")
        print("Be more specific.")
        sys.exit(1)
    return None, None


def get_assignment_group(ticket, headers):
    url = f"{JIRA_BASE}/rest/api/2/issue/{ticket}"
    r = requests.get(url, headers=headers, verify=True)
    r.raise_for_status()
    ag = r.json().get("fields", {}).get(FIELD_ASSIGNMENT_GROUP)
    return ag.get("name") if ag else None


def transition_ticket(ticket, transition_id, headers, resolution=None,
                      resolution_category=None, resolution_subcategory=None, comment=None,
                      pending_reason_id=None):
    url = f"{JIRA_BASE}/rest/api/2/issue/{ticket}/transitions"
    ag = get_assignment_group(ticket, headers)

    def build_payload(include_ag=True, include_comment=True, include_pending=True):
        fields = {}
        if resolution:
            fields["resolution"] = {"name": resolution}
        if include_ag and ag:
            fields[FIELD_ASSIGNMENT_GROUP] = {"name": ag}
        if resolution_category:
            fields[FIELD_RESOLUTION_CAT] = {"value": resolution_category}
            if resolution_subcategory:
                fields[FIELD_RESOLUTION_CAT]["child"] = {"value": resolution_subcategory}
        if include_pending and pending_reason_id:
            fields[FIELD_PENDING_REASON] = {"id": pending_reason_id}
        p = {"transition": {"id": transition_id}}
        if fields:
            p["fields"] = fields
        if include_comment and comment:
            p["update"] = {"comment": [{"add": {"body": comment}}]}
        return p

    # Attempt 1: full payload
    r = requests.post(url, headers=headers, json=build_payload(), verify=True)
    if r.status_code == 400:
        # Attempt 2: drop assignment group
        r = requests.post(url, headers=headers, json=build_payload(include_ag=False), verify=True)
    if r.status_code == 400 and comment:
        # Attempt 3: transition without comment, post comment separately
        r = requests.post(url, headers=headers, json=build_payload(include_ag=False, include_comment=False), verify=True)
        r.raise_for_status()
        return False
    r.raise_for_status()
    return True


def post_comment(ticket, comment_text, headers, internal=False):
    url = f"{JIRA_BASE}/rest/api/2/issue/{ticket}/comment"
    payload = {"body": comment_text}
    if internal:
        payload["properties"] = [{"key": "sd.public.comment", "value": {"internal": True}}]
    r = requests.post(url, headers=headers, json=payload, verify=True)
    r.raise_for_status()
    return r.json().get("id")


def get_current_user(headers):
    r = requests.get(f"{JIRA_BASE}/rest/api/2/myself", headers=headers, verify=True)
    r.raise_for_status()
    data = r.json()
    return data.get("name", "").lower(), data.get("displayName", "").lower()


def edit_last_comment(ticket, new_text, headers):
    username, displayname = get_current_user(headers)
    url = f"{JIRA_BASE}/rest/api/2/issue/{ticket}/comment"
    r = requests.get(url, headers=headers, verify=True)
    r.raise_for_status()
    comments = r.json().get("comments", [])
    my_comments = [c for c in comments
                   if username in c["author"].get("name", "").lower()
                   or displayname in c["author"].get("displayName", "").lower()]
    if not my_comments:
        my_comments = comments  # fall back to last comment overall
    last = my_comments[-1]
    edit_url = f"{JIRA_BASE}/rest/api/2/issue/{ticket}/comment/{last['id']}"
    r = requests.put(edit_url, headers=headers, json={"body": new_text}, verify=True)
    r.raise_for_status()
    return last["id"]


def apply_vars(text, var_list):
    """Replace {placeholders} in text with --var key=value pairs."""
    substitutions = {}
    for item in (var_list or []):
        if "=" not in item:
            print(f"ERROR: --var must be in key=value format, got: {item}")
            sys.exit(1)
        k, v = item.split("=", 1)
        substitutions[k.strip()] = v.strip()
    for k, v in substitutions.items():
        text = text.replace("{" + k + "}", v)
    # Warn about any unfilled placeholders
    import re
    remaining = re.findall(r"\{(\w+)\}", text)
    if remaining:
        print(f"WARNING: unfilled placeholders: {remaining}")
        print(f"  Use --var key=value to fill them in.")
    return text


def main():
    parser = argparse.ArgumentParser(
        description="Update a Jira ticket status and post a comment.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""Examples:
  python jira_update.py TICKET-123 --template mobile_inprogress
  python jira_update.py TICKET-123 --template mobile_resolved --var asset_tag=AT-00123
  python jira_update.py TICKET-123 --template mobile_scheduled --var date="June 25"
  python jira_update.py TICKET-123 --status "in progress" --comment "custom note"
  python jira_update.py TICKET-123 --list-transitions
  python jira_update.py --list-templates
        """
    )
    parser.add_argument("ticket", nargs="?", help="Ticket key, e.g. TICKET-123")
    parser.add_argument("--template", help="Template name from jira_templates.json")
    parser.add_argument("--var", action="append", metavar="key=value",
                        help="Fill in a placeholder in the template, e.g. --var asset_tag=AT-001")
    parser.add_argument("--status", help="Part of the transition name (case-insensitive).")
    parser.add_argument("--comment", help="Custom comment text to post on the ticket.")
    parser.add_argument("--edit-comment", help="Replace the last comment on the ticket with this text.")
    parser.add_argument("--internal", action="store_true",
                        help="Post comment as internal note (not visible to client).")
    parser.add_argument("--pending-reason", metavar="REASON",
                        help="Pending reason label (e.g. 'Client action required'). Defaults to 'Client action required' when transitioning to Pending.")
    parser.add_argument("--list-transitions", action="store_true",
                        help="Print all available transitions for this ticket and exit.")
    parser.add_argument("--list-templates", action="store_true",
                        help="Print all available templates and exit.")
    args = parser.parse_args()

    # List templates — no ticket needed
    if args.list_templates:
        templates = load_templates()
        print("\nAvailable templates:")
        for name, t in templates.items():
            status_label = t.get("status") or "(comment only)"
            print(f"  {name:<25} status: {status_label}")
            print(f"    {t['comment'][:80]}{'...' if len(t['comment']) > 80 else ''}")
        return

    if not args.ticket:
        parser.print_help()
        sys.exit(0)

    headers = get_headers()
    ticket = args.ticket.upper()
    transitions = get_transitions(ticket, headers)

    if args.list_transitions:
        print(f"\nAvailable transitions for {ticket}:")
        for t in transitions:
            print(f"  {t['name']}")
        print(f"\nTip: use any part of the name with --status")
        return

    status_to_use  = args.status
    comment_to_use = args.comment

    # Pending reason option IDs — these are instance-specific.
    # Find yours at: your-jira-instance/rest/api/2/field then query the field's allowedValues.
    # Replace each "YOUR_OPTION_ID" with the numeric ID from your Jira instance.
    PENDING_REASON_IDS = {
        "awaiting approval":           "YOUR_OPTION_ID",
        "client action required":      "YOUR_OPTION_ID",
        "client hold":                 "YOUR_OPTION_ID",
        "pending on change request":   "YOUR_OPTION_ID",
        "pending problem resolution":  "YOUR_OPTION_ID",
        "support contact hold":        "YOUR_OPTION_ID",
        "third party action required": "YOUR_OPTION_ID",
    }

    # Load from template if specified
    resolution_to_use = None
    resolution_cat    = None
    resolution_subcat = None

    if args.template:
        templates = load_templates()
        if args.template not in templates:
            print(f"ERROR: Template '{args.template}' not found.")
            print("Run:  python jira_update.py --list-templates")
            sys.exit(1)
        tmpl = templates[args.template]
        if tmpl.get("status") and not status_to_use:
            status_to_use = tmpl["status"]
        if tmpl.get("comment") and not comment_to_use:
            comment_to_use = apply_vars(tmpl["comment"], args.var)
        resolution_to_use = tmpl.get("resolution")
        resolution_cat    = tmpl.get("resolution_category")
        resolution_subcat = tmpl.get("resolution_subcategory")

    if args.edit_comment:
        edit_last_comment(ticket, args.edit_comment, headers)
        print(f"[OK] Last comment updated on {ticket}")
        print(f"\nView ticket: {JIRA_BASE}/browse/{ticket}")
        return

    if not status_to_use and not comment_to_use:
        parser.print_help()
        sys.exit(0)

    if status_to_use:
        tid, tname = find_transition(transitions, status_to_use)
        if not tid:
            print(f"\nERROR: No transition matching '{status_to_use}' found for {ticket}.")
            print(f"Run:  python jira_update.py {ticket} --list-transitions")
            sys.exit(1)

        # Resolve pending reason — use flag, or default to "client action required" for Pending
        pending_reason_id = None
        if "pending" in tname.lower():
            reason_label = (args.pending_reason or "client action required").lower().strip()
            pending_reason_id = PENDING_REASON_IDS.get(reason_label)
            if args.pending_reason and not pending_reason_id:
                print(f"WARNING: Unknown pending reason '{args.pending_reason}'. Valid options: {list(PENDING_REASON_IDS.keys())}")

        # Pass comment into transition so workflows that require it get satisfied
        comment_posted = transition_ticket(ticket, tid, headers,
                          resolution=resolution_to_use,
                          resolution_category=resolution_cat,
                          resolution_subcategory=resolution_subcat,
                          comment=comment_to_use,
                          pending_reason_id=pending_reason_id)
        print(f"[OK] {ticket} -> {tname}")
        if comment_to_use and comment_posted is False:
            post_comment(ticket, comment_to_use, headers, internal=args.internal)
            print(f"[OK] Comment posted separately")
        elif comment_to_use:
            print(f"[OK] Comment posted with transition")
    elif comment_to_use:
        post_comment(ticket, comment_to_use, headers, internal=args.internal)
        print(f"[OK] Comment posted")

    print(f"\nView ticket: {JIRA_BASE}/browse/{ticket}")


if __name__ == "__main__":
    main()
