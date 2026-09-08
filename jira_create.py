"""
jira_create.py — Create a new Jira ticket from a template or inline flags.

Usage:
  python jira_create.py --template software_install --var software=Zoom --var user="John Smith"
  python jira_create.py --summary "Laptop issue" --project IT --issuetype Incident --assignment-group YOUR-TEAM-GROUP
  python jira_create.py --search-user johnson
  python jira_create.py --list-templates

Options:
  --template          Template name from create_templates.json
  --var key=value     Fill template placeholder (repeatable)
  --summary           Ticket title (overrides template)
  --description       Ticket description (overrides template)
  --project           Jira project key (default: IT)
  --issuetype         Issue type name (default: Incident)
  --assignment-group  Assignment group name (required unless in template)
  --reporter          Jira username of reporter (defaults to current user)
  --assignee          Jira username to assign the ticket to (optional)
  --search-user       Search for a user by partial name
  --list-templates    Print all available templates and exit

Setup:
  setx JIRA_PAT "your_token_here"   (then restart terminal)
"""

import argparse
import os
import sys
import json
import subprocess

try:
    import requests
except ImportError:
    print("Missing dependency. Run:  pip install requests")
    sys.exit(1)

JIRA_BASE = "https://your-jira-instance.example.com"
DEFAULT_TEMPLATES_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "create_templates.json")
AI_TAG = "\n\n---\n_Created via Ticket Tools automation._"

DEFAULT_PROJECT   = "IT"
DEFAULT_ISSUETYPE = "Incident"

# Assignment group custom field — update to match your Jira schema.
# Find yours at: your-jira-instance/rest/api/2/field
FIELD_ASSIGNMENT_GROUP = "customfield_XXXXX"


def get_pat():
    pat = os.environ.get("JIRA_PAT")
    if pat:
        return pat
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
    print("  1. Go to your Jira instance > Profile > Configure API")
    print("  2. Create a Bearer token")
    print("  3. Run:  setx JIRA_PAT \"your_token_here\"  then restart terminal")
    sys.exit(1)


def get_headers():
    return {
        "Authorization": f"Bearer {get_pat()}",
        "Content-Type": "application/json",
        "Accept": "application/json",
    }


def fill_placeholders(value, vars_dict):
    if not isinstance(value, str):
        return value
    for k, v in vars_dict.items():
        value = value.replace(f"{{{k}}}", v)
    return value


def apply_vars(obj, vars_dict):
    if isinstance(obj, str):
        return fill_placeholders(obj, vars_dict)
    if isinstance(obj, dict):
        return {k: apply_vars(v, vars_dict) for k, v in obj.items()}
    return obj


def load_template(name, templates_file=None):
    path = templates_file or DEFAULT_TEMPLATES_FILE
    if not os.path.exists(path):
        print(f"ERROR: Templates file not found at {path}")
        sys.exit(1)
    with open(path) as f:
        templates = json.load(f)
    if name not in templates:
        available = ", ".join(templates.keys())
        print(f"ERROR: Template '{name}' not found.\nAvailable: {available}")
        sys.exit(1)
    return templates[name]


def get_current_user(headers):
    r = requests.get(f"{JIRA_BASE}/rest/api/2/myself", headers=headers, verify=True)
    r.raise_for_status()
    return r.json()["name"]


def get_required_fields(project, issuetype, headers):
    """Query a recent ticket's editmeta to discover required fields and their first allowed value."""
    r = requests.get(
        f"{JIRA_BASE}/rest/api/2/search",
        params={"jql": f"project={project} ORDER BY created DESC", "maxResults": 1, "fields": "summary"},
        headers=headers, verify=True
    )
    issues = r.json().get("issues", [])
    if not issues:
        return {}
    sample_key = issues[0]["key"]

    r = requests.get(f"{JIRA_BASE}/rest/api/2/issue/{sample_key}?expand=editmeta", headers=headers, verify=True)
    fields = r.json().get("editmeta", {}).get("fields", {})

    # Skip fields handled explicitly so they don't get double-set
    skip = {"summary", "reporter", "assignee", FIELD_ASSIGNMENT_GROUP}
    required = {}
    for key, meta in fields.items():
        if not meta.get("required") or key in skip:
            continue
        allowed = meta.get("allowedValues", [])
        if not allowed:
            continue
        first = allowed[0]
        entry = {"id": str(first["id"])}
        children = first.get("children", [])
        if children:
            entry["child"] = {"id": str(children[0]["id"])}
        required[key] = entry
    return required


def search_user(query, headers):
    r = requests.get(
        f"{JIRA_BASE}/rest/api/2/user/search",
        params={"username": query, "maxResults": 10},
        headers=headers, verify=True
    )
    r.raise_for_status()
    return r.json()


def create_ticket(args, vars_dict, headers):
    template = {}
    if args.template:
        raw = load_template(args.template, getattr(args, "templates_file", None))
        template = apply_vars(raw, vars_dict)

    project          = args.project          or template.get("project",          DEFAULT_PROJECT)
    issuetype        = args.issuetype        or template.get("issuetype",        DEFAULT_ISSUETYPE)
    summary          = args.summary          or template.get("summary")
    description      = args.description      or template.get("description")      or summary
    assignment_group = args.assignment_group or template.get("assignment_group")
    assignee         = args.assignee         or template.get("assignee")

    if not summary:
        print("ERROR: --summary is required (or use a template that includes one).")
        sys.exit(1)
    if not assignment_group:
        print("ERROR: --assignment-group is required (or use a template that includes one).")
        sys.exit(1)

    reporter = args.reporter or get_current_user(headers)

    fields = get_required_fields(project, issuetype, headers)

    fields.update({
        "project":              {"key": project},
        "issuetype":            {"name": issuetype},
        "summary":              summary,
        "description":          (description or summary) + AI_TAG,
        "reporter":             {"name": reporter},
        FIELD_ASSIGNMENT_GROUP: {"name": assignment_group},
    })

    if assignee:
        fields["assignee"] = {"name": assignee}

    # Template customfields (e.g. resolution category) applied last as overrides
    for k, v in template.items():
        if k.startswith("customfield_"):
            fields[k] = v

    r = requests.post(
        f"{JIRA_BASE}/rest/api/2/issue",
        headers=headers, json={"fields": fields}, verify=True
    )
    if r.status_code == 400:
        print("ERROR: Jira rejected the ticket.")
        print(json.dumps(r.json(), indent=2))
        sys.exit(1)
    r.raise_for_status()
    return r.json(), reporter, project, issuetype, summary, assignment_group


def main():
    parser = argparse.ArgumentParser(
        description="Create a new Jira ticket.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""Examples:
  python jira_create.py --template software_install --var software=Zoom --var user="John Smith"
  python jira_create.py --summary "Laptop issue" --project IT --issuetype Incident --assignment-group YOUR-TEAM-GROUP
  python jira_create.py --search-user johnson
        """
    )
    parser.add_argument("--template",         help="Template name from create_templates.json")
    parser.add_argument("--var",              metavar="key=value", action="append", default=[],
                        help="Fill a template placeholder (repeatable)")
    parser.add_argument("--summary",          help="Ticket title (overrides template)")
    parser.add_argument("--description",      help="Ticket description (overrides template)")
    parser.add_argument("--project",          help=f"Jira project key (default: {DEFAULT_PROJECT})")
    parser.add_argument("--issuetype",        help=f"Issue type (default: {DEFAULT_ISSUETYPE})")
    parser.add_argument("--assignment-group", dest="assignment_group",
                        help="Assignment group name")
    parser.add_argument("--reporter",         help="Jira username of reporter (defaults to current user)")
    parser.add_argument("--assignee",         help="Jira username to assign the ticket to")
    parser.add_argument("--search-user",      metavar="QUERY",
                        help="Search for a Jira user by partial name")
    parser.add_argument("--templates-file",   metavar="PATH",
                        help="Path to a templates JSON file (default: create_templates.json)")
    parser.add_argument("--list-templates",   action="store_true",
                        help="Print all templates in the active templates file and exit.")
    args = parser.parse_args()

    headers = get_headers()

    if args.list_templates:
        path = args.templates_file or DEFAULT_TEMPLATES_FILE
        if not os.path.exists(path):
            print(f"ERROR: Templates file not found at {path}")
            sys.exit(1)
        with open(path) as f:
            templates = json.load(f)
        print(f"\nTemplates in {os.path.basename(path)}:")
        for name, t in templates.items():
            print(f"  {name:<30} project: {t.get('project','?')} | {t.get('summary','')[:60]}")
        return

    if args.search_user:
        users = search_user(args.search_user, headers)
        if not users:
            print(f"No users found matching '{args.search_user}'")
        else:
            print(f"\nUsers matching '{args.search_user}':")
            for u in users:
                status = "active" if u.get("active") else "inactive"
                print(f"  {u['name']:<20} {u['displayName']:<30} ({status})")
        return

    if not args.template and not args.summary:
        parser.print_help()
        sys.exit(0)

    vars_dict = {}
    for v in args.var:
        if "=" not in v:
            print(f"ERROR: --var must be in key=value format, got: {v}")
            sys.exit(1)
        k, val = v.split("=", 1)
        vars_dict[k.strip()] = val.strip()

    result, reporter, project, issuetype, summary, assignment_group = create_ticket(
        args, vars_dict, headers
    )

    key = result["key"]
    print(f"[OK] Ticket created: {key}")
    print(f"     Summary:          {summary}")
    print(f"     Project:          {project} / {issuetype}")
    print(f"     Reporter:         {reporter}")
    print(f"     Assignment Group: {assignment_group}")
    if args.assignee:
        print(f"     Assignee:         {args.assignee}")
    print(f"\nView ticket: {JIRA_BASE}/browse/{key}")


if __name__ == "__main__":
    main()
