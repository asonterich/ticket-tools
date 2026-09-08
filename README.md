# Ticket Tools

A CLI toolkit for reducing time spent on repetitive administrative tasks — so you can stay focused on the actual work. Built around Jira and Outlook, it collapses multi-step manual workflows (ticket transitions, templated comments, email drafting, subtask creation) into single parameterized commands.

---

## The Problem

Administrative workflows eat into the time that should be spent on actual work. Whether it's updating a ticket status, drafting a client email, or coordinating a multi-party task like a vendor repair — each action alone is manageable, but repeated dozens of times a week across a team, it becomes a real drag on productivity. Open the system, find the record, navigate through menus, copy a template, fill in the details, switch to email, repeat.

This toolkit was built to eliminate that overhead entirely — collapsing each workflow into a single command that handles the Jira update, the comment, and the email draft in one step.

## Impact

Originally built for a single technician's daily workflow, Ticket Tools grew into a shared toolkit used across multiple teams handling IT support, facilities requests, and equipment logistics. The template system made it easy for team members to adopt without writing any code — they run a command, the ticket updates, the email drafts, and they move on.

Across the teams using it, the tool reduced the time spent on ticket documentation and client communication from several minutes per action to seconds. Workflows that previously required switching between three or four systems — Jira, Outlook, asset records, and email history — were reduced to a single CLI command with the relevant details passed as flags. The consistency of templated comments also improved the quality of ticket documentation across the board, since every technician's notes followed the same structure regardless of who worked the ticket.

---

## Architecture

```
jira_templates.json      ──►  jira_update.py    (update existing tickets)
create_templates.json    ──►  jira_create.py    (create new tickets)
email_templates.json     ──►  equipment_return.py / new_hire_setup.py
                              apple_repair_subtask.py  (multi-step Jira subtask)
                              send_email.py            (ad-hoc Outlook email)
```

All scripts authenticate via a Jira Personal Access Token stored as an environment variable. No credentials are ever hardcoded. Auth is handled silently — the token is read from the current process environment first, then from the Windows user environment on Windows (for terminals opened before `setx` was run) or the shell profile on macOS, with a clear error message if neither is found.

The Jira scripts (`jira_update.py`, `jira_create.py`, `apple_repair_subtask.py`) are fully cross-platform. The email scripts (`send_email.py`, `equipment_return.py`, `new_hire_setup.py`) detect the OS at runtime and use Outlook COM automation on Windows or AppleScript on macOS — no separate files, no flags required.

---

## What It Does

| Script | What it automates |
|---|---|
| `jira_update.py` | Transition any Jira ticket and post a comment using named templates or inline flags |
| `jira_create.py` | Create a new Jira ticket from a named template or inline flags, auto-resolving required fields |
| `apple_repair_subtask.py` | Create an Apple repair invoice subtask, assign to the finance team, and print parent + subtask links in one command |
| `equipment_return.py` | Post a Jira comment AND open a pre-filled Outlook draft for equipment return emails in one command — auto-scrapes the client's personal email from the parent ticket |
| `new_hire_setup.py` | Send HTML-formatted virtual desktop setup instructions to a new hire, auto-scraping their work email from the JSM client field |
| `send_email.py` | Send an ad-hoc email through Outlook COM with your default signature preserved |

Template files:

| File | Contents |
|---|---|
| `jira_templates.json` | 20+ named comment templates for common ticket types (mobile setup, loaner devices, workstation setup, reimaging, security keys, equipment return, Apple repair) |
| `create_templates.json` | Named creation templates for IT and facilities tickets with pre-filled project key, issue type, and assignment group |
| `email_templates.json` | Named email templates for client-facing workflows (Windows 11 upgrade coordination, mobile device pickup) |

---

## Technical Highlights

**Graceful transition retry logic** — `transition_ticket()` in `jira_update.py` builds three progressively simplified payloads and retries on `400` errors. This handles JSM instances where required fields vary by workflow state (e.g., some transitions require assignment group preservation; others reject it).

**Live required-field discovery** — `jira_create.py` queries the Jira editmeta API on a recent ticket before creation, so required fields get a valid default value even when not explicitly known. This avoids hard-coded field mappings that break when instances change.

**Dual-channel email scraping** — `equipment_return.py` walks both the current ticket and its parent ticket's full JSON payload looking for a non-org email address. Clients often include personal email in off-boarding tickets; this finds it automatically.

**Placeholder substitution with warnings** — All templates support `{placeholder}` syntax. `apply_vars()` fills them from `--var key=value` flags and warns (rather than silently failing) if any unfilled placeholders remain in the output.

**Internal comment support** — `--internal` flag in `jira_update.py` sets the `sd.public.comment` property so the note posts as an internal-only Jira Service Management note, not visible to the end user.

**Edit last comment** — `--edit-comment` locates your most recent comment on a ticket and replaces it in-place using the author identity returned from `/rest/api/2/myself`.

---

## Setup

**1. Install dependencies**

```
pip install -r requirements.txt
```

**2. Configure your Jira URL**

Open each script and update `JIRA_BASE` at the top:

```python
JIRA_BASE = "https://your-jira-instance.example.com"
```

**3. Set your Jira Personal Access Token**

Generate a token at your Jira instance under Profile > Configure API, then set it as a persistent Windows environment variable:

```
setx JIRA_PAT "your_token_here"
```

Restart your terminal after running `setx`. The scripts read the token automatically — no manual input needed.

**4. Update org-specific placeholders**

A few values need to match your environment:

- `JIRA_BASE` in every script — your Jira instance URL
- `ORG_DOMAINS` in `equipment_return.py` — your company's email domain(s) (used to distinguish work vs. personal email)
- `SITE_DROPOFF` in `equipment_return.py` — your office locations and IT contact info
- `FIELD_*` constants in each script — Jira custom field IDs (`customfield_XXXXX`); find yours at `/rest/api/2/field`
- `PENDING_REASON_IDS` in `jira_update.py` — option IDs for your Pending Reason field; query the field's `allowedValues` endpoint

---

## Usage

### jira_update.py

Transition a ticket and post a comment using a named template:

```
python jira_update.py TICKET-123 --template mobile_inprogress
python jira_update.py TICKET-123 --template loaner_laptop_pickup --var asset_tag=LT-00456
python jira_update.py TICKET-123 --template refresh_resolved --var device_type=MacBook --var asset_tag=AT-001 --var serial=ABC123 --var hostname=CORP-001 --var pickup_status="picked up" --var collected_old=yes --var previous_asset_tag=AT-000
```

Transition without a template:

```
python jira_update.py TICKET-123 --status "in progress" --comment "Looking into this now."
python jira_update.py TICKET-123 --status resolve --comment "Issue resolved."
```

Other flags:

```
python jira_update.py TICKET-123 --comment "Internal note." --internal
python jira_update.py TICKET-123 --edit-comment "Corrected comment text."
python jira_update.py TICKET-123 --list-transitions
python jira_update.py --list-templates
```

### jira_create.py

Create a new ticket from a template:

```
python jira_create.py --template software_install --var software=Zoom --var user="John Smith"
python jira_create.py --template light_replacement --var room=B204 --var site=Main
```

Create a ticket with inline flags:

```
python jira_create.py --summary "Laptop issue" --project IT --issuetype Incident --assignment-group YOUR-TEAM-GROUP
```

Other flags:

```
python jira_create.py --list-templates
python jira_create.py --search-user johnson
```

### apple_repair_subtask.py

Creates a subtask on a parent repair ticket and assigns it to the finance/payment team for invoice coordination:

```
python apple_repair_subtask.py TICKET-123 --repair-id RA12345678 --serial ABC123XYZ --cost "$250.00" --description "Logic board replacement"
python apple_repair_subtask.py TICKET-123 --repair-id RA12345678 --serial ABC123XYZ --cost "$250.00" --description "..." --notes "Repair email attached"
python apple_repair_subtask.py TICKET-123 --repair-id RA12345678 --serial ABC123XYZ --cost "$250.00" --description "..." --assignee "other.user"
```

### equipment_return.py

Posts a standard Jira comment and opens a pre-filled Outlook draft in one command:

```
python equipment_return.py TICKET-123 "John Smith"
python equipment_return.py TICKET-123 "John Smith" --site office_b
python equipment_return.py TICKET-123 "John Smith" --remote
python equipment_return.py TICKET-123 "John Smith" john@gmail.com
python equipment_return.py TICKET-123 "John Smith" --email-only
```

The script automatically scrapes the client's personal email from the parent ticket if not provided. The `--remote` flag switches to a template that includes monitor, dock, and FedEx shipping instructions.

### new_hire_setup.py

Sends virtual desktop setup instructions to a new hire:

```
python new_hire_setup.py TICKET-123 "John Smith"
python new_hire_setup.py TICKET-123 "John Smith" jsmith@your-company.com
```

Auto-scrapes the work email from the ticket's JSM client field. Falls back to manual input if not found.

### send_email.py

Send an ad-hoc email with your Outlook default signature:

```
python send_email.py --to user@example.com --subject "Your Subject" --body "Message body."
python send_email.py --to user@example.com --subject "Subject" --body "<p>HTML body</p>" --html
```

---

## Template Reference

### jira_templates.json

Templates that drive `jira_update.py`. Each entry defines a `status` (transition name), `comment`, and optional resolution fields.

| Template | Transition | Variables |
|---|---|---|
| `mobile_inprogress` | In Progress | — |
| `mobile_scheduled` | — (comment only) | `{date}` |
| `mobile_resolved` | Resolve | `{device_type}`, `{asset_tag}` |
| `loaner_phone_inprogress` | In Progress | `{office}` |
| `loaner_phone_pickup` | Pending | `{asset_tag}` |
| `loaner_phone_resolved` | Resolve | `{asset_tag}` |
| `loaner_laptop_inprogress` | In Progress | `{office}` |
| `loaner_laptop_pickup` | Pending | `{asset_tag}` |
| `loaner_laptop_resolved` | Resolve | `{asset_tag}` |
| `workstation_setup_inprogress` | In Progress | — |
| `workstation_setup_resolved` | Resolve | — |
| `equipment_return` | — (comment only) | — |
| `yubikey_inprogress` | In Progress | `{office}` |
| `yubikey_shipped` | In Progress | — |
| `yubikey_resolved` | Resolve | — |
| `token_inprogress` | In Progress | `{office}` |
| `token_shipped` | In Progress | — |
| `token_resolved` | Resolve | — |
| `win11_resolved` | Resolve | — |
| `refresh_resolved` | Resolve | `{device_type}`, `{pickup_status}`, `{hostname}`, `{asset_tag}`, `{serial}`, `{collected_old}`, `{previous_asset_tag}` |
| `reimage_resolved` | Resolve | — |
| `apple_repair_dropoff` | In Progress | `{repair_id}` |
| `apple_repair_ready` | — (comment only) | `{repair_id}` |

Fill variables using `--var key=value`. Multiple `--var` flags are supported.

### create_templates.json

Templates that drive `jira_create.py`. Each entry defines a `project`, `issuetype`, `assignment_group`, `summary`, and `description`. All values support `{placeholder}` substitution via `--var`.

| Template | Project | Summary |
|---|---|---|
| `software_install` | IT | Please install {software} on {user} computer |
| `access_request` | IT | Access request - {access_type} for {user} |
| `tech_remove` | IT | WO - Remove technology/hardware {room} |
| `tech_needed` | IT | WO - Technology Needed {room} |
| `monitor_replace` | IT | WO - Replace monitors {room} |
| `move_relocation` | IT | WO - {item} move {room} |
| `conference_tech_reconnect` | IT | WO - Reconnecting conference table tech {room} |
| `light_replacement` | FACILITIES | WO - Overhead light replacement {room} |
| `restroom_plumbing` | FACILITIES | WO - {description} Restroom {location} |
| `furniture_remove` | FACILITIES | WO - Furniture remove {room} |
| `furniture_deliver` | FACILITIES | WO - Furniture delivery {room} |
| `furniture_repair` | FACILITIES | WO - Furniture repair {room} |
| `office_cleaning` | FACILITIES | WO - Office Cleaning {room} |
| `carpet_cleaning` | FACILITIES | WO - Carpet Cleaning {room} |
| `hvac_issue` | FACILITIES | WO - HVAC issue {room} |

### email_templates.json

Templates for client-facing emails. Loaded by custom scripts or referenced directly.

| Template | Subject | Variables |
|---|---|---|
| `win11_secure_boot` | Windows 11 Upgrade – Action Required | `{first_name}`, `{device_model}`, `{ticket}`, `{office}` |
| `win11_schedule` | Windows 11 Upgrade – Scheduling | `{first_name}`, `{device_model}`, `{ticket}` |
| `mobile_pickup` | Mobile Device Ready for Pickup | `{first_name}`, `{ticket}`, `{office}` |

---

## Requirements

- Python 3.8+
- Jira Server or Data Center with Personal Access Token support
- Microsoft Outlook installed (Windows or macOS)

```
pip install requests pywin32   # Windows
pip install requests           # macOS (pywin32 not needed)
```

**macOS:** Scripts use `osascript` (AppleScript) to drive Outlook — no additional dependencies required. Set your token in your shell profile instead of `setx`:

```bash
export JIRA_PAT="your_token_here"   # add to ~/.zshrc or ~/.bash_profile
```

**Windows:** Scripts use `win32com.client` (pywin32) for Outlook COM automation and fall back to the Windows user environment for the PAT if not set in the current session.

---

## Notes on Custom Fields

This toolkit was built against a Jira Service Management instance with specific custom fields. The field ID constants (`FIELD_ASSIGNMENT_GROUP`, `FIELD_RESOLUTION_CAT`, `FIELD_PENDING_REASON`, `FIELD_CLIENT_USER`) at the top of each script are placeholders — replace the `customfield_XXXXX` values with the actual field IDs from your Jira instance.

To find your field IDs:
- Query `/rest/api/2/field` on your Jira instance for a full list
- Use `--list-transitions` to discover valid status transition names
- Query a field's `allowedValues` to find option IDs for select fields like Pending Reason

The retry logic in `transition_ticket()` handles common JSM transition quirks where required fields vary by workflow state — it attempts the full payload first, then falls back gracefully if the instance rejects it.
