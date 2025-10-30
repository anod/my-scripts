#!/usr/bin/env python3
"""
Export tasks from Microsoft To Do into a CSV compatible with Todoist import.

Usage:
  1. Install dependencies:  pip install msal requests
  2. Set environment variables:
        MS_TODO_CLIENT_ID=<your Azure AD app's client id>
        MS_TODO_TENANT_ID=<tenant id or 'common'> (optional, defaults to 'common')
     Your app must have delegated permission Tasks.Read (or Tasks.ReadWrite) granted.
  3. Run:  python ms_todo_to_todoist.py output.csv
  4. Import the generated CSV in Todoist (Web: Settings > Import).

Todoist CSV columns produced:
    Type,Content,Priority,Due Date,Due Time,Description

Priority mapping:
    Microsoft To Do importance (low, normal, high) -> Todoist priority number
        high   -> 1 (highest)
        normal -> 3
        low    -> 4 (lowest)

Authentication:
    Uses Device Code flow via MSAL. You will be prompted to visit a URL and enter a code.

Notes:
    - All lists are exported. The list name is appended to Description.
    - Tasks with subtasks: Each checklist item becomes a separate task line (prefixed by parent task title) unless --no-checklists is set.
    - If a task is flagged/has "hasAttachments" this is noted in the description.
    - If due date has time component it will appear in Due Time (HH:MM in 24h). Otherwise Due Time is blank.
    - Completed tasks are skipped by default; use --include-completed to include them.

"""
import os
import sys
import csv
import datetime as dt
import argparse
from typing import List, Dict, Any

import requests
import msal

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SCOPES = ["Tasks.Read"]  # or Tasks.ReadWrite if needed

PRIORITY_MAP = {
    "high": 1,
    "normal": 3,
    "low": 4,
}

def get_token() -> str:
    client_id = os.getenv("MS_TODO_CLIENT_ID")
    if not client_id:
        print("ERROR: MS_TODO_CLIENT_ID environment variable not set", file=sys.stderr)
        sys.exit(1)
    tenant_id = os.getenv("MS_TODO_TENANT_ID", "common")
    authority = f"https://login.microsoftonline.com/{{tenant_id}}"
    app = msal.PublicClientApplication(client_id=client_id, authority=authority)

    # Try silent first (will almost always fail first run)
    accounts = app.get_accounts()
    result = None
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])

    if not result:
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            print("Failed to create device flow", file=sys.stderr)
            sys.exit(1)
        print("To authenticate, navigate to:", flow["verification_uri"])
        print("Enter the code:", flow["user_code"])
        result = app.acquire_token_by_device_flow(flow)  # blocks

    if "access_token" not in result:
        print("Failed to obtain access token:", result.get("error_description"), file=sys.stderr)
        sys.exit(1)
    return result["access_token"]

def graph_get(token: str, url: str) -> Dict[str, Any]:
    headers = {"Authorization": f"Bearer {{token}}"}
    r = requests.get(url, headers=headers)
    if r.status_code >= 400:
        raise RuntimeError(f"Graph API error {{r.status_code}}: {{r.text}}")
    return r.json()

def fetch_lists(token: str) -> List[Dict[str, Any]]:
    data = graph_get(token, f"{{GRAPH_BASE}}/me/todo/lists")
    return data.get("value", [])

def fetch_tasks(token: str, list_id: str) -> List[Dict[str, Any]]:
    # Expand checklist items
    data = graph_get(token, f"{{GRAPH_BASE}}/me/todo/lists/{{list_id}}/tasks?$expand=checklistItems")
    return data.get("value", [])

def parse_due(due: Dict[str, Any]):
    if not due or not due.get("dateTime"):
        return "", ""
    # dateTime is ISO 8601 in user's time zone (timeZone field). We'll parse naive and output date/time.
    try:
        raw = due["dateTime"]
        # Some values may lack time; attempt parse
        dt_obj = dt.datetime.fromisoformat(raw.replace("Z", "+00:00"))
        date_str = dt_obj.date().isoformat()
        time_str = dt_obj.time().strftime("%H:%M") if dt_obj.time() != dt.time(0, 0) else ""
        return date_str, time_str
    except Exception:
        return "", ""

def importance_to_priority(importance: str) -> int:
    return PRIORITY_MAP.get(importance, 3)

def sanitize(text: str) -> str:
    if text is None:
        return ""
    return text.replace("\r", " ").replace("\n", " ").strip()

def task_rows(task: Dict[str, Any], list_name: str, include_checklists: bool, parent_prefix: str = "") -> List[List[str]]:
    rows = []
    if task.get("status") == "completed":
        # Caller filters; keep anyway if asked
        pass
    title = sanitize(task.get("title", ""))
    importance = task.get("importance", "normal")
    priority = importance_to_priority(importance)
    due_date, due_time = parse_due(task.get("dueDateTime"))
    body = sanitize(task.get("body", {}).get("content", ""))
    notes = []
    if body:
        notes.append(body)
    if task.get("hasAttachments"):
        notes.append("[Has Attachments]")
    notes.append(f"List: {{list_name}}")
    description = " 
".join(notes)

    content = f"{{parent_prefix}}{{title}}" if parent_prefix else title
    rows.append(["task", content, str(priority), due_date, due_time, description])

    if include_checklists:
        cl_items = task.get("checklistItems", []) or []
        for item in cl_items:
            if item.get("isChecked"):
                continue  # skip completed checklist items (importing them as tasks usually not desired)
            cl_title = sanitize(item.get("displayName", ""))
            rows.append(["task", f"{{title}} > {{cl_title}}", str(priority), due_date, due_time, f"Subtask from '{{title}}' | List: {{list_name}}" ])
    return rows

def export_csv(token: str, output_path: str, include_completed: bool, no_checklists: bool):
    lists = fetch_lists(token)
    all_rows: List[List[str]] = []
    for lst in lists:
        list_name = lst.get("displayName", "Unnamed List")
        tasks = fetch_tasks(token, lst.get("id"))
        for t in tasks:
            if not include_completed and t.get("status") == "completed":
                continue
            rows = task_rows(t, list_name, include_checklists=not no_checklists)
            all_rows.extend(rows)
    # Write CSV
    with open(output_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["Type", "Content", "Priority", "Due Date", "Due Time", "Description"])
        writer.writerows(all_rows)
    print(f"Exported {{len(all_rows)}} tasks to {{output_path}}")

def main():
    parser = argparse.ArgumentParser(description="Export Microsoft To Do tasks to Todoist CSV format")
    parser.add_argument("output", help="Output CSV filename")
    parser.add_argument("--include-completed", action="store_true", help="Include completed tasks")
    parser.add_argument("--no-checklists", action="store_true", help="Do not export checklist items as separate tasks")
    args = parser.parse_args()

    token = get_token()
    export_csv(token, args.output, include_completed=args.include_completed, no_checklists=args.no_checklists)

if __name__ == "__main__":
    main()