"""
outlook_bridge.py — Outlook COM bridge for Claude Code
Requires: pywin32 (pip install pywin32), Outlook Desktop App installed and signed in.

Usage:
  python outlook_bridge.py folders
  python outlook_bridge.py list [--folder INBOX] [--count 20]
  python outlook_bridge.py search "keyword" [--folder INBOX] [--count 20]
  python outlook_bridge.py read <entry_id>
  python outlook_bridge.py send --to addr [--cc addr] --subject "..." --body "..." [--html]
  python outlook_bridge.py reply <entry_id> --body "..." [--html] [--all] [--draft]
  python outlook_bridge.py forward <entry_id> --to addr [--cc addr] --body "..." [--draft]
  python outlook_bridge.py move <entry_id> --folder "Folder Name"
  python outlook_bridge.py delete <entry_id>
  python outlook_bridge.py cal-list --from YYYY-MM-DD --to YYYY-MM-DD [--cal personal|<name>]

Output: JSON (use --pretty for human-readable)

Shared calendars are configured in outlook_bridge_config.json (same directory as this script,
or ~/.claude/outlook_bridge_config.json). See outlook_bridge_config.json for format.
"""

import argparse
import json
import os
import sys
from datetime import datetime

# Force UTF-8 output so non-ASCII characters in email bodies don't crash on Windows
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

try:
    import win32com.client
    import pywintypes
except ImportError:
    print(json.dumps({"error": "pywin32 not installed. Run: pip install pywin32"}))
    sys.exit(1)

# Outlook default folder constants
FOLDER_INBOX = 6
FOLDER_SENT = 5
FOLDER_DRAFTS = 16
FOLDER_DELETED = 3

FOLDER_NAMES = {
    "inbox": FOLDER_INBOX,
    "sent": FOLDER_SENT,
    "drafts": FOLDER_DRAFTS,
    "deleted": FOLDER_DELETED,
    "trash": FOLDER_DELETED,
}


def load_config():
    """Load shared_calendars from config file. Returns dict (may be empty)."""
    candidates = [
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "outlook_bridge_config.json"),
        os.path.expanduser("~/.claude/outlook_bridge_config.json"),
    ]
    for path in candidates:
        if os.path.exists(path):
            try:
                with open(path, encoding="utf-8") as f:
                    data = json.load(f)
                return data.get("shared_calendars", {})
            except Exception:
                pass
    return {}


def get_outlook():
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        ns = outlook.GetNamespace("MAPI")
        return outlook, ns
    except Exception as e:
        err(f"Cannot connect to Outlook: {e}")


def err(msg):
    print(json.dumps({"error": str(msg)}))
    sys.exit(1)


def fmt_time(t):
    if t is None:
        return None
    try:
        return str(t)[:19]  # "YYYY-MM-DD HH:MM:SS"
    except Exception:
        return str(t)


def mail_summary(item):
    try:
        return {
            "id": item.EntryID,
            "subject": item.Subject or "(no subject)",
            "from": item.SenderEmailAddress or item.SenderName,
            "from_name": item.SenderName,
            "to": item.To,
            "cc": item.CC,
            "received": fmt_time(item.ReceivedTime),
            "sent": fmt_time(item.SentOn),
            "unread": item.UnRead,
            "has_attachments": item.Attachments.Count > 0,
            "size": item.Size,
        }
    except Exception as e:
        return {"error": str(e)}


def mail_full(item):
    s = mail_summary(item)
    try:
        s["body"] = item.Body
        s["attachments"] = [
            {"name": a.FileName, "size": a.Size}
            for a in item.Attachments
        ]
    except Exception as e:
        s["body_error"] = str(e)
    return s


def resolve_folder(ns, folder_name):
    """Resolve folder name: default names or walk full path like 'Projects/Client'."""
    key = folder_name.strip().lower()
    if key in FOLDER_NAMES:
        return ns.GetDefaultFolder(FOLDER_NAMES[key])

    # Walk path segments from inbox root
    inbox = ns.GetDefaultFolder(FOLDER_INBOX)
    root = inbox.Parent  # mailbox root

    parts = folder_name.replace("\\", "/").split("/")
    current = root
    for part in parts:
        found = None
        try:
            for f in current.Folders:
                if f.Name.lower() == part.lower():
                    found = f
                    break
        except Exception:
            pass
        if found is None:
            # Try from inbox
            try:
                for f in inbox.Folders:
                    if f.Name.lower() == part.lower():
                        found = f
                        break
            except Exception:
                pass
        if found is None:
            err(f"Folder not found: '{part}' in path '{folder_name}'")
        current = found
    return current


def cmd_folders(ns, args, pretty):
    """List all folders recursively."""
    def walk(folder, depth=0):
        result = []
        try:
            for f in folder.Folders:
                result.append({
                    "name": f.Name,
                    "path": "  " * depth + f.Name,
                    "count": f.Items.Count,
                    "unread": f.UnReadItemCount,
                })
                result.extend(walk(f, depth + 1))
        except Exception:
            pass
        return result

    inbox = ns.GetDefaultFolder(FOLDER_INBOX)
    root = inbox.Parent
    folders = walk(root)
    output({"folders": folders}, pretty)


def cmd_list(ns, args, pretty):
    folder = resolve_folder(ns, args.folder)
    items = folder.Items
    items.Sort("[ReceivedTime]", True)  # newest first

    count = min(args.count, items.Count)
    results = []
    mail_count = 0
    for i in range(1, count + 1):
        try:
            item = items[i]
            if item.Class == 43:  # olMail
                results.append(mail_summary(item))
                mail_count += 1
        except Exception:
            pass

    output({"folder": folder.Name, "total": mail_count, "showing": len(results), "emails": results}, pretty)


def cmd_search(ns, args, pretty):
    folder = resolve_folder(ns, args.folder)
    items = folder.Items
    items.Sort("[ReceivedTime]", True)

    q = args.query.lower()
    results = []
    limit = args.count * 10  # scan up to 10x to find matches

    for i in range(1, min(items.Count, limit) + 1):
        if len(results) >= args.count:
            break
        try:
            item = items[i]
            if item.Class != 43:
                continue
            subj = (item.Subject or "").lower()
            body = (item.Body or "").lower()
            sender = (item.SenderEmailAddress or "").lower()
            sender_name = (item.SenderName or "").lower()
            to = (item.To or "").lower()
            if q in subj or q in body or q in sender or q in sender_name or q in to:
                results.append(mail_summary(item))
        except Exception:
            pass

    output({"query": args.query, "folder": folder.Name, "found": len(results), "emails": results}, pretty)


def cmd_read(ns, args, pretty):
    try:
        item = ns.GetItemFromID(args.entry_id)
        output(mail_full(item), pretty)
    except Exception as e:
        err(f"Cannot read email (ID may be invalid): {e}")


def cmd_send(outlook, args, pretty):
    try:
        mail = outlook.CreateItem(0)  # olMailItem
        mail.Subject = args.subject or ""
        if args.html:
            mail.HTMLBody = args.body
        else:
            mail.Body = args.body

        for addr in args.to:
            r = mail.Recipients.Add(addr.strip())
            r.Type = 1  # olTo

        if args.cc:
            for addr in args.cc:
                r = mail.Recipients.Add(addr.strip())
                r.Type = 2  # olCC

        mail.Recipients.ResolveAll()
        mail.Send()
        output({"status": "sent", "subject": args.subject, "to": args.to}, pretty)
    except Exception as e:
        err(f"Send failed: {e}")


def cmd_reply(ns, args, pretty):
    try:
        item = ns.GetItemFromID(args.entry_id)
        reply = item.ReplyAll() if args.all else item.Reply()
        html_body = "<div>" + args.body.replace("\n", "<br>") + "</div><br>"
        reply.HTMLBody = html_body + reply.HTMLBody
        if hasattr(args, 'attach_email') and args.attach_email:
            import tempfile
            for eid in args.attach_email:
                attached_item = ns.GetItemFromID(eid)
                tmp = tempfile.mktemp(suffix=".msg")
                attached_item.SaveAs(tmp, 3)
                reply.Attachments.Add(tmp)
                os.remove(tmp)
        if args.draft:
            reply.Save()
            output({"status": "saved_to_drafts", "original_subject": item.Subject, "reply_all": args.all}, pretty)
        else:
            reply.Send()
            output({"status": "replied", "original_subject": item.Subject, "reply_all": args.all}, pretty)
    except Exception as e:
        err(f"Reply failed: {e}")


def cmd_forward(ns, outlook, args, pretty):
    try:
        item = ns.GetItemFromID(args.entry_id)
        fwd = item.Forward()
        if args.body:
            html_body = "<div>" + args.body.replace("\n", "<br>") + "</div><br>"
            fwd.HTMLBody = html_body + fwd.HTMLBody

        for addr in args.to:
            r = fwd.Recipients.Add(addr.strip())
            r.Type = 1

        if args.cc:
            for addr in args.cc:
                r = fwd.Recipients.Add(addr.strip())
                r.Type = 2

        if args.draft:
            fwd.Save()
            output({"status": "saved_to_drafts", "original_subject": item.Subject}, pretty)
        else:
            fwd.Recipients.ResolveAll()
            fwd.Send()
            output({"status": "forwarded", "original_subject": item.Subject}, pretty)
    except Exception as e:
        err(f"Forward failed: {e}")


def cmd_move(ns, args, pretty):
    try:
        item = ns.GetItemFromID(args.entry_id)
        dest = resolve_folder(ns, args.folder)
        item.Move(dest)
        output({"status": "moved", "folder": dest.Name}, pretty)
    except Exception as e:
        err(f"Move failed: {e}")


def get_calendar_folder(ns, cal, shared_calendars):
    """Resolve calendar: 'personal' = default Outlook calendar; any other name = shared calendar from config."""
    if cal == "personal":
        return ns.GetDefaultFolder(9)  # olFolderCalendar

    email = shared_calendars.get(cal)
    if not email:
        err(f"Shared calendar '{cal}' not found in config. Check outlook_bridge_config.json.")

    try:
        recip = ns.CreateRecipient(email)
        recip.Resolve()
        if not recip.Resolved:
            err(f"Could not resolve shared calendar '{cal}' ({email}) — ensure you have access in Outlook.")
        return ns.GetSharedDefaultFolder(recip, 9)
    except Exception as e:
        err(f"Cannot access shared calendar '{cal}': {e}")


def cmd_cal_list(ns, args, shared_calendars, pretty):
    """List calendar events in a date range."""
    try:
        from datetime import timedelta
        cal_folder = get_calendar_folder(ns, args.cal, shared_calendars)
        items = cal_folder.Items
        items.IncludeRecurrences = True
        items.Sort("[Start]")

        try:
            from_dt = datetime.strptime(args.from_date, "%Y-%m-%d")
            to_dt = datetime.strptime(args.to_date, "%Y-%m-%d").replace(
                hour=23, minute=59, second=59
            )
        except ValueError as e:
            err(f"Invalid date format (use YYYY-MM-DD): {e}")

        fmt = "%m/%d/%Y %H:%M"
        restriction = (
            f"[Start] >= '{from_dt.strftime(fmt)}' AND [Start] <= '{to_dt.strftime(fmt)}'"
        )
        filtered = items.Restrict(restriction)

        results = []
        for item in filtered:
            try:
                results.append({
                    "id": item.EntryID,
                    "subject": item.Subject or "(no subject)",
                    "start": fmt_time(item.Start),
                    "end": fmt_time(item.End),
                    "location": item.Location or "",
                    "all_day": item.AllDayEvent,
                    "is_recurring": item.IsRecurring,
                    "categories": item.Categories or "",
                    "body_preview": (item.Body or "")[:200],
                })
            except Exception as e:
                results.append({"error": str(e)})

        output({
            "calendar": args.cal,
            "from": args.from_date,
            "to": args.to_date,
            "count": len(results),
            "events": results,
        }, pretty)

    except Exception as e:
        err(f"cal-list failed: {e}")


def cmd_delete(ns, args, pretty):
    try:
        item = ns.GetItemFromID(args.entry_id)
        subj = item.Subject
        item.Delete()
        output({"status": "deleted", "subject": subj}, pretty)
    except Exception as e:
        err(f"Delete failed: {e}")


def output(data, pretty):
    if pretty:
        print(json.dumps(data, ensure_ascii=False, indent=2, default=str))
    else:
        print(json.dumps(data, ensure_ascii=False, default=str))


def main():
    shared_calendars = load_config()

    # Strip --pretty from argv early so it works before or after subcommand
    pretty_flag = "--pretty" in sys.argv
    if pretty_flag:
        sys.argv.remove("--pretty")

    cal_choices = ["personal"] + list(shared_calendars.keys())

    parser = argparse.ArgumentParser(description="Outlook COM bridge for Claude Code")
    sub = parser.add_subparsers(dest="cmd")

    sub.add_parser("folders", help="List all folders")

    p_list = sub.add_parser("list", help="List recent emails")
    p_list.add_argument("--folder", default="inbox")
    p_list.add_argument("--count", type=int, default=20)

    p_search = sub.add_parser("search", help="Search emails")
    p_search.add_argument("query")
    p_search.add_argument("--folder", default="inbox")
    p_search.add_argument("--count", type=int, default=20)

    p_read = sub.add_parser("read", help="Read full email by EntryID")
    p_read.add_argument("entry_id")

    p_send = sub.add_parser("send", help="Send a new email")
    p_send.add_argument("--to", nargs="+", required=True)
    p_send.add_argument("--cc", nargs="+")
    p_send.add_argument("--subject", required=True)
    p_send.add_argument("--body", required=True)
    p_send.add_argument("--html", action="store_true")

    p_reply = sub.add_parser("reply", help="Reply to an email")
    p_reply.add_argument("entry_id")
    p_reply.add_argument("--body", required=True)
    p_reply.add_argument("--html", action="store_true")
    p_reply.add_argument("--all", action="store_true", help="Reply all")
    p_reply.add_argument("--draft", action="store_true", help="Save to Drafts instead of sending")
    p_reply.add_argument("--attach-email", nargs="+", dest="attach_email", help="Entry IDs of emails to attach as .msg files")

    p_fwd = sub.add_parser("forward", help="Forward an email")
    p_fwd.add_argument("entry_id")
    p_fwd.add_argument("--to", nargs="+", required=True)
    p_fwd.add_argument("--cc", nargs="+")
    p_fwd.add_argument("--body", default="")
    p_fwd.add_argument("--draft", action="store_true", help="Save to Drafts instead of sending")

    p_move = sub.add_parser("move", help="Move email to folder")
    p_move.add_argument("entry_id")
    p_move.add_argument("--folder", required=True)

    p_del = sub.add_parser("delete", help="Move email to Deleted Items")
    p_del.add_argument("entry_id")

    p_cal_list = sub.add_parser("cal-list", help="List calendar events in a date range")
    p_cal_list.add_argument("--from", dest="from_date", required=True, metavar="YYYY-MM-DD")
    p_cal_list.add_argument("--to", dest="to_date", required=True, metavar="YYYY-MM-DD")
    p_cal_list.add_argument(
        "--cal", default="personal", choices=cal_choices,
        help=f"Calendar to query: personal (default){', or a shared calendar name from your config' if shared_calendars else ''}"
    )

    args = parser.parse_args()
    if not args.cmd:
        parser.print_help()
        sys.exit(0)

    outlook, ns = get_outlook()
    pretty = pretty_flag

    if args.cmd == "folders":
        cmd_folders(ns, args, pretty)
    elif args.cmd == "list":
        cmd_list(ns, args, pretty)
    elif args.cmd == "search":
        cmd_search(ns, args, pretty)
    elif args.cmd == "read":
        cmd_read(ns, args, pretty)
    elif args.cmd == "send":
        cmd_send(outlook, args, pretty)
    elif args.cmd == "reply":
        cmd_reply(ns, args, pretty)
    elif args.cmd == "forward":
        cmd_forward(ns, outlook, args, pretty)
    elif args.cmd == "move":
        cmd_move(ns, args, pretty)
    elif args.cmd == "delete":
        cmd_delete(ns, args, pretty)
    elif args.cmd == "cal-list":
        cmd_cal_list(ns, args, shared_calendars, pretty)


if __name__ == "__main__":
    main()
