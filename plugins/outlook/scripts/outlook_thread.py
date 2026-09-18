#!/usr/bin/env python3
"""READ-ONLY. Returns every message in a conversation (thread), oldest first, with full bodies.

    python outlook_thread.py -Subject "Q3 budget"
    python outlook_thread.py -EntryID 00000000ABCD... -OutFile thread.json

The anchor and the topic fallback are found through Outlook's Table object; the messages of the
thread are opened one by one (their bodies are the point of this script).
"""
import outlook_com as oc


def find_anchor(a, ns):
    if a.entryid:
        return ns.GetItemFromID(a.entryid)
    if not (a.subject or a.conversationid):
        raise SystemExit("Provide -EntryID, -ConversationID or -Subject.")
    folder = oc.get_folder(a.folder, a.store, ns)
    dasl = f'@SQL="urn:schemas:httpmail:subject" LIKE \'%{oc.dasl_literal(a.subject)}%\'' if a.subject else ""
    for row in oc.scan_mail(folder, dasl):
        if not a.conversationid or row.summary["ConversationID"] == a.conversationid:
            return oc.open_item(row, ns)
    raise SystemExit("No matching message found.")


def collect(anchor, ns):
    """The thread's mail items, how they were found, and the folder path of each item where known."""
    messages, method, paths = [], "", {}
    try:
        conv = anchor.GetConversation()
        if conv is not None:
            method = "GetConversation"
            table = conv.GetTable()
            while not table.EndOfTable:
                row = table.GetNextRow()
                try:
                    cls = str(oc._safe(lambda: row.Item("MessageClass"), "") or "")
                    if cls and not cls.upper().startswith("IPM.NOTE"):
                        continue  # a meeting request or report in the thread: not opened
                    m = ns.GetItemFromID(str(row.Item("EntryID")))
                    if cls or int(m.Class) == oc.OL_MAIL_ITEM:
                        messages.append(m)
                except Exception:
                    pass
    except Exception:
        pass
    if not messages:
        method = "ConversationTopic"
        topic = str(anchor.ConversationTopic)
        root = anchor.Parent.Store.GetRootFolder()
        dasl = f'@SQL="urn:schemas:httpmail:thread-topic" = \'{oc.dasl_literal(topic)}\''
        for f in oc.mail_folders_recursive(root):
            for row in oc.scan_mail(f, dasl):
                paths[row.summary["EntryID"]] = row.summary["Folder"]
                messages.append(oc.open_item(row, ns))
    return messages, method, paths


def run(a, ns=None):
    ns = ns or oc.connect()
    anchor = find_anchor(a, ns)
    messages, method, paths = collect(anchor, ns)
    seen, summaries = set(), []
    for m in messages:
        eid = str(m.EntryID)
        if eid in seen:
            continue
        seen.add(eid)
        s = oc.mail_summary(m, include_body=True, folder_path=paths.get(eid))
        if len(s["Body"]) > a.maxbodychars:
            s["Body"] = s["Body"][:a.maxbodychars] + "\n[... truncated ...]"
        recipients = oc.recipients_by_type(m)  # one pass over the Recipients collection
        s["ToRecipients"], s["CcRecipients"] = recipients[1], recipients[2]
        summaries.append(s)
    summaries.sort(key=lambda x: x["ReceivedTime"] or "")
    return {
        "Anchor": str(anchor.EntryID), "Topic": str(anchor.ConversationTopic),
        "ConversationID": str(oc._safe(lambda: anchor.ConversationID, "")), "Method": method,
        "Count": len(summaries),
        "Participants": sorted({s["FromAddress"] for s in summaries if s["FromAddress"]}),
        "Messages": summaries,
    }


def parser():
    ap = oc.ArgParser(description=__doc__)
    ap.opt("-EntryID", default="")
    ap.opt("-ConversationID", default="")
    ap.opt("-Subject", default="", help="substring; newest match in -Folder anchors the thread")
    ap.opt("-Folder", default="")
    ap.opt("-Store", default="")
    ap.opt("-MaxBodyChars", type=int, default=20000)
    oc.add_common_output(ap)
    return ap


def main(argv=None):
    a = parser().parse_args(argv)
    oc.apply_settings(a)
    oc.write_json(run(a), a.out_file)


if __name__ == "__main__":
    main()
