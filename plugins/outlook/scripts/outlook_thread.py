#!/usr/bin/env python3
"""READ-ONLY. Returns every message in a conversation (thread), oldest first, with full bodies.

    python outlook_thread.py -Subject "Q3 budget"
    python outlook_thread.py -EntryID 00000000ABCD... -OutFile thread.json
"""
import outlook_com as oc


def find_anchor(a, ns):
    if a.entryid:
        return ns.GetItemFromID(a.entryid)
    if not (a.subject or a.conversationid):
        raise SystemExit("Provide -EntryID, -ConversationID or -Subject.")
    folder = oc.get_folder(a.folder, a.store, ns)
    items = folder.Items
    if a.subject:
        items = items.Restrict(f'@SQL="urn:schemas:httpmail:subject" LIKE \'%{oc.dasl_literal(a.subject)}%\'')
    items.Sort("[ReceivedTime]", True)
    for it in oc.iter_items(items):
        if int(oc._safe(lambda: it.Class, 0)) != oc.OL_MAIL_ITEM:
            continue
        if not a.conversationid or str(oc._safe(lambda: it.ConversationID, "")) == a.conversationid:
            return it
    raise SystemExit("No matching message found.")


def collect(anchor, ns):
    messages, method = [], ""
    try:
        conv = anchor.GetConversation()
        if conv is not None:
            method = "GetConversation"
            table = conv.GetTable()
            while not table.EndOfTable:
                row = table.GetNextRow()
                try:
                    m = ns.GetItemFromID(str(row.Item("EntryID")))
                    if int(m.Class) == oc.OL_MAIL_ITEM:
                        messages.append(m)
                except Exception:
                    pass
    except Exception:
        pass
    if not messages:
        method = "ConversationTopic"
        topic = str(anchor.ConversationTopic)
        root = anchor.Parent.Store.GetRootFolder()
        for f in oc.mail_folders_recursive(root):
            r = f.Items.Restrict(f'@SQL="urn:schemas:httpmail:thread-topic" = \'{oc.dasl_literal(topic)}\'')
            for m in oc.iter_items(r):
                if int(oc._safe(lambda: m.Class, 0)) == oc.OL_MAIL_ITEM:
                    messages.append(m)
    return messages, method


def run(a, ns=None):
    ns = ns or oc.connect()
    anchor = find_anchor(a, ns)
    messages, method = collect(anchor, ns)
    seen, summaries = set(), []
    for m in messages:
        eid = str(m.EntryID)
        if eid in seen:
            continue
        seen.add(eid)
        s = oc.mail_summary(m, include_body=True)
        if len(s["Body"]) > a.maxbodychars:
            s["Body"] = s["Body"][:a.maxbodychars] + "\n[... truncated ...]"
        s["ToRecipients"] = oc.recipient_list(m, 1)
        s["CcRecipients"] = oc.recipient_list(m, 2)
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
    oc.write_json(run(a), a.out_file)


if __name__ == "__main__":
    main()
