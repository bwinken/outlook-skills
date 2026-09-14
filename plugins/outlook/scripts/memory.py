#!/usr/bin/env python3
"""Memory files for the Outlook skills: one Markdown file per fact cluster, with YAML front matter.

Layout (user level; a working-directory .outlook-skills/ may hold the same tree and is read too):
    ~/.outlook-skills/memory/<category>/<slug>.md

    ---
    title: Alice Chen
    category: people
    tags: [legal, contoso]
    created: 2026-09-14T10:02:11
    updated: 2026-09-14T10:02:11
    source: bootstrap          # or "user"
    ---
    - 法務窗口，alice.chen@contoso.com
    - 合約相關的信都由她發起

Categories: people, folders, projects, preferences, recurring (others allowed).

Usage:
    python memory.py list [--category people] [--json]
    python memory.py find "alice" [--json]                 # matches title, tags, body (case-insensitive)
    python memory.py new --category people --title "Alice Chen" --tags legal,contoso --body "- 法務窗口" [--source bootstrap] [--local]
    python memory.py append <path-or-title> --body "- 新事實"   # appends bullet(s), bumps updated
    python memory.py touch <path-or-title>                      # bump updated after editing the body by hand
    python memory.py show <path-or-title>
    python memory.py remove <path-or-title>

Only .outlook-skills folders are written. Outlook is never touched.
"""
import argparse
import datetime as dt
import json
import os
import re
import sys
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import settings as ps  # noqa: E402

CATEGORIES = ["people", "folders", "projects", "preferences", "recurring"]


# ---------------------------------------------------------------- front matter
def _now():
    return dt.datetime.now().replace(microsecond=0).isoformat()


def _dump_fm(meta: dict) -> str:
    lines = ["---"]
    for k in ("title", "category", "tags", "created", "updated", "source"):
        if k not in meta:
            continue
        v = meta[k]
        if k == "tags":
            lines.append("tags: [" + ", ".join(str(t) for t in v) + "]")
        else:
            sv = str(v)
            if k not in ("created", "updated") and (re.search(r'[:#\[\]{}"\']', sv) or sv != sv.strip()):
                sv = json.dumps(sv, ensure_ascii=False)
            lines.append(f"{k}: {sv}")
    for k, v in meta.items():
        if k not in ("title", "category", "tags", "created", "updated", "source"):
            lines.append(f"{k}: {json.dumps(v, ensure_ascii=False) if not isinstance(v, str) else v}")
    lines.append("---")
    return "\n".join(lines) + "\n"


def _parse_fm(text: str):
    if not text.startswith("---"):
        return {}, text
    m = re.match(r"^---\n(.*?)\n---\n?", text, re.S)
    if not m:
        return {}, text
    meta = {}
    for line in m.group(1).splitlines():
        if ":" not in line:
            continue
        k, v = line.split(":", 1)
        k, v = k.strip(), v.strip()
        if k == "tags":
            v = v.strip("[]")
            meta[k] = [t.strip().strip('"').strip("'") for t in v.split(",") if t.strip()]
        else:
            if v.startswith('"') and v.endswith('"'):
                try:
                    v = json.loads(v)
                except Exception:
                    v = v[1:-1]
            meta[k] = v
    return meta, text[m.end():]


def slugify(title: str) -> str:
    s = title.strip().lower()
    s = re.sub(r"[\\/:*?\"<>|#%&{}$!'@+`=]", " ", s)
    s = re.sub(r"[\s_.,;()\[\]]+", "-", s).strip("-")
    return s or "untitled"


# ---------------------------------------------------------------- store
def roots(local_first=False):
    ld = ps.local_dir()
    r = [ps.user_dir()]
    if ld:
        r.append(ld)
    return list(reversed(r)) if local_first else r


def all_files():
    out = []
    for root in roots():
        md = root / "memory"
        if not md.is_dir():
            continue
        for p in sorted(md.glob("*/*.md")):   # only <category>/<title>.md; memory/README.md is not a note
            try:
                text = p.read_text(encoding="utf-8")
            except Exception:
                continue
            meta, body = _parse_fm(text)
            meta.setdefault("title", p.stem)
            meta.setdefault("category", p.parent.name if p.parent != md else "")
            meta.setdefault("tags", [])
            out.append({"path": str(p), "scope": "local" if root != ps.user_dir() else "user", "meta": meta, "body": body})
    return out


def resolve_ref(ref: str):
    p = Path(ref)
    if p.is_file():
        return p
    low = ref.lower()
    hits = [f for f in all_files() if f["meta"]["title"].lower() == low or Path(f["path"]).stem == slugify(ref)]
    if len(hits) == 1:
        return Path(hits[0]["path"])
    if not hits:
        raise SystemExit(f"No memory file matches '{ref}'. Use `memory.py find` to look it up.")
    raise SystemExit("Ambiguous: " + ", ".join(h["path"] for h in hits))


def index_summary():
    files = all_files()
    per = {}
    for f in files:
        per[f["meta"]["category"]] = per.get(f["meta"]["category"], 0) + 1
    return {"count": len(files), "per_category": per,
            "entries": [{"title": f["meta"]["title"], "category": f["meta"]["category"], "tags": f["meta"].get("tags", []),
                         "updated": f["meta"].get("updated"), "scope": f["scope"], "path": f["path"]} for f in files]}


# ---------------------------------------------------------------- commands
def cmd_list(a):
    files = all_files()
    if a.category:
        files = [f for f in files if f["meta"]["category"] == a.category]
    if a.json:
        print(json.dumps([{"path": f["path"], "scope": f["scope"], **f["meta"]} for f in files], ensure_ascii=False, indent=2))
        return
    if not files:
        print("(no memory files)")
        return
    for f in files:
        m = f["meta"]
        print(f"[{m['category']}] {m['title']}  tags={','.join(m.get('tags', []))}  updated={m.get('updated', '?')}  ({f['scope']}) {f['path']}")


def cmd_find(a):
    q = a.query.lower()
    hits = []
    for f in all_files():
        m = f["meta"]
        hay = " ".join([m["title"], " ".join(m.get("tags", [])), f["body"]]).lower()
        if q in hay:
            where = "title" if q in m["title"].lower() else ("tags" if q in " ".join(m.get("tags", [])).lower() else "body")
            hits.append({"path": f["path"], "title": m["title"], "category": m["category"], "tags": m.get("tags", []), "match": where})
    if a.json:
        print(json.dumps(hits, ensure_ascii=False, indent=2))
    elif not hits:
        print("(no match)")
    else:
        for h in hits:
            print(f"[{h['category']}] {h['title']}  (match: {h['match']})  {h['path']}")


def cmd_new(a):
    root = (ps.local_dir() or (Path.cwd() / ps.DIRNAME)) if a.local else ps.user_dir()
    d = root / "memory" / a.category
    d.mkdir(parents=True, exist_ok=True)
    p = d / (slugify(a.title) + ".md")
    if p.exists() and not a.force:
        raise SystemExit(f"Exists: {p}. Use `append` to add to it, or --force to overwrite.")
    now = _now()
    meta = {"title": a.title, "category": a.category, "tags": [t.strip() for t in (a.tags or "").split(",") if t.strip()],
            "created": now, "updated": now, "source": a.source}
    body = (a.body or "").rstrip() + "\n"
    p.write_text(_dump_fm(meta) + body, encoding="utf-8")
    print(json.dumps({"path": str(p), **meta}, ensure_ascii=False))


def cmd_append(a):
    p = resolve_ref(a.ref)
    meta, body = _parse_fm(p.read_text(encoding="utf-8"))
    meta["updated"] = _now()
    if a.tags:
        for t in a.tags.split(","):
            t = t.strip()
            if t and t not in meta.setdefault("tags", []):
                meta["tags"].append(t)
    body = body.rstrip() + "\n" + a.body.rstrip() + "\n"
    p.write_text(_dump_fm(meta) + body, encoding="utf-8")
    print(json.dumps({"updated": str(p), "title": meta.get("title"), "tags": meta.get("tags", [])}, ensure_ascii=False))


def cmd_touch(a):
    p = resolve_ref(a.ref)
    meta, body = _parse_fm(p.read_text(encoding="utf-8"))
    meta["updated"] = _now()
    p.write_text(_dump_fm(meta) + body, encoding="utf-8")
    print(json.dumps({"touched": str(p), "updated": meta["updated"]}, ensure_ascii=False))


def cmd_show(a):
    p = resolve_ref(a.ref)
    print(p.read_text(encoding="utf-8"))


def cmd_remove(a):
    p = resolve_ref(a.ref)
    p.unlink()
    print(json.dumps({"removed": str(p)}, ensure_ascii=False))


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    p = sub.add_parser("list"); p.add_argument("--category"); p.add_argument("--json", action="store_true"); p.set_defaults(fn=cmd_list)
    p = sub.add_parser("find"); p.add_argument("query"); p.add_argument("--json", action="store_true"); p.set_defaults(fn=cmd_find)
    p = sub.add_parser("new"); p.add_argument("--category", required=True); p.add_argument("--title", required=True)
    p.add_argument("--tags", default=""); p.add_argument("--body", default=""); p.add_argument("--source", default="user")
    p.add_argument("--local", action="store_true"); p.add_argument("--force", action="store_true"); p.set_defaults(fn=cmd_new)
    p = sub.add_parser("append"); p.add_argument("ref"); p.add_argument("--body", required=True); p.add_argument("--tags", default=""); p.set_defaults(fn=cmd_append)
    p = sub.add_parser("touch"); p.add_argument("ref"); p.set_defaults(fn=cmd_touch)
    p = sub.add_parser("show"); p.add_argument("ref"); p.set_defaults(fn=cmd_show)
    p = sub.add_parser("remove"); p.add_argument("ref"); p.set_defaults(fn=cmd_remove)
    a = ap.parse_args(argv)
    a.fn(a)


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    main()
