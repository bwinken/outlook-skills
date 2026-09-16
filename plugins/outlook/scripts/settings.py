#!/usr/bin/env python3
"""Resolve the plugin's settings and memory files.

Two layers, the closer one wins key by key (deep merge):
    ~/.outlook-skills/settings.json          user level, shared across projects
    <cwd or a parent>/.outlook-skills/settings.json   working-directory level

Each folder may also hold memory/<category>/<title>.md notes Claude keeps for the user
(see memory.py). Both layers are indexed; the working-directory folder is preferred for
new notes when it exists.

Usage:
    python settings.py show                # merged settings + which file set each key + memory paths
    python settings.py init                # create ~/.outlook-skills/ with settings.json ({}), settings.example.json, memory/<category>/
    python settings.py init --local        # same, in ./.outlook-skills/
    python settings.py set working_hours.end 17:30 [--local]
    python settings.py set rerank.auto_consent true
    python settings.py memory              # memory index (use memory.py for details)
    python settings.py profile show        # reply-habit / writing-style profile (profile.md)
    python settings.py profile write --file draft.md [--local]
    python settings.py path                # print the resolved folders

Nothing here touches Outlook. The only writes are to the plugin's own folders.
"""
import argparse
import json
import os
import sys
from pathlib import Path

DIRNAME = ".outlook-skills"
MEMORY_SOFT_LIMIT = 300   # memory files before `show` suggests pruning
CATEGORIES = ["people", "folders", "projects", "preferences", "recurring"]

DEFAULTS = {
    "language": "zh-TW",
    "working_hours": {"start": "09:00", "end": "18:00", "days": [1, 2, 3, 4, 5]},
    "availability": {"min_slot_minutes": 30},
    "search": {
        "default_lookback_days": 90,
        "default_folder": "Inbox",
        "all_folders": False,
        "max_candidates": 300,
        "direct_read_max": 20,
    },
    "store": None,
    "rerank": {
        "gateway": None,
        "model": None,
        "api_key": None,
        "auto_consent": False,
    },
    "status": {"skip_com": False},
    "send": {
        "approver": None,                 # name in the footer; default: the Outlook user's display name
        "footer": "--\nDrafted by Claude, reviewed and approved by {approver}.",
        "quote_original": True,           # replies carry the original message below the footer
        "dialog_timeout_seconds": 300,    # the confirmation window closes as "cancel" after this
    },
    "meeting": {
        "default_duration_minutes": 60,
        "reminder_minutes": 15,
    },
}

MEMORY_TEMPLATE = """# Outlook memory

One Markdown file per topic, in a category folder, with YAML front matter:

    memory/people/alice-chen.md
    ---
    title: Alice Chen
    category: people
    tags: [legal, contoso]
    created: 2026-09-14T10:02:11
    updated: 2026-09-14T10:02:11
    source: user
    ---
    - 法務窗口，alice.chen@contoso.com

Categories: people (人物), folders (資料夾), projects (專案與主題), preferences (偏好),
recurring (定期事務：週報、月結、固定會議). Manage with scripts/memory.py or the
outlook-memory skill. Only aliases, meanings, keywords, decisions and preferences
belong here; never mail bodies, attachments or credentials.
"""


def user_dir() -> Path:
    return Path(os.path.expanduser("~")) / DIRNAME


def local_dir(start=None):
    p = (start or Path.cwd()).resolve()
    home = Path(os.path.expanduser("~")).resolve()
    for d in [p, *p.parents]:
        cand = d / DIRNAME
        if cand.is_dir() and cand != user_dir():
            return cand
        if d == home:
            break
    return None


def _load(path: Path):
    try:
        with open(path, "r", encoding="utf-8-sig") as fh:
            data = json.load(fh)
        return data if isinstance(data, dict) else {}
    except FileNotFoundError:
        return {}
    except Exception as e:
        raise SystemExit(f"Cannot parse {path}: {e}")


def _merge(base, over, src, sources, prefix=""):
    out = dict(base)
    for k, v in over.items():
        key = f"{prefix}{k}"
        if isinstance(v, dict) and isinstance(out.get(k), dict):
            out[k] = _merge(out[k], v, src, sources, key + ".")
        else:
            out[k] = v
            sources[key] = src
    return out


def resolve():
    sources = {}
    merged = json.loads(json.dumps(DEFAULTS))
    layers = []
    u = user_dir() / "settings.json"
    if u.is_file():
        merged = _merge(merged, _load(u), str(u), sources)
        layers.append(str(u))
    ld = local_dir()
    if ld and (ld / "settings.json").is_file():
        lp = ld / "settings.json"
        merged = _merge(merged, _load(lp), str(lp), sources)
        layers.append(str(lp))
    return merged, sources, layers, ld


def cmd_show(args):
    merged, sources, layers, ld = resolve()
    try:
        import memory as pm
        idx = pm.index_summary()
    except Exception as e:
        idx = {"count": 0, "per_category": {}, "entries": [], "error": str(e)}
    hint = None
    if idx["count"] > MEMORY_SOFT_LIMIT:
        hint = f"{idx['count']} memory files (soft limit {MEMORY_SOFT_LIMIT}); suggest pruning stale ones"
    profile = next((str(d / "profile.md") for d in ([ld] if ld else []) + [user_dir()] if (d / "profile.md").is_file()), None)
    print(json.dumps({
        "first_run": not user_dir().exists(),   # true = ~/.outlook-skills has never been created: run outlook-setup
        "profile": profile,                     # writing / reply-habit profile (outlook-setup stage 4), or null
        "settings": merged,
        "sources": sources,           # key -> file that set it (keys absent here are defaults)
        "layers": layers,             # files actually read, in precedence order (later wins)
        "memory": idx,                # index only: title, category, tags, updated, path. Read files on demand with memory.py show
        "memory_hint": hint,
        "local_dir": str(ld) if ld else None,
        "user_dir": str(user_dir()),
    }, ensure_ascii=False, indent=2))


def cmd_init(args):
    target = (Path.cwd() / DIRNAME) if args.local else user_dir()
    target.mkdir(parents=True, exist_ok=True)
    created = []
    sp = target / "settings.json"
    if not sp.exists():
        # Empty on purpose: only keys written here override the layer below.
        sp.write_text("{}\n", encoding="utf-8")
        created.append(str(sp))
    ex = target / "settings.example.json"
    if not ex.exists():
        ex.write_text(json.dumps(DEFAULTS, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        created.append(str(ex))
    for c in CATEGORIES:
        d = target / "memory" / c
        if not d.exists():
            d.mkdir(parents=True)
            created.append(str(d) + os.sep)
    rd = target / "memory" / "README.md"
    if not rd.exists():
        rd.write_text(MEMORY_TEMPLATE, encoding="utf-8")
        created.append(str(rd))
    print(json.dumps({"dir": str(target), "created": created, "already_present": [str(p) for p in (sp, ex) if str(p) not in created]}, ensure_ascii=False, indent=2))


def _coerce(v: str):
    low = v.lower()
    if low in ("true", "false"):
        return low == "true"
    if low in ("null", "none"):
        return None
    try:
        return int(v)
    except ValueError:
        pass
    try:
        return float(v)
    except ValueError:
        pass
    if v.startswith("[") or v.startswith("{"):
        try:
            return json.loads(v)
        except Exception:
            pass
    return v


def cmd_set(args):
    target = ((local_dir() or (Path.cwd() / DIRNAME)) if args.local else user_dir())
    target.mkdir(parents=True, exist_ok=True)
    sp = target / "settings.json"
    data = _load(sp) if sp.exists() else {}
    node = data
    parts = args.key.split(".")
    for p in parts[:-1]:
        node = node.setdefault(p, {})
        if not isinstance(node, dict):
            raise SystemExit(f"{args.key}: '{p}' is not an object")
    node[parts[-1]] = _coerce(args.value)
    sp.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"file": str(sp), "key": args.key, "value": node[parts[-1]]}, ensure_ascii=False))


def cmd_memory(args):
    import memory as pm
    print(json.dumps(pm.index_summary(), ensure_ascii=False, indent=2))


def cmd_profile(args):
    """profile show | profile write --file <md>  : the reply-habit and writing-style profile (profile.md)."""
    target = ((local_dir() or (Path.cwd() / DIRNAME)) if args.local else user_dir()) / "profile.md"
    if args.action == "show":
        print(target.read_text(encoding="utf-8") if target.is_file() else "(no profile.md; run outlook-setup)")
        return
    src = Path(args.file).read_text(encoding="utf-8")
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(src, encoding="utf-8")
    print(json.dumps({"written": str(target), "bytes": len(src.encode("utf-8"))}, ensure_ascii=False))


def cmd_path(args):
    _, _, _, ld = resolve()
    print(json.dumps({"user_dir": str(user_dir()), "local_dir": str(ld) if ld else None}, ensure_ascii=False))


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    sub.add_parser("show").set_defaults(fn=cmd_show)
    p = sub.add_parser("init"); p.add_argument("--local", action="store_true"); p.set_defaults(fn=cmd_init)
    p = sub.add_parser("set"); p.add_argument("key"); p.add_argument("value"); p.add_argument("--local", action="store_true"); p.set_defaults(fn=cmd_set)
    sub.add_parser("memory").set_defaults(fn=cmd_memory)
    p = sub.add_parser("profile"); p.add_argument("action", choices=["show", "write"]); p.add_argument("--file"); p.add_argument("--local", action="store_true"); p.set_defaults(fn=cmd_profile)
    sub.add_parser("path").set_defaults(fn=cmd_path)
    args = ap.parse_args(argv)
    args.fn(args)


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    main()
