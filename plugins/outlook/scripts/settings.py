#!/usr/bin/env python3
"""Resolve the plugin's settings and memory files.

Two layers, the closer one wins key by key (deep merge):
    ~/.outlook-skills/settings.json          user level, shared across projects
    <cwd or a parent>/.outlook-skills/settings.json   working-directory level

Each folder may also hold memory.md, free-form notes Claude keeps for the user
(contact aliases, folder meanings, project keywords, preferences). Both memory
files are read; the working-directory one is preferred for new notes when it exists.

Usage:
    python settings.py show                # merged settings + which file set each key + memory paths
    python settings.py init                # create ~/.outlook-skills/ with an empty settings.json, settings.example.json (all keys), memory.md
    python settings.py init --local        # same, in ./.outlook-skills/
    python settings.py set working_hours.end 17:30 [--local]
    python settings.py set rerank.auto_consent true
    python settings.py memory              # print the memory files' contents
    python settings.py path                # print the resolved folders

Nothing here touches Outlook. The only writes are to the plugin's own folders.
"""
import argparse
import json
import os
import sys
from pathlib import Path

DIRNAME = ".outlook-skills"

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
}

MEMORY_TEMPLATE = """# Outlook memory

Notes Claude keeps for the Outlook skills. Only aliases, meanings, keywords,
decisions and preferences go here; never mail bodies or attachments.

## 人物與別名
<!-- - Alice = Alice Chen <alice.chen@contoso.com>，法務窗口 -->

## 資料夾
<!-- - Inbox/Vendors：供應商往來 -->

## 專案關鍵字
<!-- - Q3 預算：方案 B，1.5M，VP review 9/19 -->

## 偏好
<!-- - 表格日期用 MM/dd；回覆用繁體中文 -->
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
    memory = [str(p) for p in [user_dir() / "memory.md", (ld / "memory.md") if ld else None] if p and p.is_file()]
    return merged, sources, layers, memory, ld


def cmd_show(args):
    merged, sources, layers, memory, ld = resolve()
    print(json.dumps({
        "settings": merged,
        "sources": sources,           # key -> file that set it (keys absent here are defaults)
        "layers": layers,             # files actually read, in precedence order (later wins)
        "memory": memory,             # memory.md files that exist
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
    mp = target / "memory.md"
    if not mp.exists():
        mp.write_text(MEMORY_TEMPLATE, encoding="utf-8")
        created.append(str(mp))
    print(json.dumps({"dir": str(target), "created": created, "already_present": [str(p) for p in (sp, ex, mp) if str(p) not in created]}, ensure_ascii=False, indent=2))


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
    _, _, _, memory, _ = resolve()
    if not memory:
        print("(no memory.md found; run `settings.py init` to create one)")
        return
    for m in memory:
        print(f"===== {m}")
        print(Path(m).read_text(encoding="utf-8"))


def cmd_path(args):
    _, _, _, _, ld = resolve()
    print(json.dumps({"user_dir": str(user_dir()), "local_dir": str(ld) if ld else None}, ensure_ascii=False))


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    sub.add_parser("show").set_defaults(fn=cmd_show)
    p = sub.add_parser("init"); p.add_argument("--local", action="store_true"); p.set_defaults(fn=cmd_init)
    p = sub.add_parser("set"); p.add_argument("key"); p.add_argument("value"); p.add_argument("--local", action="store_true"); p.set_defaults(fn=cmd_set)
    sub.add_parser("memory").set_defaults(fn=cmd_memory)
    sub.add_parser("path").set_defaults(fn=cmd_path)
    args = ap.parse_args(argv)
    args.fn(args)


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    main()
