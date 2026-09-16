#!/usr/bin/env python3
"""Keep plugins/outlook-mcp/scripts/ a verbatim copy of plugins/outlook/scripts/.

plugins/outlook/scripts is the source of truth (the skills plugin). The MCP plugin needs the same
files inside its own directory because Claude Code copies only a plugin's own folder when it
installs it, and git symlinks are not reliable on Windows. So the copy is committed, and CI fails
when the two drift.

    python tools/sync_scripts.py            # copy source -> mcp (adds, updates, removes stale files)
    python tools/sync_scripts.py --check    # exit 1 and list the differences, change nothing
"""
import argparse
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "plugins" / "outlook" / "scripts"
DST = ROOT / "plugins" / "outlook-mcp" / "scripts"


def _files(d: Path):
    return {p.name: p for p in d.glob("*.py")} if d.is_dir() else {}


def differences():
    src, dst = _files(SRC), _files(DST)
    out = []
    for name in sorted(set(src) | set(dst)):
        if name not in dst:
            out.append(("missing in mcp", name))
        elif name not in src:
            out.append(("stale in mcp", name))
        elif src[name].read_bytes() != dst[name].read_bytes():
            out.append(("differs", name))
    return out


def sync():
    DST.mkdir(parents=True, exist_ok=True)
    done = []
    for kind, name in differences():
        if kind == "stale in mcp":
            (DST / name).unlink()
        else:
            shutil.copyfile(SRC / name, DST / name)
        done.append((kind, name))
    for p in DST.glob("__pycache__"):
        shutil.rmtree(p, ignore_errors=True)
    return done


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--check", action="store_true", help="report differences and exit 1 if any; change nothing")
    a = ap.parse_args(argv)
    if a.check:
        diffs = differences()
        for kind, name in diffs:
            print(f"{kind}: {name}")
        if diffs:
            print(f"\n{DST.relative_to(ROOT)} is out of sync with {SRC.relative_to(ROOT)}; run: python tools/sync_scripts.py")
            sys.exit(1)
        print("mcp scripts in sync")
        return
    done = sync()
    print("\n".join(f"{k}: {n}" for k, n in done) or "already in sync")


if __name__ == "__main__":
    main()
