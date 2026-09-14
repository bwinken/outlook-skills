#!/usr/bin/env python3
"""Install the Outlook skills into an Agent Skills host other than Claude Code (Roo Code, or any
tool that reads .agents/skills/).

Claude Code loads this folder as a plugin and expands ${CLAUDE_PLUGIN_ROOT} itself. Other hosts
do not know that variable, so this script copies each skills/<name>/ folder to the host's skills
directory and rewrites ${CLAUDE_PLUGIN_ROOT} to this folder's absolute path. Scripts stay here in
one place; only SKILL.md and reference.md are copied.

Usage:
    python install.py --roo              # ./.roo/skills/<name>/          (this project, Roo Code)
    python install.py --roo --global     # ~/.roo/skills/<name>/          (all projects, Roo Code)
    python install.py --agents           # ./.agents/skills/<name>/       (Agent Skills standard path)
    python install.py --dest <dir>       # any directory
    python install.py --roo --uninstall  # remove what a previous run created
    python install.py --roo --dry-run

Re-run after pulling updates; existing copies are replaced.
"""
import argparse
import shutil
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
SKILLS = HERE / "skills"
VAR = "${CLAUDE_PLUGIN_ROOT}"
MARKER = ".installed-by-outlook-skills"


def dest_dir(args) -> Path:
    if args.dest:
        return Path(args.dest).expanduser().resolve()
    base = Path.home() if args.global_ else Path.cwd()
    if args.roo:
        return base / ".roo" / "skills"
    if args.agents:
        return base / ".agents" / "skills"
    raise SystemExit("Choose --roo, --agents or --dest <dir>.")


def rewrite(text: str, root: Path) -> str:
    return text.replace(VAR, root.as_posix())


def install(dest: Path, dry: bool):
    root = HERE
    done = []
    for skill in sorted(p for p in SKILLS.iterdir() if p.is_dir() and (p / "SKILL.md").is_file()):
        target = dest / skill.name
        if dry:
            done.append(f"{skill.name} -> {target}")
            continue
        if target.exists():
            if not (target / MARKER).exists():
                raise SystemExit(f"{target} exists and was not created by this installer; remove it first.")
            shutil.rmtree(target)
        target.mkdir(parents=True)
        for f in skill.iterdir():
            if f.is_file() and f.suffix.lower() == ".md":
                (target / f.name).write_text(rewrite(f.read_text(encoding="utf-8"), root), encoding="utf-8")
            elif f.is_file():
                shutil.copy2(f, target / f.name)
        (target / MARKER).write_text(f"source: {skill}\n", encoding="utf-8")
        done.append(f"{skill.name} -> {target}")
    return done


def uninstall(dest: Path, dry: bool):
    done = []
    for skill in sorted(p for p in SKILLS.iterdir() if p.is_dir()):
        target = dest / skill.name
        if (target / MARKER).exists():
            if not dry:
                shutil.rmtree(target)
            done.append(str(target))
    return done


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--roo", action="store_true", help="Roo Code: .roo/skills/")
    ap.add_argument("--agents", action="store_true", help="Agent Skills standard: .agents/skills/")
    ap.add_argument("--dest", help="explicit skills directory")
    ap.add_argument("--global", dest="global_", action="store_true", help="use the home directory instead of the current directory")
    ap.add_argument("--uninstall", action="store_true")
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args(argv)
    dest = dest_dir(args)
    if args.uninstall:
        removed = uninstall(dest, args.dry_run)
        print(("Would remove:\n" if args.dry_run else "Removed:\n") + ("\n".join(removed) or "(nothing)"))
        return
    done = install(dest, args.dry_run)
    print(("Would install:\n" if args.dry_run else "Installed:\n") + "\n".join(done))
    print(f"\n{VAR} rewritten to {HERE.as_posix()}")
    print("Scripts run from that folder; keep this clone in place, or re-run install.py after moving it.")
    print("Settings and memory still live in ~/.outlook-skills/ and ./.outlook-skills/, shared with Claude Code.")


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    main()
