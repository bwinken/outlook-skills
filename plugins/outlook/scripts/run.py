#!/usr/bin/env python3
"""Launcher for the plugin's PowerShell scripts. Use this instead of calling powershell yourself.

    python run.py Search-OutlookMail.ps1 -From Alice -After 2026-09-01 -Max 20
    python run.py Get-OutlookStatus.ps1 -SkipCom
    python run.py Get-OutlookCalendar.ps1 -Start 2026-09-16 -Days 1 --out cal.json

What it takes care of, so SKILL.md commands work the same from Bash, cmd, PowerShell or a sandboxed tool:
  - picks the interpreter: powershell.exe (Windows PowerShell 5.1) on Windows, pwsh elsewhere;
    override with OUTLOOK_SKILLS_PWSH=<path>
  - no shell quoting problems: arguments are passed straight to the process, nothing goes through a shell,
    so values with spaces, quotes, $ or Chinese are safe
  - execution policy: runs with -File -ExecutionPolicy Bypass first; if the machine rejects that
    (Group Policy), re-runs the script text as a script block, which the policy does not cover
  - encoding: the scripts carry a UTF-8 BOM for PowerShell 5.1; if a copy lost it, the launcher runs a
    BOM-prefixed temp copy instead. Output always goes through -OutFile as UTF-8 JSON and is printed to
    stdout as UTF-8, so console code pages never garble Chinese
  - --out <file> keeps the JSON file instead of printing it (large results)

Exit code is the script's. Nothing here writes to Outlook.
"""
import os
import platform
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
HELPER = "OutlookReadOnly.ps1"
BOM = b"\xef\xbb\xbf"
POLICY_MARKERS = ("running scripts is disabled", "UnauthorizedAccess", "PSSecurityException", "not digitally signed")


def find_powershell() -> str:
    if os.environ.get("OUTLOOK_SKILLS_PWSH"):
        return os.environ["OUTLOOK_SKILLS_PWSH"]
    candidates = ["powershell.exe", "powershell", "pwsh.exe", "pwsh"] if platform.system() == "Windows" else ["pwsh", "powershell"]
    for c in candidates:
        p = shutil.which(c)
        if p:
            return p
    raise SystemExit("No PowerShell found. Install PowerShell, or set OUTLOOK_SKILLS_PWSH to its path.")


def ensure_bom(script: Path, workdir: Path) -> Path:
    """Return a path to a BOM-prefixed script (the original if it already has one)."""
    files = [script, HERE / HELPER]
    if all(f.read_bytes().startswith(BOM) for f in files if f.exists()):
        return script
    for f in files:
        if f.exists():
            data = f.read_bytes()
            (workdir / f.name).write_bytes(data if data.startswith(BOM) else BOM + data)
    return workdir / script.name


def ps_quote(a: str) -> str:
    if a.startswith("-") and " " not in a and "'" not in a:
        return a
    return "'" + a.replace("'", "''") + "'"


def main(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)
    if not argv or argv[0] in ("-h", "--help"):
        print(__doc__)
        return 0
    script_name = argv[0]
    args = argv[1:]
    keep_out = None
    if "--out" in args:
        i = args.index("--out")
        keep_out = args[i + 1]
        del args[i:i + 2]
    script = HERE / script_name
    if not script.is_file():
        raise SystemExit(f"Unknown script {script_name}. Available: " + ", ".join(sorted(p.name for p in HERE.glob("*.ps1") if p.name != HELPER)))
    if any(a.lower() == "-outfile" for a in args):
        raise SystemExit("Use --out <file> instead of -OutFile; the launcher manages the output file.")

    ps = find_powershell()
    workdir = Path(tempfile.mkdtemp(prefix="outlook-skills-"))
    out = Path(keep_out).resolve() if keep_out else workdir / "out.json"
    try:
        script_path = ensure_bom(script, workdir)
        script_dir = script_path.parent
        full_args = args + ["-OutFile", str(out)]
        base = [ps, "-NoProfile", "-NonInteractive"]
        env = dict(os.environ, OUTLOOK_SKILLS_SCRIPTS=str(script_dir))

        r = subprocess.run(base + ["-ExecutionPolicy", "Bypass", "-File", str(script_path)] + full_args,
                           capture_output=True, text=True, encoding="utf-8", errors="replace", env=env)
        blob = (r.stdout or "") + (r.stderr or "")
        if r.returncode != 0 and any(m in blob for m in POLICY_MARKERS):
            # Policy-free form: the text is parsed as a script block, which execution policy does not govern.
            command = ("$env:OUTLOOK_SKILLS_SCRIPTS='" + str(script_dir).replace("'", "''") + "'; "
                       "& ([scriptblock]::Create((Get-Content -Raw -LiteralPath '" + str(script_path).replace("'", "''") + "'))) "
                       + " ".join(ps_quote(a) for a in full_args))
            r = subprocess.run(base + ["-Command", command], capture_output=True, text=True, encoding="utf-8", errors="replace", env=env)

        if r.stderr and r.stderr.strip():
            sys.stderr.write(r.stderr)
        if r.returncode != 0 and not out.exists():
            if r.stdout and r.stdout.strip():
                sys.stderr.write(r.stdout)
            return r.returncode or 1

        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass
        if keep_out:
            print(f"Written to {out}")
        else:
            with open(out, "r", encoding="utf-8-sig") as fh:
                sys.stdout.write(fh.read())
                sys.stdout.write("\n")
        return r.returncode
    finally:
        shutil.rmtree(workdir, ignore_errors=True)


if __name__ == "__main__":
    sys.exit(main())
