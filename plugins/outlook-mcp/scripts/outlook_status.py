#!/usr/bin/env python3
"""READ-ONLY. Reports Outlook version, profiles, accounts, data files (.pst/.ost) and folder counts.

    python outlook_status.py
    python outlook_status.py -SkipCom -OutFile status.json
"""
import datetime as dt
import os
import platform
from pathlib import Path

import outlook_com as oc


def registry_info(report: dict):
    if platform.system() != "Windows":
        report["Warnings"].append("Not Windows: registry and COM information unavailable.")
        return
    import winreg
    for ver in ("16.0", "15.0", "14.0"):
        key = rf"Software\Microsoft\Office\{ver}\Outlook"
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key) as k:
                report["OfficeVersion"] = ver
                try:
                    report["DefaultProfile"] = winreg.QueryValueEx(k, "DefaultProfile")[0]
                except OSError:
                    pass
        except OSError:
            continue
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key + r"\Profiles") as pk:
                i, names = 0, []
                while True:
                    try:
                        names.append(winreg.EnumKey(pk, i)); i += 1
                    except OSError:
                        break
                report["Profiles"] = names
        except OSError:
            pass
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key + r"\Preferences") as prk:
                report["NewOutlookEnabled"] = int(winreg.QueryValueEx(prk, "UseNewOutlook")[0]) == 1
        except OSError:
            pass
        break
    if not report["OfficeVersion"]:
        report["Warnings"].append("No Classic Outlook registry hive found under HKCU. Outlook may not be installed or never launched.")
    if report["NewOutlookEnabled"]:
        report["Warnings"].append("New Outlook toggle is ON. COM automation only works with Classic Outlook.")


def data_files_on_disk(report: dict):
    candidates = []
    if os.environ.get("LOCALAPPDATA"):
        candidates.append(Path(os.environ["LOCALAPPDATA"]) / "Microsoft" / "Outlook")
    docs = Path.home() / "Documents"
    candidates += [docs / "Outlook Files", docs / "Outlook 檔案"]
    for d in candidates:
        if not d.is_dir():
            continue
        for p in list(d.rglob("*.pst")) + list(d.rglob("*.ost")):
            st = p.stat()
            report["DataFilesOnDisk"].append({
                "Path": str(p), "Type": p.suffix[1:].upper(), "SizeMB": round(st.st_size / 1048576, 1),
                "Modified": dt.datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%dT%H:%M:%S"),
            })


_ACCOUNT_TYPES = {0: "Exchange", 1: "IMAP", 2: "POP3", 3: "HTTP", 4: "EAS", 5: "Other"}
_STORE_TYPES = {0: "PrimaryExchangeMailbox", 1: "ExchangePublicFolder", 2: "DelegateExchangeMailbox", 3: "NotExchange", 4: "AdditionalExchangeMailbox"}


def com_info(report: dict):
    ns = oc.connect()
    report["OutlookExeVersion"] = str(oc.application().Version)
    for a in ns.Accounts:
        report["Accounts"].append({
            "DisplayName": str(a.DisplayName), "SmtpAddress": str(a.SmtpAddress), "UserName": str(a.UserName),
            "Type": _ACCOUNT_TYPES.get(int(a.AccountType), "Unknown"),
        })
    for s in oc.get_stores(ns):
        path = oc._safe(lambda: str(s.FilePath), "") or ""
        size = round(os.path.getsize(path) / 1048576, 1) if path and os.path.exists(path) else None
        folders = []
        try:
            for f in s.GetRootFolder().Folders:
                folders.append({
                    "Name": str(f.Name), "Items": int(f.Items.Count),
                    "Unread": oc._safe(lambda: int(f.UnReadItemCount)), "SubFolders": int(f.Folders.Count),
                })
        except Exception as e:
            report["Warnings"].append(f"Could not enumerate folders of store '{s.DisplayName}': {e}")
        report["Stores"].append({
            "DisplayName": str(s.DisplayName), "FilePath": path, "SizeMB": size,
            "StoreType": _STORE_TYPES.get(int(s.ExchangeStoreType), "Unknown"),
            "IsDataFileStore": bool(s.IsDataFileStore), "IsCachedExchange": bool(oc._safe(lambda: s.IsCachedExchange, False)),
            "Folders": folders,
        })


def parser():
    ap = oc.ArgParser(description=__doc__)
    ap.flag("-SkipCom", help="only inspect registry and file system; do not talk to Outlook")
    oc.add_common_output(ap)
    return ap


def run(a, ns=None):
    report = {
        "GeneratedAt": dt.datetime.now().strftime("%Y-%m-%dT%H:%M:%S"),
        "Machine": os.environ.get("COMPUTERNAME") or platform.node(), "User": os.environ.get("USERNAME") or os.environ.get("USER"),
        "OfficeVersion": None, "OutlookExeVersion": None, "NewOutlookEnabled": None, "DefaultProfile": None, "Profiles": [],
        "Accounts": [], "Stores": [], "DataFilesOnDisk": [], "Warnings": [],
    }
    registry_info(report)
    data_files_on_disk(report)
    if not a.skipcom:
        try:
            com_info(report)
        except SystemExit as e:
            report["Warnings"].append(f"COM automation failed: {e}")
        except Exception as e:
            report["Warnings"].append(f"COM automation failed: {e}")
    return report


def main(argv=None):
    a = parser().parse_args(argv)
    oc.write_json(run(a), a.out_file)


if __name__ == "__main__":
    main()
