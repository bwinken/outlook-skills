# outlook-status reference

## 1. Script output (JSON)

Top level:

| Field | Type | Notes |
|---|---|---|
| `GeneratedAt` | ISO datetime | local time |
| `OfficeVersion` | `"16.0"` / `"15.0"` / `"14.0"` / null | registry hive found; null = Classic Outlook never installed or never run |
| `OutlookExeVersion` | string | e.g. `16.0.17928.20114`; missing when `-SkipCom` or COM failed |
| `NewOutlookEnabled` | bool / null | true means the "New Outlook" toggle is on; COM will not work |
| `DefaultProfile` | string | |
| `Profiles[]` | string | profile names |
| `Accounts[]` | object | `DisplayName`, `SmtpAddress`, `UserName`, `Type` (Exchange / IMAP / POP3 / HTTP / EAS / Other) |
| `Stores[]` | object | see below |
| `DataFilesOnDisk[]` | object | `Path`, `Type` (PST / OST), `SizeMB`, `Modified` |
| `Warnings[]` | string | anything that failed; always show these to the user |

`Stores[]` item:

| Field | Notes |
|---|---|
| `DisplayName` | as shown in Outlook's folder pane |
| `FilePath` | .ost / .pst path; empty for online-only stores |
| `SizeMB` | null when the file is not readable |
| `StoreType` | PrimaryExchangeMailbox / AdditionalExchangeMailbox / DelegateExchangeMailbox / ExchangePublicFolder / NotExchange |
| `IsDataFileStore` | true for .pst |
| `IsCachedExchange` | true for .ost cached mode |
| `Folders[]` | top-level folders: `Name`, `Items`, `Unread` (null for non-mail folders), `SubFolders` |

## 2. Presentation template

Match the user's language. Labels below are Traditional Chinese; translate them if the user writes in English.

```
## Outlook 狀態

**環境**：Classic Outlook {OutlookExeVersion}，Profile「{DefaultProfile}」{；New Outlook 已啟用，COM 功能無法使用 — only if NewOutlookEnabled}

**帳號**
| 帳號 | 類型 | 地址 |
|---|---|---|
| {DisplayName} | {Type} | {SmtpAddress} |

**資料檔**
| 名稱 | 類型 | 大小 | 路徑 | 收件匣未讀 |
|---|---|---|---|---|
| {DisplayName} | {OST 快取 / PST 資料檔 / 線上} | {SizeMB} MB | {FilePath} | {Folders[Inbox].Unread} |

**未掛載的資料檔**（only if DataFilesOnDisk has paths not present in Stores）
- {Path}（{SizeMB} MB，最後修改 {Modified}）

**警告**（only if Warnings non-empty）
- {each warning}

**建議**（only when something is worth saying）
- {e.g. OST 已達 38 GB，接近 50 GB 上限，建議在 Outlook 內封存舊郵件}
- {e.g. 有 2 個 .pst 在磁碟上但沒有掛載，若已不用可考慮備份後移除}
```

Rules:
- Size: one decimal in MB below 1024 MB, otherwise convert to GB with one decimal.
- Never propose that the plugin itself compacts, deletes or archives. Suggestions are for the user to do in Outlook.
- If `Stores` is empty and `Warnings` mentions COM, show the file-system findings and say the user can rerun after opening Classic Outlook.
