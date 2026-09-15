# outlook-attachments reference

## 1. Script output (JSON)

```
{
  "Query": { <search filters> + Name, Ext[], MinSizeKB, Sort, Top, SaveTo },
  "MailsScanned": <mails with attachments that matched the search>,
  "Count": <attachments after filtering>, "TotalSizeKB": <sum>,
  "Results": [ { FileName, Ext, SizeKB, ReceivedTime, From, FromAddress, Subject, Folder, EntryID, SavedTo? }, ... ]
}
```

`SavedTo` appears only when `-SaveTo` was used. `Results` is already sorted per `-Sort` and cut to `-Top`.

## 2. Presentation template

Match the user's language; labels below are Traditional Chinese.

```
在 {資料夾/store 描述} 找到 {Count} 個附件（共 {TotalSizeKB → MB}），顯示前 {shown} 個，依{大小|日期|檔名}排序：

| # | 檔案 | 大小 | 寄件者 | 日期 | 主旨 |
|---|---|---|---|---|---|
| 1 | Q3_report.pptx | 4.6 MB | David WY Chen | 09/11 | FYI: 季報 |
| 2 | 合約草稿_v3_legal.docx | 180 KB | Cassie Tsai | 09/12 | Re: 合約草稿 v3 - 法務意見 |

要存出哪幾個？告訴我資料夾即可。（只會複製檔案，Outlook 不會有任何變動）
```

After `-SaveTo`:

```
已複製 {n} 個檔案到 {folder}：
- Q3_report.pptx（4.6 MB）
- 合約草稿_v3_legal.docx（180 KB）
```

Rules:
- Size: KB below 1024, otherwise MB with one decimal.
- Show at most 20 rows; say how many more. Date `MM/dd`; subject trimmed to ~50 characters.
- Duplicated file names across mails are separate rows; when the user asks for "the" file, prefer the newest.
- Do not describe file contents you have not read; the plugin lists and copies, nothing more.
