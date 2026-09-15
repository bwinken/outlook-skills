<#
.SYNOPSIS
    READ-ONLY. Returns every message in a conversation (thread), oldest first, with full bodies.

.DESCRIPTION
    Identify the thread by one of:
      -EntryID         an EntryID from Search-OutlookMail output (most precise)
      -ConversationID  a ConversationID from Search-OutlookMail output
      -Subject         substring; the newest matching mail in the folder anchors the thread

.EXAMPLE
    # Normal:
    powershell -NoProfile -ExecutionPolicy Bypass -File <this script> [args]
    # Execution policy enforced by Group Policy:
    powershell -NoProfile -Command "$env:OUTLOOK_SKILLS_SCRIPTS='<scripts dir>'; & ([scriptblock]::Create((Get-Content -Raw -LiteralPath '<scripts dir>\<this script>'))) [args]"

    Get-OutlookThread.ps1 -Subject "Q3 budget"
    Get-OutlookThread.ps1 -EntryID 00000000ABCD... -OutFile thread.json
#>
[CmdletBinding()]
param(
    [string]$EntryID = '',
    [string]$ConversationID = '',
    [string]$Subject = '',
    [string]$Folder = '',
    [string]$Store = '',
    [int]$MaxBodyChars = 20000,
    [string]$OutFile = ''
)

# ---- Load the shared read-only helpers without going through the execution policy.
#      $PSScriptRoot is set when run with -File; when run via -Command/[scriptblock]::Create it is
#      empty, so the caller sets OUTLOOK_SKILLS_SCRIPTS to this folder instead.
$scriptsDir = if ($PSScriptRoot) { $PSScriptRoot } elseif ($env:OUTLOOK_SKILLS_SCRIPTS) { $env:OUTLOOK_SKILLS_SCRIPTS } else {
    throw 'Cannot locate OutlookReadOnly.ps1. Run with -File, or set $env:OUTLOOK_SKILLS_SCRIPTS to the plugin scripts folder.'
}
. ([scriptblock]::Create((Get-Content -Raw -LiteralPath (Join-Path $scriptsDir 'OutlookReadOnly.ps1'))))
Initialize-OutlookConsole

$ns = Connect-Outlook
$anchor = $null

if ($EntryID) {
    $anchor = $ns.GetItemFromID($EntryID)
} elseif ($Subject -or $ConversationID) {
    $folder = Get-OlFolder -Path $Folder -Store $Store
    $items = $folder.Items
    if ($Subject) {
        $items = $items.Restrict("@SQL=""urn:schemas:httpmail:subject"" LIKE '%$(ConvertTo-DaslLiteral $Subject)%'")
    }
    $items.Sort('[ReceivedTime]', $true)
    $it = $items.GetFirst()
    while ($it -ne $null) {
        if ([int]$it.Class -eq 43) {
            if (-not $ConversationID -or [string]$it.ConversationID -eq $ConversationID) { $anchor = $it; break }
        }
        $it = $items.GetNext()
    }
} else {
    throw 'Provide -EntryID, -ConversationID or -Subject.'
}
if (-not $anchor) { throw 'No matching message found.' }

# ---- Collect conversation members
$messages = @()
$method = ''
try {
    $conv = $anchor.GetConversation()
    if ($conv -ne $null) {
        $method = 'GetConversation'
        $table = $conv.GetTable()
        while (-not $table.EndOfTable) {
            $row = $table.GetNextRow()
            try {
                $m = $ns.GetItemFromID([string]$row.Item('EntryID'))
                if ([int]$m.Class -eq 43) { $messages += $m }
            } catch { }
        }
    }
} catch { }

if ($messages.Count -eq 0) {
    # Fallback: same ConversationTopic across all mail folders of the store (works for POP/PST stores).
    $method = 'ConversationTopic'
    $topic = [string]$anchor.ConversationTopic
    $storeRoot = $anchor.Parent.Store.GetRootFolder()
    foreach ($f in Get-OlMailFoldersRecursive -Folder $storeRoot) {
        $r = $f.Items.Restrict("@SQL=""urn:schemas:httpmail:thread-topic"" = '$(ConvertTo-DaslLiteral $topic)'")
        $m = $r.GetFirst()
        while ($m -ne $null) {
            if ([int]$m.Class -eq 43) { $messages += $m }
            $m = $r.GetNext()
        }
    }
}

$summaries = @()
$seen = @{}
foreach ($m in $messages) {
    $id = [string]$m.EntryID
    if ($seen.ContainsKey($id)) { continue }
    $seen[$id] = $true
    $s = ConvertTo-OlMailSummary -Mail $m -IncludeBody
    if ($s.Body.Length -gt $MaxBodyChars) { $s.Body = $s.Body.Substring(0, $MaxBodyChars) + "`n[... truncated ...]" }
    $s | Add-Member -NotePropertyName ToRecipients -NotePropertyValue @(Get-OlRecipientList -Item $m -Type 1)
    $s | Add-Member -NotePropertyName CcRecipients -NotePropertyValue @(Get-OlRecipientList -Item $m -Type 2)
    $summaries += $s
}
$summaries = @($summaries | Sort-Object -Property ReceivedTime)

$out = [pscustomobject][ordered]@{
    Anchor         = [string]$anchor.EntryID
    Topic          = [string]$anchor.ConversationTopic
    ConversationID = $(try { [string]$anchor.ConversationID } catch { '' })
    Method         = $method
    Count          = $summaries.Count
    Participants   = @($summaries | ForEach-Object { $_.FromAddress } | Where-Object { $_ } | Sort-Object -Unique)
    Messages       = $summaries
}
Write-OlJson -InputObject $out -OutFile $OutFile
