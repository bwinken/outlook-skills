<#
.SYNOPSIS
    READ-ONLY. Aggregated overview of the mailbox for building the plugin's memory:
    top correspondents, folder tree with counts, frequent conversation topics,
    likely newsletters, and recurring meetings. No message bodies are returned.

.EXAMPLE
    # Normal:
    powershell -NoProfile -ExecutionPolicy Bypass -File Get-OutlookOverview.ps1 -Days 180 -OutFile overview.json
    # Execution policy enforced by Group Policy:
    powershell -NoProfile -Command "$env:OUTLOOK_SKILLS_SCRIPTS='<scripts dir>'; & ([scriptblock]::Create((Get-Content -Raw -LiteralPath '<scripts dir>\Get-OutlookOverview.ps1'))) -Days 180"
#>
[CmdletBinding()]
param(
    [int]$Days = 180,            # look-back window for mail
    [int]$MaxItems = 3000,       # cap on mails scanned per direction (received / sent), newest first
    [int]$Top = 30,              # rows per ranking
    [string]$Store = '',
    [string]$OutFile = ''
)

$scriptsDir = if ($PSScriptRoot) { $PSScriptRoot } elseif ($env:OUTLOOK_SKILLS_SCRIPTS) { $env:OUTLOOK_SKILLS_SCRIPTS } else {
    throw 'Cannot locate OutlookReadOnly.ps1. Run with -File, or set $env:OUTLOOK_SKILLS_SCRIPTS to the plugin scripts folder.'
}
. ([scriptblock]::Create((Get-Content -Raw -LiteralPath (Join-Path $scriptsDir 'OutlookReadOnly.ps1'))))
Initialize-OutlookConsole

$ns = Connect-Outlook
$since = (Get-Date).Date.AddDays(-$Days)
$me = ''
try { $me = [string]$ns.CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress } catch { }
if (-not $me) { try { $me = [string]$ns.CurrentUser.Address } catch { } }

function Add-Count([hashtable]$table, [string]$key, [hashtable]$extra) {
    if (-not $key) { return }
    if (-not $table.ContainsKey($key)) { $table[$key] = @{ Key = $key; Count = 0 } + $extra }
    $table[$key].Count++
}

# ---- folders ------------------------------------------------------------------------------
$root = if ($Store) { (Get-OlStores | Where-Object { $_.DisplayName -eq $Store } | Select-Object -First 1).GetRootFolder() } else { $ns.GetDefaultFolder(6).Parent }
$folders = @()
foreach ($f in Get-OlMailFoldersRecursive -Folder $root) {
    $unread = $null; try { $unread = [int]$f.UnReadItemCount } catch { }
    $newest = $null
    try { $it = $f.Items; $it.Sort('[ReceivedTime]', $true); $first = $it.GetFirst(); if ($first) { $newest = (Get-Date $first.ReceivedTime).ToString('s') } } catch { }
    $folders += [pscustomobject]@{ Path = [string]$f.FolderPath; Items = [int]$f.Items.Count; Unread = $unread; Newest = $newest }
}

# ---- received mail: senders, topics, newsletters ---------------------------------------------
$senders = @{}; $topics = @{}; $scannedIn = 0
$inbox = if ($Store) { $root.Folders | Where-Object { $_.Name -in 'Inbox','收件匣' } | Select-Object -First 1 } else { $ns.GetDefaultFolder(6) }
foreach ($f in Get-OlMailFoldersRecursive -Folder $inbox) {
    $items = $f.Items; $items.Sort('[ReceivedTime]', $true)
    $m = $items.GetFirst()
    while ($m -ne $null -and $scannedIn -lt $MaxItems) {
        if ([int]$m.Class -eq 43) {
            $rt = Get-Date $m.ReceivedTime
            if ($rt -lt $since) { break }
            $scannedIn++
            $addr = (Get-OlSenderSmtp -Mail $m).ToLower()
            Add-Count $senders $addr @{ Name = [string]$m.SenderName; Last = $rt.ToString('s'); Unsub = $false; Folder = [string]$f.FolderPath }
            if ($rt.ToString('s') -gt $senders[$addr].Last) { $senders[$addr].Last = $rt.ToString('s') }
            try { if ($m.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/string/{00020386-0000-0000-C000-000000000046}/List-Unsubscribe')) { $senders[$addr].Unsub = $true } } catch { }
            $topic = [string]$m.ConversationTopic
            if ($topic) { Add-Count $topics $topic @{ Last = $rt.ToString('s'); Sample = $addr } }
        }
        $m = $items.GetNext()
    }
}

# ---- sent mail: recipients ---------------------------------------------------------------------
$recips = @{}; $scannedOut = 0
$sent = if ($Store) { $root.Folders | Where-Object { $_.Name -in 'Sent Items','寄件備份' } | Select-Object -First 1 } else { $ns.GetDefaultFolder(5) }
if ($sent) {
    $items = $sent.Items; $items.Sort('[SentOn]', $true)
    $m = $items.GetFirst()
    while ($m -ne $null -and $scannedOut -lt $MaxItems) {
        if ([int]$m.Class -eq 43) {
            $st = $null; try { $st = Get-Date $m.SentOn } catch { }
            if ($st -and $st -lt $since) { break }
            $scannedOut++
            foreach ($r in Get-OlRecipientList -Item $m -Type 1) {
                $a = ([string]$r.Address).ToLower()
                Add-Count $recips $a @{ Name = [string]$r.Name }
            }
        }
        $m = $items.GetNext()
    }
}

# ---- recurring meetings --------------------------------------------------------------------------
$recurring = @()
try {
    $cal = if ($Store) { $root.Folders | Where-Object { $_.Name -in 'Calendar','行事曆' } | Select-Object -First 1 } else { $ns.GetDefaultFolder(9) }
    $items = $cal.Items; $items.IncludeRecurrences = $false
    $it = $items.GetFirst()
    while ($it -ne $null) {
        if ([int]$it.Class -eq 26 -and $it.IsRecurring) {
            $rp = $it.GetRecurrencePattern()
            $type = switch ([int]$rp.RecurrenceType) { 0 {'Daily'} 1 {'Weekly'} 2 {'Monthly'} 3 {'MonthNth'} 5 {'Yearly'} 6 {'YearNth'} default {'Other'} }
            $recurring += [pscustomobject]@{
                Subject = [string]$it.Subject; Organizer = [string]$it.Organizer; Pattern = $type
                Interval = [int]$rp.Interval; DayOfWeekMask = [int]$rp.DayOfWeekMask
                StartTime = (Get-Date $rp.StartTime).ToString('HH:mm'); DurationMinutes = [int]$rp.Duration
                PatternEnd = if ($rp.NoEndDate) { $null } else { (Get-Date $rp.PatternEndDate).ToString('yyyy-MM-dd') }
                Attendees = [string]$it.RequiredAttendees
            }
        }
        $it = $items.GetNext()
    }
} catch { }

function Top-Rows($table, $n) { @($table.Values | Sort-Object -Property Count -Descending | Select-Object -First $n | ForEach-Object { [pscustomobject]$_ }) }

$out = [pscustomobject][ordered]@{
    GeneratedAt   = (Get-Date).ToString('s')
    Me            = $me
    Window        = [pscustomobject]@{ Since = $since.ToString('s'); Days = $Days; ScannedReceived = $scannedIn; ScannedSent = $scannedOut; MaxItems = $MaxItems }
    Folders       = $folders
    TopSenders    = Top-Rows $senders $Top
    TopRecipients = Top-Rows $recips $Top
    TopTopics     = Top-Rows $topics $Top
    Newsletters   = @($senders.Values | Where-Object { $_.Unsub } | Sort-Object -Property Count -Descending | Select-Object -First $Top | ForEach-Object { [pscustomobject]$_ })
    RecurringMeetings = $recurring
}
Write-OlJson -InputObject $out -OutFile $OutFile
