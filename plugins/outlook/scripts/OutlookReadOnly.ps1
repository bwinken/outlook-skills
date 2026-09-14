<#
.SYNOPSIS
    Shared READ-ONLY helpers for the Outlook skills plugin.

.DESCRIPTION
    Every function in this module only reads from Outlook through COM automation.
    The module deliberately exposes NO wrapper for Save, Send, Delete, Move, Copy,
    MarkAsRead/UnRead, Display, or any property setter. Scripts that import this
    module must follow the same rule: never call a mutating COM method.

    Requires Classic Outlook (Outlook 2013/2016/2019/2021/Microsoft 365) on Windows.
    "New Outlook" (olk.exe) has no COM object model and is not supported.

    Recommended host: Windows PowerShell 5.1 (powershell.exe). PowerShell 7 also works.

    This file is a plain script library, not a module, on purpose: the skill scripts load it
    with  . ([scriptblock]::Create((Get-Content -Raw <path>)))  which is not subject to the
    PowerShell execution policy. Import-Module on a .psm1 would be blocked on machines where
    Group Policy enforces Restricted/AllSigned.
#>

# Outlook default folder constants (OlDefaultFolders)
$script:OlFolder = @{
    DeletedItems = 3
    Outbox       = 4
    SentMail     = 5
    Inbox        = 6
    Calendar     = 9
    Contacts     = 10
    Journal      = 11
    Notes        = 12
    Tasks        = 13
    Drafts       = 16
    Junk         = 23
}

# MAPI property tags used through PropertyAccessor (read-only lookups)
$script:PR_SENDER_SMTP_ADDRESS = 'http://schemas.microsoft.com/mapi/proptag/0x5D01001F'
$script:PR_SMTP_ADDRESS        = 'http://schemas.microsoft.com/mapi/proptag/0x39FE001F'

$script:Outlook   = $null
$script:Namespace = $null

function Initialize-OutlookConsole {
    # Make sure non-ASCII subjects/bodies survive the trip to stdout.
    try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }
    try { $global:OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }
}

function Connect-Outlook {
    <#
    .SYNOPSIS
        Returns the MAPI namespace. Attaches to a running Outlook or starts one in the background.
    #>
    if ($script:Namespace) { return $script:Namespace }
    try {
        # Outlook is single-instance: New-Object attaches to the running one if present.
        $script:Outlook = New-Object -ComObject 'Outlook.Application'
    } catch {
        throw "Cannot start Outlook COM automation. Classic Outlook must be installed. New Outlook (olk.exe) is not supported. Error: $($_.Exception.Message)"
    }
    $script:Namespace = $script:Outlook.GetNamespace('MAPI')
    return $script:Namespace
}

function Get-OutlookApplication {
    if (-not $script:Outlook) { [void](Connect-Outlook) }
    return $script:Outlook
}

function Get-OlDefaultFolder {
    param([Parameter(Mandatory)][string]$Name)
    $ns = Connect-Outlook
    if (-not $script:OlFolder.ContainsKey($Name)) { throw "Unknown default folder '$Name'." }
    return $ns.GetDefaultFolder($script:OlFolder[$Name])
}

function Get-OlStores {
    $ns = Connect-Outlook
    $result = @()
    foreach ($s in $ns.Stores) { $result += $s }
    return $result
}

function Get-OlFolder {
    <#
    .SYNOPSIS
        Resolves a folder from a path like "Inbox", "Inbox/Projects/Alpha", or "\\Store Name\Inbox\Sub".
        Empty path returns the default Inbox.
    #>
    param(
        [string]$Path = '',
        [string]$Store = ''
    )
    $ns = Connect-Outlook
    if ([string]::IsNullOrWhiteSpace($Path)) {
        if ($Store) {
            return (Get-OlStores | Where-Object { $_.DisplayName -eq $Store } | Select-Object -First 1).GetDefaultFolder($script:OlFolder.Inbox)
        }
        return $ns.GetDefaultFolder($script:OlFolder.Inbox)
    }

    $parts = @($Path -split '[\\/]' | Where-Object { $_ -ne '' })

    # Absolute path "\\Store\Folder\..." names the store in the first segment.
    $root = $null
    if ($Path.StartsWith('\\') -or $Store) {
        if ($Store) {
            $storeName = $Store
        } else {
            $storeName = $parts[0]
            $parts = @($parts | Select-Object -Skip 1)
        }
        $st = Get-OlStores | Where-Object { $_.DisplayName -eq $storeName } | Select-Object -First 1
        if (-not $st) { throw "Store '$storeName' not found." }
        $root = $st.GetRootFolder()
    } else {
        # Relative paths start from the default store. Well-known names map to default folders.
        $first = $parts[0]
        $known = $script:OlFolder.Keys | Where-Object { $_ -ieq $first -or ($first -ieq 'Sent Items' -and $_ -eq 'SentMail') -or ($first -ieq 'Deleted Items' -and $_ -eq 'DeletedItems') -or ($first -ieq 'Junk Email' -and $_ -eq 'Junk') } | Select-Object -First 1
        if ($known) {
            $root = $ns.GetDefaultFolder($script:OlFolder[$known])
            $parts = @($parts | Select-Object -Skip 1)
        } else {
            $root = $ns.GetDefaultFolder($script:OlFolder.Inbox).Parent
        }
    }

    $current = $root
    foreach ($p in $parts) {
        $next = $null
        foreach ($f in $current.Folders) {
            if ($f.Name -ieq $p) { $next = $f; break }
        }
        if (-not $next) { throw "Folder '$p' not found under '$($current.FolderPath)'." }
        $current = $next
    }
    return $current
}

function Get-OlMailFoldersRecursive {
    <#
    .SYNOPSIS
        Enumerates all mail folders (DefaultItemType = olMailItem) under a folder, including itself.
    #>
    param([Parameter(Mandatory)]$Folder)
    $out = @()
    if ($Folder.DefaultItemType -eq 0) { $out += $Folder }
    foreach ($f in $Folder.Folders) { $out += Get-OlMailFoldersRecursive -Folder $f }
    return $out
}

function Get-OlSenderSmtp {
    param([Parameter(Mandatory)]$Mail)
    try {
        $v = $Mail.PropertyAccessor.GetProperty($script:PR_SENDER_SMTP_ADDRESS)
        if ($v) { return [string]$v }
    } catch { }
    try {
        if ($Mail.SenderEmailType -eq 'EX') {
            $u = $Mail.Sender.GetExchangeUser()
            if ($u -and $u.PrimarySmtpAddress) { return $u.PrimarySmtpAddress }
        }
    } catch { }
    try { return [string]$Mail.SenderEmailAddress } catch { return '' }
}

function Get-OlRecipientList {
    param([Parameter(Mandatory)]$Item, [int]$Type = 0)
    # Type: 1 = To, 2 = CC, 3 = BCC, 0 = all
    $list = @()
    try {
        foreach ($r in $Item.Recipients) {
            if ($Type -ne 0 -and $r.Type -ne $Type) { continue }
            $addr = ''
            try { $addr = [string]$r.PropertyAccessor.GetProperty($script:PR_SMTP_ADDRESS) } catch { }
            if (-not $addr) { try { $addr = [string]$r.Address } catch { } }
            $list += [pscustomobject]@{ Name = [string]$r.Name; Address = $addr }
        }
    } catch { }
    return $list
}

function Get-OlAttachmentList {
    param([Parameter(Mandatory)]$Item)
    $list = @()
    try {
        foreach ($a in $Item.Attachments) {
            $list += [pscustomobject]@{
                FileName = [string]$a.FileName
                Size     = [int64]$a.Size
                Type     = [int]$a.Type   # 1 = file, 5 = embedded, 6 = OLE
            }
        }
    } catch { }
    return $list
}

function ConvertTo-OlMailSummary {
    <#
    .SYNOPSIS
        Projects a MailItem into a plain PSCustomObject (safe for ConvertTo-Json).
    #>
    param(
        [Parameter(Mandatory)]$Mail,
        [switch]$IncludeBody,
        [int]$PreviewLength = 200
    )
    $body = ''
    try { $body = [string]$Mail.Body } catch { }
    $preview = if ($body.Length -gt $PreviewLength) { $body.Substring(0, $PreviewLength) + '...' } else { $body }

    $sentOn = $null
    try { $sentOn = (Get-Date $Mail.SentOn).ToString('s') } catch { }
    $flag = 0
    try { $flag = [int]$Mail.FlagStatus } catch { }
    $convId = ''
    try { $convId = [string]$Mail.ConversationID } catch { }
    $attachments = @(Get-OlAttachmentList -Item $Mail)

    $obj = [ordered]@{
        EntryID           = [string]$Mail.EntryID
        Folder            = [string]$Mail.Parent.FolderPath
        ReceivedTime      = (Get-Date $Mail.ReceivedTime).ToString('s')
        SentOn            = $sentOn
        From              = [string]$Mail.SenderName
        FromAddress       = Get-OlSenderSmtp -Mail $Mail
        To                = [string]$Mail.To
        CC                = [string]$Mail.CC
        Subject           = [string]$Mail.Subject
        Unread            = [bool]$Mail.UnRead
        HasAttachments    = ($attachments.Count -gt 0)
        Attachments       = $attachments
        Size              = [int64]$Mail.Size
        Importance        = [int]$Mail.Importance
        FlagStatus        = $flag
        Categories        = [string]$Mail.Categories
        ConversationID    = $convId
        ConversationTopic = [string]$Mail.ConversationTopic
        BodyPreview       = ($preview -replace '\s+', ' ').Trim()
    }
    if ($IncludeBody) { $obj['Body'] = $body }
    return [pscustomobject]$obj
}

function ConvertTo-OlAppointmentSummary {
    param([Parameter(Mandatory)]$Appt)
    $organizer = ''
    try { $organizer = [string]$Appt.Organizer } catch { }

    $busy = switch ([int]$Appt.BusyStatus) { 0 {'Free'} 1 {'Tentative'} 2 {'Busy'} 3 {'OutOfOffice'} 4 {'WorkingElsewhere'} default {'Unknown'} }
    $meeting = switch ([int]$Appt.MeetingStatus) { 0 {'NonMeeting'} 1 {'Meeting'} 3 {'Received'} 5 {'Canceled'} 7 {'ReceivedAndCanceled'} default {'Unknown'} }
    $response = switch ([int]$Appt.ResponseStatus) { 0 {'None'} 1 {'Organized'} 2 {'Tentative'} 3 {'Accepted'} 4 {'Declined'} 5 {'NotResponded'} default {'Unknown'} }

    $body = ''
    try { $body = (([string]$Appt.Body) -replace '\s+', ' ').Trim() } catch { }
    if ($body.Length -gt 300) { $body = $body.Substring(0, 300) + '...' }

    return [pscustomobject][ordered]@{
        EntryID           = [string]$Appt.EntryID
        Subject           = [string]$Appt.Subject
        Start             = (Get-Date $Appt.Start).ToString('s')
        End               = (Get-Date $Appt.End).ToString('s')
        DurationMinutes   = [int]$Appt.Duration
        AllDayEvent       = [bool]$Appt.AllDayEvent
        Location          = [string]$Appt.Location
        Organizer         = $organizer
        RequiredAttendees = [string]$Appt.RequiredAttendees
        OptionalAttendees = [string]$Appt.OptionalAttendees
        BusyStatus        = $busy
        MeetingStatus     = $meeting
        ResponseStatus    = $response
        IsRecurring       = [bool]$Appt.IsRecurring
        Categories        = [string]$Appt.Categories
        BodyPreview       = $body
    }
}

function Write-OlJson {
    <#
    .SYNOPSIS
        Emits an object as UTF-8 JSON to stdout, or to -OutFile if given.
    #>
    param(
        [Parameter(Mandatory)]$InputObject,
        [string]$OutFile = '',
        [int]$Depth = 8
    )
    $json = $InputObject | ConvertTo-Json -Depth $Depth
    if ($OutFile) {
        [System.IO.File]::WriteAllText($OutFile, $json, (New-Object System.Text.UTF8Encoding($false)))
        Write-Output "Written to $OutFile"
    } else {
        Write-Output $json
    }
}

function ConvertTo-DaslLiteral {
    # Escapes a value for use inside a single-quoted DASL string literal.
    param([string]$Value)
    return ($Value -replace "'", "''")
}
