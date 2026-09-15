<#
.SYNOPSIS
    READ-ONLY. Searches mail items by sender, subject, body, date range, attachments and unread state.

.EXAMPLE
    # Normal:
    powershell -NoProfile -ExecutionPolicy Bypass -File <this script> [args]
    # Execution policy enforced by Group Policy:
    powershell -NoProfile -Command "$env:OUTLOOK_SKILLS_SCRIPTS='<scripts dir>'; & ([scriptblock]::Create((Get-Content -Raw -LiteralPath '<scripts dir>\<this script>'))) [args]"

    Search-OutlookMail.ps1 -From "alice" -After 2026-09-01 -Max 20
    Search-OutlookMail.ps1 -Subject "invoice" -HasAttachments -AllFolders
    Search-OutlookMail.ps1 -Folder "Inbox/Projects" -Unread -IncludeBody -OutFile hits.json
    Search-OutlookMail.ps1 -AnyOf "報價","quote","pricing" -After 2026-06-01 -AllFolders -Max 300 -PreviewLength 500 -OutFile candidates.json
#>
[CmdletBinding()]
param(
    [string]$From = '',            # matches sender name or address (substring)
    [string]$To = '',              # matches To/CC display string (substring)
    [string]$Subject = '',         # substring
    [string]$Body = '',            # substring in plain-text body
    [string]$Text = '',            # substring in subject OR body
    [string[]]$AnyOf = @(),        # several terms, subject OR body, joined with OR (keyword expansion)
    [datetime]$After,
    [datetime]$Before,
    [switch]$HasAttachments,
    [switch]$Unread,
    [string]$Folder = '',          # e.g. "Inbox", "Sent Items", "Inbox/Projects", "\\Store\Inbox"
    [string]$Store = '',
    [switch]$AllFolders,           # recurse through every mail folder of the store
    [int]$Max = 50,
    [switch]$IncludeBody,
    [int]$PreviewLength = 200,     # BodyPreview length; raise to ~500 when feeding rerank.py
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

# ---- Build DASL filter for the text/flag criteria (dates are applied in-process for exactness)
$clauses = @()
if ($From) {
    $v = ConvertTo-DaslLiteral $From
    $clauses += "(""urn:schemas:httpmail:fromname"" LIKE '%$v%' OR ""urn:schemas:httpmail:fromemail"" LIKE '%$v%')"
}
if ($To) {
    $v = ConvertTo-DaslLiteral $To
    $clauses += "(""urn:schemas:httpmail:displayto"" LIKE '%$v%' OR ""urn:schemas:httpmail:displaycc"" LIKE '%$v%')"
}
if ($Subject) { $clauses += """urn:schemas:httpmail:subject"" LIKE '%$(ConvertTo-DaslLiteral $Subject)%'" }
if ($Body)    { $clauses += """urn:schemas:httpmail:textdescription"" LIKE '%$(ConvertTo-DaslLiteral $Body)%'" }
if ($Text) {
    $v = ConvertTo-DaslLiteral $Text
    $clauses += "(""urn:schemas:httpmail:subject"" LIKE '%$v%' OR ""urn:schemas:httpmail:textdescription"" LIKE '%$v%')"
}
if ($AnyOf.Count -gt 0) {
    $ors = @()
    foreach ($term in $AnyOf) {
        if ([string]::IsNullOrWhiteSpace($term)) { continue }
        $v = ConvertTo-DaslLiteral $term
        $ors += """urn:schemas:httpmail:subject"" LIKE '%$v%' OR ""urn:schemas:httpmail:textdescription"" LIKE '%$v%'"
    }
    if ($ors.Count -gt 0) { $clauses += '(' + ($ors -join ' OR ') + ')' }
}
if ($HasAttachments) { $clauses += """urn:schemas:httpmail:hasattachment"" = 1" }
if ($Unread)         { $clauses += """urn:schemas:httpmail:read"" = 0" }
$dasl = if ($clauses.Count -gt 0) { '@SQL=' + ($clauses -join ' AND ') } else { '' }

# ---- Resolve folders
$root = Get-OlFolder -Path $Folder -Store $Store
$folders = if ($AllFolders) { @(Get-OlMailFoldersRecursive -Folder $root) } else { @($root) }

$results = @()
foreach ($f in $folders) {
    if ($results.Count -ge $Max) { break }
    $items = $f.Items
    if ($dasl) { $items = $items.Restrict($dasl) }
    $items.Sort('[ReceivedTime]', $true)   # newest first

    $item = $items.GetFirst()
    while ($item -ne $null) {
        if ($results.Count -ge $Max) { break }
        # Only MailItem (Class 43); skip meeting requests, reports etc.
        if ([int]$item.Class -eq 43) {
            $rt = Get-Date $item.ReceivedTime
            if ($PSBoundParameters.ContainsKey('After') -and $rt -lt $After) { break }   # sorted desc: nothing older will match
            $skip = $false
            if ($PSBoundParameters.ContainsKey('Before') -and $rt -ge $Before) { $skip = $true }
            if (-not $skip) {
                $results += ConvertTo-OlMailSummary -Mail $item -IncludeBody:$IncludeBody -PreviewLength $PreviewLength
            }
        }
        $item = $items.GetNext()
    }
}

$out = [pscustomobject][ordered]@{
    Query = [pscustomobject][ordered]@{
        From = $From; To = $To; Subject = $Subject; Body = $Body; Text = $Text; AnyOf = @($AnyOf)
        After = if ($PSBoundParameters.ContainsKey('After')) { $After.ToString('s') } else { $null }
        Before = if ($PSBoundParameters.ContainsKey('Before')) { $Before.ToString('s') } else { $null }
        HasAttachments = [bool]$HasAttachments; Unread = [bool]$Unread
        Folders = @($folders | ForEach-Object { [string]$_.FolderPath })
        Dasl = $dasl; Max = $Max
    }
    Count   = $results.Count
    Results = $results
}
Write-OlJson -InputObject $out -OutFile $OutFile
