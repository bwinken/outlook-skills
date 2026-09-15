<#
.SYNOPSIS
    READ-ONLY. Lists calendar items in a date range (recurrences expanded) and flags overlaps.

.EXAMPLE
    # Normal:
    powershell -NoProfile -ExecutionPolicy Bypass -File <this script> [args]
    # Execution policy enforced by Group Policy:
    powershell -NoProfile -Command "$env:OUTLOOK_SKILLS_SCRIPTS='<scripts dir>'; & ([scriptblock]::Create((Get-Content -Raw -LiteralPath '<scripts dir>\<this script>'))) [args]"

    Get-OutlookCalendar.ps1                       # today
    Get-OutlookCalendar.ps1 -Days 7               # next 7 days
    Get-OutlookCalendar.ps1 -Start 2026-09-15 -End 2026-09-20 -OutFile cal.json
#>
[CmdletBinding()]
param(
    [datetime]$Start = (Get-Date).Date,
    [datetime]$End,
    [int]$Days = 1,
    [string]$Store = '',
    [switch]$IncludeFree,          # include items whose BusyStatus is Free
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

if (-not $PSBoundParameters.ContainsKey('End')) { $End = $Start.Date.AddDays($Days) }

$ns = Connect-Outlook
$cal = if ($Store) {
    (Get-OlStores | Where-Object { $_.DisplayName -eq $Store } | Select-Object -First 1).GetDefaultFolder(9)
} else {
    $ns.GetDefaultFolder(9)
}

# Canonical recurrence-safe pattern: sort by Start, IncludeRecurrences, then Restrict with an overlap test.
# Jet syntax dates use the current culture's short date/time format.
$items = $cal.Items
$items.Sort('[Start]')
$items.IncludeRecurrences = $true
$filter = "[Start] < '$($End.ToString('g'))' AND [End] > '$($Start.ToString('g'))'"
$items = $items.Restrict($filter)

$appts = @()
$it = $items.GetFirst()
while ($it -ne $null) {
    if ([int]$it.Class -eq 26) {   # olAppointment
        $s = ConvertTo-OlAppointmentSummary -Appt $it
        if ($IncludeFree -or $s.BusyStatus -ne 'Free') { $appts += $s }
    }
    $it = $items.GetNext()
}
$appts = @($appts | Sort-Object -Property Start)

# ---- Overlap detection (ignores all-day items)
$conflicts = @()
for ($i = 0; $i -lt $appts.Count; $i++) {
    if ($appts[$i].AllDayEvent) { continue }
    for ($j = $i + 1; $j -lt $appts.Count; $j++) {
        if ($appts[$j].AllDayEvent) { continue }
        if ([datetime]$appts[$j].Start -lt [datetime]$appts[$i].End) {
            $conflicts += [pscustomobject]@{
                A = $appts[$i].Subject; AStart = $appts[$i].Start; AEnd = $appts[$i].End
                B = $appts[$j].Subject; BStart = $appts[$j].Start; BEnd = $appts[$j].End
            }
        } else { break }
    }
}

$out = [pscustomobject][ordered]@{
    Range     = [pscustomobject]@{ Start = $Start.ToString('s'); End = $End.ToString('s'); Filter = $filter }
    Calendar  = [string]$cal.FolderPath
    Count     = $appts.Count
    Conflicts = $conflicts
    Items     = $appts
}
Write-OlJson -InputObject $out -OutFile $OutFile
