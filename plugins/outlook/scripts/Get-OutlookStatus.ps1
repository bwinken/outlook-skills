<#
.SYNOPSIS
    READ-ONLY. Reports Outlook version, profiles, accounts, data files (.pst/.ost) and folder counts.

.EXAMPLE
    # Normal:
    powershell -NoProfile -ExecutionPolicy Bypass -File <this script> [args]
    # Execution policy enforced by Group Policy:
    powershell -NoProfile -Command "$env:OUTLOOK_SKILLS_SCRIPTS='<scripts dir>'; & ([scriptblock]::Create((Get-Content -Raw -LiteralPath '<scripts dir>\<this script>'))) [args]"

    powershell -NoProfile -ExecutionPolicy Bypass -File Get-OutlookStatus.ps1
    powershell -NoProfile -ExecutionPolicy Bypass -File Get-OutlookStatus.ps1 -OutFile status.json
#>
[CmdletBinding()]
param(
    [string]$OutFile = '',
    [switch]$SkipCom      # Only inspect registry and file system; do not talk to Outlook.
)

# ---- Load the shared read-only helpers without going through the execution policy.
#      $PSScriptRoot is set when run with -File; when run via -Command/[scriptblock]::Create it is
#      empty, so the caller sets OUTLOOK_SKILLS_SCRIPTS to this folder instead.
$scriptsDir = if ($PSScriptRoot) { $PSScriptRoot } elseif ($env:OUTLOOK_SKILLS_SCRIPTS) { $env:OUTLOOK_SKILLS_SCRIPTS } else {
    throw 'Cannot locate OutlookReadOnly.ps1. Run with -File, or set $env:OUTLOOK_SKILLS_SCRIPTS to the plugin scripts folder.'
}
. ([scriptblock]::Create((Get-Content -Raw -LiteralPath (Join-Path $scriptsDir 'OutlookReadOnly.ps1'))))
Initialize-OutlookConsole

$report = [ordered]@{
    GeneratedAt   = (Get-Date).ToString('s')
    Machine       = $env:COMPUTERNAME
    User          = $env:USERNAME
    OfficeVersion = $null
    NewOutlookEnabled = $null
    DefaultProfile = $null
    Profiles      = @()
    Accounts      = @()
    Stores        = @()
    DataFilesOnDisk = @()
    Warnings      = @()
}

# ---- Registry: Office version, profiles, New Outlook toggle -------------------------------
foreach ($ver in '16.0', '15.0', '14.0') {
    $key = "HKCU:\Software\Microsoft\Office\$ver\Outlook"
    if (Test-Path $key) {
        $report.OfficeVersion = $ver
        $o = Get-ItemProperty -Path $key -ErrorAction SilentlyContinue
        if ($o -and $o.PSObject.Properties['DefaultProfile']) { $report.DefaultProfile = $o.DefaultProfile }
        $pk = "$key\Profiles"
        if (Test-Path $pk) {
            $report.Profiles = @(Get-ChildItem -Path $pk -ErrorAction SilentlyContinue | ForEach-Object { $_.PSChildName })
        }
        $pref = Get-ItemProperty -Path "$key\Preferences" -ErrorAction SilentlyContinue
        if ($pref -and $pref.PSObject.Properties['UseNewOutlook']) {
            $report.NewOutlookEnabled = ([int]$pref.UseNewOutlook -eq 1)
        }
        break
    }
}
if (-not $report.OfficeVersion) { $report.Warnings += 'No Classic Outlook registry hive found under HKCU. Outlook may not be installed or never launched.' }
if ($report.NewOutlookEnabled) { $report.Warnings += 'New Outlook toggle is ON. COM automation only works with Classic Outlook.' }

# ---- File system: .pst/.ost in default locations ---------------------------------------
$candidates = @(
    (Join-Path $env:LOCALAPPDATA 'Microsoft\Outlook'),
    (Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'Outlook Files'),
    (Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'Outlook 檔案')
)
foreach ($dir in $candidates) {
    if (Test-Path $dir) {
        Get-ChildItem -Path $dir -Include *.pst, *.ost -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object {
            $report.DataFilesOnDisk += [pscustomobject]@{
                Path      = $_.FullName
                Type      = $_.Extension.TrimStart('.').ToUpper()
                SizeMB    = [math]::Round($_.Length / 1MB, 1)
                Modified  = $_.LastWriteTime.ToString('s')
            }
        }
    }
}

# ---- COM: accounts, stores, folder counts ---------------------------------------------
if (-not $SkipCom) {
    try {
        $ns  = Connect-Outlook
        $app = Get-OutlookApplication
        $report['OutlookExeVersion'] = [string]$app.Version

        foreach ($a in $ns.Accounts) {
            $type = switch ([int]$a.AccountType) { 0 {'Exchange'} 1 {'IMAP'} 2 {'POP3'} 3 {'HTTP'} 4 {'EAS'} 5 {'Other'} default {'Unknown'} }
            $report.Accounts += [pscustomobject]@{
                DisplayName = [string]$a.DisplayName
                SmtpAddress = [string]$a.SmtpAddress
                UserName    = [string]$a.UserName
                Type        = $type
            }
        }

        foreach ($s in Get-OlStores) {
            $path = ''
            try { $path = [string]$s.FilePath } catch { }
            $sizeMB = $null
            if ($path -and (Test-Path $path)) { $sizeMB = [math]::Round((Get-Item $path).Length / 1MB, 1) }
            $storeType = switch ([int]$s.ExchangeStoreType) { 0 {'PrimaryExchangeMailbox'} 1 {'ExchangePublicFolder'} 2 {'DelegateExchangeMailbox'} 3 {'NotExchange'} 4 {'AdditionalExchangeMailbox'} default {'Unknown'} }

            $folders = @()
            try {
                foreach ($f in $s.GetRootFolder().Folders) {
                    $unread = $null
                    try { $unread = [int]$f.UnReadItemCount } catch { }
                    $folders += [pscustomobject]@{
                        Name       = [string]$f.Name
                        Items      = [int]$f.Items.Count
                        Unread     = $unread
                        SubFolders = [int]$f.Folders.Count
                    }
                }
            } catch { $report.Warnings += "Could not enumerate folders of store '$($s.DisplayName)': $($_.Exception.Message)" }

            $report.Stores += [pscustomobject]@{
                DisplayName     = [string]$s.DisplayName
                FilePath        = $path
                SizeMB          = $sizeMB
                StoreType       = $storeType
                IsDataFileStore = [bool]$s.IsDataFileStore
                IsCachedExchange = [bool]$s.IsCachedExchange
                Folders         = $folders
            }
        }
    } catch {
        $report.Warnings += "COM automation failed: $($_.Exception.Message)"
    }
}

Write-OlJson -InputObject ([pscustomobject]$report) -OutFile $OutFile
