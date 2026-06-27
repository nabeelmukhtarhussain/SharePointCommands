<#
.SYNOPSIS
    Generates an Excel report of every SharePoint Online document library across
    a Microsoft 365 tenant, including each library's address, size, and the
    parent site's visibility and external-sharing status.

.DESCRIPTION
    A single, self-contained script. Paste it into ANY PowerShell window
    (even Windows PowerShell 5.1) and it will:
      1. Detect PowerShell 7; install it automatically if missing.
      2. Re-launch itself in PowerShell 7.
      3. Install the required modules (PnP.PowerShell, ImportExcel).
      4. Register a one-time Entra ID app for interactive sign-in (cached after).
      5. Grant the running admin Site Collection Administrator rights on every
         site so no site is skipped for lack of access.
      6. Scan every document library and export a multi-sheet Excel workbook.

.OUTPUT
    LibraryReport.xlsx with the following worksheets:
      - All Libraries : Site, URL, Visibility, External Sharing, Library,
                        Address, Item count, Size (GB), Category
      - Under 50GB / 50-100GB / Over 100GB : libraries grouped by size
      - Summary       : totals and counts

.REQUIREMENTS
    - A SharePoint Administrator account.
    - Run the window "As Administrator" (needed for the PowerShell 7 install).
    - First run prompts for one Entra app consent (Global Admin may be required).

.NOTES
    Visibility (Public/Private) and external-sharing status are reported at the
    SITE level and applied to that site's libraries. Per-library unique
    permissions are intentionally not enumerated (too slow for large tenants).
    The Public/Private column requires Group.Read.All consent; without it that
    column is left blank, while External Sharing still works.
#>

# ===================== MAIN REPORT SCRIPT (payload) =====================
$Payload = @'
$ErrorActionPreference = "Stop"

if (-not (Get-Module -ListAvailable PnP.PowerShell)) {
    Write-Host "Installing PnP.PowerShell..." -ForegroundColor Cyan
    Install-Module PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
}
if (-not (Get-Module -ListAvailable ImportExcel)) {
    Write-Host "Installing ImportExcel..." -ForegroundColor Cyan
    Install-Module ImportExcel -Scope CurrentUser -Force -AllowClobber
}
Import-Module PnP.PowerShell
Import-Module ImportExcel

# ---- Sign in once; tenant is derived from the admin's email (no separate prompt) ----
function Register-AppAndGetClientId {
    param([string]$Tenant)
    try {
        $r = Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "SPODocLibraryReport" -Tenant $Tenant
        $cid = $r.'AzureAppId/ClientId'; if (-not $cid) { $cid = $r.ClientId }; if (-not $cid) { $cid = $r.'Client Id' }
        return $cid
    } catch {
        Write-Host "Auto-registration issue: $($_.Exception.Message)" -ForegroundColor DarkYellow
        return $null
    }
}

# Helper: connect, and if the cached app belongs to a different tenant, self-heal.
function Connect-Smart {
    param([string]$Url, [string]$TenantDomain, [string]$CacheFile)

    $cid = $null
    if (Test-Path $CacheFile) { $cid = (Get-Content $CacheFile -Raw).Trim() }

    for ($attempt = 1; $attempt -le 2; $attempt++) {
        if (-not $cid) {
            Write-Host "Registering an Entra ID app for this tenant (one-time; sign in + consent)..." -ForegroundColor Yellow
            $cid = Register-AppAndGetClientId -Tenant $TenantDomain
            if ($cid) {
                $cid | Out-File -FilePath $CacheFile -Encoding ascii
                Write-Host "Waiting 60s for permissions to propagate..." -ForegroundColor Cyan
                Start-Sleep -Seconds 60
            } else {
                throw "Could not register the app. A Global Administrator may need to consent."
            }
        }
        try {
            Connect-PnPOnline -Url $Url -Interactive -ClientId $cid -ErrorAction Stop
            return $cid   # success
        } catch {
            # Stale / wrong-tenant app id -> clear cache and register fresh, then retry once.
            if ("$($_.Exception.Message)" -match "was not found in the directory|AADSTS700016|AADSTS90002") {
                Write-Host "Cached app does not belong to this tenant - re-registering automatically..." -ForegroundColor DarkYellow
                Remove-Item $CacheFile -Force -ErrorAction SilentlyContinue
                $cid = $null
                continue
            }
            throw
        }
    }
    throw "Unable to connect after re-registering."
}

# Derive the tenant from the admin's email so nothing extra has to be remembered.
$AdminUpn = Read-Host "Enter your SharePoint admin email (e.g. admin@contoso.onmicrosoft.com)"
$TenantDomain = ($AdminUpn -split "@")[1]
$TenantName   = ($TenantDomain -split "\.")[0]

# Most tenants: the guess is correct. For custom/vanity domains, allow a quick override.
$override = Read-Host "Detected tenant '$TenantName' - press Enter to accept, or type the correct SharePoint tenant name"
if ($override) { $TenantName = $override.Trim() }
$TenantDomain = "$TenantName.onmicrosoft.com"

$AdminUrl   = "https://$TenantName-admin.sharepoint.com"
$RootUrl    = "https://$TenantName.sharepoint.com"
$scriptDir  = (Get-Location).Path
$cfgFile    = Join-Path $scriptDir "pnp-clientid-$TenantName.txt"

Write-Host "Connecting to the SharePoint Admin Center for '$TenantName'..." -ForegroundColor Cyan
$ClientId = Connect-Smart -Url $AdminUrl -TenantDomain $TenantDomain -CacheFile $cfgFile

# ---- All sites ----
$sites = Get-PnPTenantSite -ErrorAction SilentlyContinue | Where-Object { $_.Template -notlike "RedirectSite*" }
Write-Host "Found $($sites.Count) sites." -ForegroundColor Green

# ---- M365 Group visibility map (Public/Private). Skips gracefully if no consent ----
$groupVis = @{}
try {
    Get-PnPMicrosoft365Group -ErrorAction Stop | ForEach-Object { $groupVis[$_.GroupId.ToString()] = $_.Visibility }
} catch {
    Write-Host "  (Skipping Public/Private column - Group.Read.All consent not granted. External Sharing will still be reported.)" -ForegroundColor DarkYellow
}

# ---- Pass 1: grant Site Collection Admin on every site (removes 'no access' skips) ----
Write-Host "Ensuring access (Site Collection Administrator)..." -ForegroundColor Cyan
foreach ($site in $sites) {
    try { Set-PnPTenantSite -Identity $site.Url -Owners $AdminUpn -ErrorAction Stop }
    catch { Write-Host "  ! Could not grant admin on: $($site.Url)" -ForegroundColor DarkYellow }
}

# ---- Pass 2: scan libraries ----
$rows = New-Object System.Collections.Generic.List[object]
Write-Host "Scanning libraries...`n" -ForegroundColor Green
foreach ($site in $sites) {

    $gid = if ($site.GroupId) { $site.GroupId.ToString() } else { "" }
    $visibility = if ($gid -and $groupVis.ContainsKey($gid)) { $groupVis[$gid] } else { "-" }
    $sharing = switch ("$($site.SharingCapability)") {
        "Disabled"                        { "Private (internal only)" }
        "ExistingExternalUserSharingOnly" { "Shared (existing guests)" }
        "ExternalUserSharingOnly"         { "Shared (external users)" }
        "ExternalUserAndGuestSharing"     { "Shared (external + anyone)" }
        default                           { "$($site.SharingCapability)" }
    }

    try {
        Connect-PnPOnline -Url $site.Url -Interactive -ClientId $ClientId
        $libs = Get-PnPList -Includes RootFolder | Where-Object { $_.BaseTemplate -eq 101 -and -not $_.Hidden }
        foreach ($lib in $libs) {
            try { $sizeGB = [math]::Round((Get-PnPFolderStorageMetric -List $lib).TotalSize / 1GB, 2) }
            catch { $sizeGB = 0 }
            $address = $RootUrl + $lib.RootFolder.ServerRelativeUrl
            $category = if ($sizeGB -lt 50) { "Under 50GB" } elseif ($sizeGB -lt 100) { "50-100GB" } else { "Over 100GB" }
            $rows.Add([PSCustomObject]@{
                SiteTitle=$site.Title; SiteUrl=$site.Url
                Visibility=$visibility; ExternalSharing=$sharing
                LibraryName=$lib.Title; Address=$address
                Items=$lib.ItemCount; SizeGB=$sizeGB; Category=$category
            })
            Write-Host ("  {0,-8} GB | {1,-22} | {2} | {3}" -f $sizeGB, $visibility, $lib.Title, $site.Title) -ForegroundColor Gray
        }
    } catch { Write-Host "  ! Still no access: $($site.Url)" -ForegroundColor DarkYellow }
}

# ---- Build Excel (falls back to a timestamped name if the file is open) ----
$xlsx = Join-Path $scriptDir "LibraryReport.xlsx"
try {
    if (Test-Path $xlsx) { Remove-Item $xlsx -Force -ErrorAction Stop }
} catch {
    $xlsx = Join-Path $scriptDir ("LibraryReport_{0}.xlsx" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
    Write-Host "Existing file is open - writing to a new file: $xlsx" -ForegroundColor DarkYellow
}

$sorted  = $rows | Sort-Object SizeGB -Descending
$under50 = $sorted | Where-Object { $_.Category -eq "Under 50GB" }
$mid     = $sorted | Where-Object { $_.Category -eq "50-100GB" }
$over100 = $sorted | Where-Object { $_.Category -eq "Over 100GB" }
$summary = [PSCustomObject]@{
    "Total Libraries"=$rows.Count
    "Combined Size (GB)"=[math]::Round(($rows | Measure-Object SizeGB -Sum).Sum,2)
    "Under 50 GB"=$under50.Count; "50 - 100 GB"=$mid.Count; "Over 100 GB"=$over100.Count
    "Public sites"=($rows | Where-Object {$_.Visibility -eq "Public"} | Select-Object SiteUrl -Unique).Count
    "Private sites"=($rows | Where-Object {$_.Visibility -eq "Private"} | Select-Object SiteUrl -Unique).Count
}
$xp = @{ AutoSize=$true; FreezeTopRow=$true; BoldTopRow=$true }
$sorted  | Export-Excel -Path $xlsx -WorksheetName "All Libraries" -TableName "AllLibs" -TableStyle Medium2 @xp
$under50 | Export-Excel -Path $xlsx -WorksheetName "Under 50GB"   -TableName "Under50" -TableStyle Medium6 @xp
$mid     | Export-Excel -Path $xlsx -WorksheetName "50-100GB"     -TableName "Mid"     -TableStyle Medium5 @xp
$over100 | Export-Excel -Path $xlsx -WorksheetName "Over 100GB"   -TableName "Over100" -TableStyle Medium3 @xp
$summary | Export-Excel -Path $xlsx -WorksheetName "Summary" -Title "Document Library Size Report" -TitleBold @xp

Write-Host "`n=================== SUMMARY ===================" -ForegroundColor Cyan
Write-Host ("Total document libraries : {0}" -f $rows.Count)   -ForegroundColor White
Write-Host ("Under 50 GB              : {0}" -f $under50.Count) -ForegroundColor Green
Write-Host ("50 - 100 GB             : {0}" -f $mid.Count)      -ForegroundColor Green
Write-Host ("Over 100 GB             : {0}" -f $over100.Count)  -ForegroundColor Green
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host "`nExcel report ready: $xlsx" -ForegroundColor Cyan
Disconnect-PnPOnline
'@
# ===================== end of payload =====================


# ---------- Already in PowerShell 7 -> run directly ----------
if ($PSVersionTable.PSVersion.Major -ge 7) {
    Invoke-Expression $Payload
    return
}

# ---------- Otherwise locate or install PowerShell 7 ----------
function Get-Pwsh {
    $cmd = (Get-Command pwsh -ErrorAction SilentlyContinue).Source
    if ($cmd) { return $cmd }
    foreach ($p in @(
        "$env:ProgramFiles\PowerShell\7\pwsh.exe",
        "${env:ProgramFiles(x86)}\PowerShell\7\pwsh.exe",
        "$env:LOCALAPPDATA\Microsoft\PowerShell\7\pwsh.exe"
    )) { if (Test-Path $p) { return $p } }
    return $null
}

[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor 3072
$pwshPath = Get-Pwsh

if (-not $pwshPath) {
    Write-Host "PowerShell 7 not found - installing it (this may take a moment)..." -ForegroundColor Yellow
    try {
        Invoke-Expression "& { $(Invoke-RestMethod https://aka.ms/install-powershell.ps1) } -UseMSI -Quiet"
    } catch {
        Write-Host "Auto-install failed: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Open the window 'As Administrator', or install manually:" -ForegroundColor Yellow
        Write-Host "  https://github.com/PowerShell/PowerShell/releases/latest" -ForegroundColor Yellow
        return
    }
    Start-Sleep -Seconds 5
    $pwshPath = Get-Pwsh
}

if (-not $pwshPath) {
    Write-Host "Installed, but pwsh.exe was not detected yet. Close this window, open a new one, and paste the script again." -ForegroundColor Yellow
    return
}

# ---------- Write payload to a temp file and run it in PowerShell 7 ----------
$temp = Join-Path $env:TEMP "SPODocLibraryReport_payload.ps1"
$Payload | Out-File -FilePath $temp -Encoding utf8 -Force
Write-Host "Running the report in PowerShell 7..." -ForegroundColor Cyan
& $pwshPath -NoExit -ExecutionPolicy Bypass -File $temp
