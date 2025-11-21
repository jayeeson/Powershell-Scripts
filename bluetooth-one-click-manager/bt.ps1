<#
.SYNOPSIS
    Bluetooth one-click manager
#>

param(
    [switch]$List,
    [string]$Connect,
    [string]$Disconnect,
    [string]$Forget,
    [string]$Reset,
    [switch]$AddDevice,
    [switch]$GenerateShortcuts,
    [switch]$Help,
    [switch]$Edit
)

$ErrorActionPreference = 'Stop'
$ScriptPath = $PSScriptRoot
$ConfigFile = Join-Path $ScriptPath "btconfig.json"

# Create config if missing
if (-not (Test-Path $ConfigFile)) {
    @{ manual_connect = @(); auto_connect_allowed = @() } | ConvertTo-Json -Depth 10 | Set-Content $ConfigFile -Encoding UTF8
    Write-Host "Created empty btconfig.json" -ForegroundColor Green
}
$Config = Get-Content $ConfigFile -Raw | ConvertFrom-Json

# =============================================================================
# Helpers
# =============================================================================
function Resolve-Device($id) {
    $id = $id.Trim()
    $clean = $id -replace '[-:]',''

    $all = $Config.manual_connect + $Config.auto_connect_allowed

    # MAC match
    $m = $all | Where-Object { $_.mac -replace '[-:]','' -eq $clean }
    if ($m) { return $m }

    # Friendly name partial match
    $m = $all | Where-Object { $_.friendly_name -and $_.friendly_name -match "(?i)$([regex]::Escape($id))" }
    if ($m.Count -eq 1) { return $m[0] }
    if ($m.Count -gt 1) { Write-Error "Ambiguous name '$id' → $($m.friendly_name -join ', ')"; return $null }

    Write-Error "Device not found: $id"
    return $null
}

function Get-BestName($dev) { $dev.friendly_name ?? $dev.name ?? $dev.mac }
function Save-Config { $Config | ConvertTo-Json -Depth 10 | Set-Content $ConfigFile -Encoding UTF8 }

function Show-Status {
    Write-Host "`nCurrently known devices:" -ForegroundColor Cyan
    Write-Host "(     MAC                  Name                  Description)`n" -ForegroundColor Cyan
    btdiscovery.exe | ForEach-Object { Write-Host "$_" }
}

function Toast($title, $text) {
    try { New-BurntToastNotification -Text $title,$text -ErrorAction SilentlyContinue }
    catch { Write-Host "$title — $text" -ForegroundColor Yellow }
}

function Get-IconPath($dev, $action) {
    if ($dev.icons -and $dev.icons.$action) {
        return Join-Path $ScriptPath "icons/$($dev.icons.$action)"
    }
    return Join-Path $ScriptPath "icons/bluetooth.ico"
}

function Write-Help {
    Write-Host @"
Bluetooth Manager — official commands only

    bt.ps1 -List
    bt.ps1 -Connect "MyMouse"       (name or MAC)
    bt.ps1 -Disconnect "Headset"
    bt.ps1 -Reset "Keyboard"
    bt.ps1 -AddDevice
    bt.ps1 -GenerateShortcuts

Requires: 
    1. Bluetooth Command Line Tools 1.2.0.56 (btpair.exe, btdiscovery.exe)
       Download: https://bluetoothinstaller.com/bluetooth-command-line-tools/CS-307.exe
       Make sure btpair.exe, btdiscovery.exe and btcom.exe are either in the system PATH or in the same folder as this script.
    2. PowerShell 7+
       Download: https://github.com/PowerShell/PowerShell/releases
       Make sure pwsh.exe is either in the system PATH or in the same folder as this script.
    3. (Optional) BurntToast module for notifications:
       In PowerShell 7+, run `Install-Module -Name BurntToast`
"@ -ForegroundColor Cyan
}

# =============================================================================
# Actions
# =============================================================================
if ($List) {
    Write-Host "`nManual-connect devices:`n" -ForegroundColor Magenta
    $Config.manual_connect | Format-Table @{L="Name";E={Get-BestName $_}}, mac

    Write-Host "`nAuto-connect / resettable devices:`n" -ForegroundColor Green
    $Config.auto_connect_allowed | Format-Table @{L="Name";E={Get-BestName $_}}, mac

    Show-Status
    return
}

if ($Connect) {
    $dev = Resolve-Device $Connect; if (!$dev) { exit 1 }
    Write-Host "Make sure you are in pairing mode..." -ForegroundColor Yellow
    Write-Host "Connecting $(Get-BestName $dev) ..." -ForegroundColor Cyan

    btpair.exe -p -b "$($dev.mac)" | Out-Null
        
    if ($LASTEXITCODE -ne 0) { Toast "Bluetooth" "Connect failed"; return }
    return
}

if ($Disconnect) {
    $dev = Resolve-Device $Disconnect; if (!$dev) { exit 1 }
    Write-Host "Disconnecting $(Get-BestName $dev) ..." -ForegroundColor Cyan

    if( -not $dev.service_classes ) {
        Write-Host "Can't disconnect from this device — no service classes defined." -ForegroundColor Yellow
        return
    }

    foreach ( $class in $dev.service_classes ) {
        btcom.exe -r -b $dev.mac -s $class | Out-Null
    }
    return
}

if ($Forget) {
    $dev = Resolve-Device $Forget; if (!$dev) { exit 1 }
    Write-Host "Forgetting device $(Get-BestName $dev) ..." -ForegroundColor Cyan
    btpair -u -b "$($dev.mac)" | Out-Null
    return
}

if ($Reset) {
    $dev = Resolve-Device $Reset; if (!$dev) { exit 1 }
    $name = Get-BestName $dev
    Write-Host "`nResetting $name ..." -ForegroundColor Yellow
    
    try {
        btpair -u -b "$($dev.mac)" | Out-Null
        Write-Host "Device forgotten." -ForegroundColor Cyan
    } catch {
        Write-Error "Failed to remove device"  
    }

    btpair.exe -p -b "$($dev.mac)" | Out-Null

    if ($LASTEXITCODE -ne 0) { Toast "Bluetooth Reset failed" "Could not connect to $name" }
    return
}

if ($AddDevice) {
    Write-Host "Put your device into pairing mode now..." -ForegroundColor Cyan
    Read-Host "Press ENTER when ready"
    $scan = btdiscovery.exe -i10

    $devices = @()
    foreach ($line in $scan -split "`r?`n") {
        if ($line -match '^\(([0-9A-Fa-f:]{17})\)\s+(.+?)\s+(\S+)\s*$') {
            $devices += [pscustomobject]@{ Mac = $matches[1]; Name = $matches[2].Trim() }
        }
    }

    # Filter list into two different lists: known and unknown devices
    $knownDevices = @()
    $unknownDevices = @()
    foreach ($dev in $devices) {
        $isKnown = $false
        foreach ($kdev in $Config.manual_connect + $Config.auto_connect_allowed) {
            if ($dev.Mac -eq $kdev.mac) {
                $knownDevices += $dev
                $isKnown = $true
                break
            }
        }
        if (-not $isKnown) {
            $unknownDevices += $dev
        }
    }
    # $knownDevices = $Config.manual_connect + $Config.auto_connect_allowed | ForEach-Object { $_.mac }
    # $unknownDevices = $devices | ForEach-Object { $_.Mac }

    Write-Host "`nKnown devices:" -ForegroundColor Magenta
    $knownDevices | ForEach-Object { Write-Host "$([char](65 + [array]::IndexOf($knownDevices,$_))): $($_.Mac)   $($_.Name)" }
    Write-Host "`nNew devices found:" -ForegroundColor Green
    $unknownDevices | ForEach-Object { Write-Host "$([array]::IndexOf($unknownDevices,$_) +1): $($_.Mac)   $($_.Name)" }

    return;

    if ($unknownDevices.Count -eq 0) { Write-Host "No new devices found." -ForegroundColor Red; return }

    $unknownDevices | ForEach-Object { Write-Host "$([array]::IndexOf($unknownDevices,$_) +1): $($_.Mac)   $($_.Name)" }
    $choice = Read-Host "`nEnter number or MAC"
    if ($choice -match '^\d+$' -and $choice -le $unknownDevices.Count) { $sel = $unknownDevices[$choice-1] }
    else { $sel = $unknownDevices | Where-Object { $_.Mac -replace '[-:]','' -like "*$($choice -replace '[-:]','')*" } | Select-Object -First 1 }
    if (!$sel) { Write-Error "Invalid selection"; return }

    Write-Host "Pairing device $($sel.Name) [$($sel.Mac)] ..." -ForegroundColor Cyan
    btpair.exe -p -b "$($sel.Mac)" | Out-Null
    if ($LASTEXITCODE -ne 0) { Write-Error "Pairing failed... Adding device to the config anyway.`n";}

    $friendly = Read-Host "Friendly name [$($sel.Name)]"
    if (!$friendly) { $friendly = $sel.Name }

    $isAudio = (Read-Host "Is this an audio device? (y/n)") -match '^y'
    if ($isAudio) {
        # Write-Host "Since it's an audio device, setting service classs 110B (Audio Sink) and 111e (Handsfree)." -ForegroundColor Yellow
        $serviceClasses = @("110B", "111e")
    } else {
        $serviceClasses = @()
    }

    $type = Read-Host "1) Manual connect only   2) Auto-connect + reset [1/2]"

    $icons = @{}
    if ($isAudio -and $type -eq "1") {
        $icons = @{
            add = "bluetooth-buds-add.ico"
            remove = "bluetooth-buds-remove.ico"
        }
    }

    $shortcuts = @()
    if ($type -eq "1") {
        if ((Read-Host "Create desktop shortcut to Connect? (y/n)") -match '^y') { $shortcuts += "connect" }
        if ((Read-Host "Create desktop shortcut to Disconnect? (y/n)") -match '^y') { $shortcuts += "disconnect" }
        if ((Read-Host "Create desktop shortcut to Forget? (y/n)") -match '^y') { $shortcuts += "forget" }
        $Config.manual_connect += $new
    } else {
        if( (Read-Host "Create desktop shortcut to Reset (a combination of Connect + Forget)? (y/n)") -match '^y') { $shortcuts += "reset" }
        $Config.auto_connect_allowed += $new
    }
    
    $new = @{ 
        friendly_name = $friendly;
        name = $sel.Name;
        mac = $sel.Mac;
        service_classes = $serviceClasses;
        icons = $icons;
        shortcuts = $shortcuts
    }
    Save-Config
    Write-Host "Device added!" -ForegroundColor Green
    & $PSCommandPath -GenerateShortcuts
    return
}

if ($GenerateShortcuts) {
    $desktop = [Environment]::GetFolderPath("Desktop")
    foreach ($dev in $Config.manual_connect) {
        $safe = (Get-BestName $dev) -replace '[<>:"/\\|?*]', '_'
        
        # Connect shortcut
        $bat = Join-Path $ScriptPath "bat/" "Connect-$safe.bat"
        $content = "@echo off`r`npwsh -NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -Connect `"$($dev.mac)`""
        Set-Content $bat $content -Encoding ASCII

        if ($dev.shortcuts -contains "connect") {
            $lnk = "$desktop\Connect $(Get-BestName $dev).lnk"
            $ws = New-Object -ComObject WScript.Shell
            $s = $ws.CreateShortcut($lnk)
            $s.TargetPath = $bat
            $s.IconLocation =  Get-IconPath $dev "add"
            $s.Save()
            Write-Host "Desktop shortcut → Connect $(Get-BestName $dev)"
        }
        
        # Forget shortcut
        $batDisc = Join-Path $ScriptPath "bat/" "Forget-$safe.bat"
        $contentDisc = "@echo off`r`npwsh -NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -Forget `"$($dev.mac)`""
        Set-Content $batDisc $contentDisc -Encoding ASCII

        if ($dev.shortcuts -contains "forget") {
            $lnkDisc = "$desktop\Forget $(Get-BestName $dev).lnk"
            $ws = New-Object -ComObject WScript.Shell
            $sDisc = $ws.CreateShortcut($lnkDisc)
            $sDisc.TargetPath = $batDisc
            $sDisc.IconLocation =  Get-IconPath $dev "remove"
            $sDisc.Save()
            Write-Host "Desktop shortcut → Forget $(Get-BestName $dev)"
        }
    }
    
    foreach ($dev in $Config.auto_connect_allowed) {
        $safe = (Get-BestName $dev) -replace '[<>:"/\\|?*]', '_'
        
        # Reset shortcut
        $batReset = Join-Path $ScriptPath "bat/" "Reset-$safe.bat"
        $contentReset = "@echo off`r`npwsh -NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -Reset `"$($dev.mac)`""
        Set-Content $batReset $contentReset -Encoding ASCII
        
        if ($dev.shortcuts -contains "reset") {
            $lnkReset = "$desktop\Reset $(Get-BestName $dev).lnk"
            $ws = New-Object -ComObject WScript.Shell
            $sReset = $ws.CreateShortcut($lnkReset)
            $sReset.TargetPath = $batReset
            $sReset.IconLocation =  Get-IconPath $dev "reset"
            $sReset.Save()
            Write-Host "Desktop shortcut → Reset $(Get-BestName $dev)"
        }
    }
    
    Write-Host "`nShortcuts updated!" -ForegroundColor Green
    return
}

if ($Edit) {
    Start-Process notepad.exe $ConfigFile
    return
}

if ($Help) {
    Write-Help
    return
}

Write-Help
