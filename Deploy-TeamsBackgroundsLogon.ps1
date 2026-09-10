<#
.SYNOPSIS
    Deploys custom Microsoft Teams (new client) background images. Applies them in the USER's
    context, at LOGON ONLY - never on an all-day wall-clock schedule. Runs windowless.

.DESCRIPTION
    Replacement for earlier Deploy-TeamsBackgrounds.ps1 scripts that used a <TimeTrigger>
    repeating every 30 minutes against the Users group. That produced a visible console flash
    every tick, for every logged-on session, all day.

    Changes:
      * Trigger is a single <LogonTrigger> with a delay (default 10 min) and NO <Repetition> at
        all. Combined with a per-user once-per-day guard, that means at most one real run per
        user per day. Nothing runs on a wall-clock schedule.
      * The task action is wscript.exe running a tiny VBS shim, which starts PowerShell with
        WindowStyle 0. -WindowStyle Hidden alone does NOT prevent the flash: conhost paints a
        window before PowerShell can hide itself. wscript.exe is windowless from the start.
      * Distinct paths, script name and task name so it cannot collide with older installs.

    TRADE-OFF being accepted: new Teams has not initialised at the moment of logon, so the run is
    delayed by -LogonDelayMinutes to let it come up. If Teams still is not up by then, this user
    gets nothing from that sign-in - there is no retry. The daily marker is therefore written ONLY
    by a run that actually reached Teams, so a second sign-in the same day still gets a real
    attempt. A user who signs in once, Teams is slow, and the machine stays on will wait until
    their next sign-in.

    Modes:
      -Mode Stage   (default; run once by Intune device-context / RMM as SYSTEM)
            Builds the 1920x1080 + 280x158 GUID-named pairs once into a machine payload folder,
            writes the VBS shim, registers the logon task, removes any legacy task if present,
            and kicks the task once so already-logged-on users are covered immediately.

      -Mode Apply   (invoked by the task, in the logged-on user's context)
            Exits 0 immediately if this user already completed a run today. Exits 0 if Teams has
            not initialised for this user - without consuming the daily slot. Otherwise CONVERGES
            the user's Uploads folder against the payload (restores missing, refreshes changed).

.PARAMETER Mode
    Stage (default) or Apply. You only ever invoke Stage; the task invokes Apply.

.PARAMETER ZipUrl
    Public URL to a .zip of images.

.PARAMETER ManifestUrl
    Public URL to a JSON array: [ { "Name": "...", "Url": "https://.../image.jpg" }, ... ].

.PARAMETER ImageUrls
    Direct list of public image URLs (alternative to -ManifestUrl).

.PARAMETER ImageNames
    Optional friendly names, matched positionally to -ImageUrls.

.PARAMETER ImageFolder
    Local folder of images. Must be passed explicitly to win over the default auto-detection.

.PARAMETER LogonDelayMinutes
    How long after sign-in to wait before running, giving new Teams time to start. Default 10.

.PARAMETER UseShim
    Use a wscript.exe VBS launcher to prevent the conhost flash. Default $true.
    Set to $false only if VBS is blocked by AppLocker/SRP/ASR in your environment.

.PARAMETER Force
    Bypass the once-per-day guard and re-copy even when content is unchanged.

.PARAMETER Uninstall
    Remove the task, shim, payload, and every GUID-named file this tool deployed.

.EXAMPLE
    # Stage using a remotely hosted zip (run as SYSTEM via Intune or RMM):
    powershell.exe -ExecutionPolicy Bypass -NoProfile -File Deploy-TeamsBackgroundsLogon.ps1 `
        -ZipUrl "https://yourstorage.blob.core.windows.net/container/backgrounds.zip"

.EXAMPLE
    # Stage using images in an Images\ subfolder next to the script:
    powershell.exe -ExecutionPolicy Bypass -NoProfile -File Deploy-TeamsBackgroundsLogon.ps1

.EXAMPLE
    # Uninstall:
    powershell.exe -ExecutionPolicy Bypass -NoProfile -File Deploy-TeamsBackgroundsLogon.ps1 -Uninstall

.NOTES
    Exit codes:
      0   Success (Stage completed; or Apply completed / Teams not yet signed in)
      10  Download or zip-extraction failed
      11  No images supplied / source produced zero usable entries
      12  Payload staging produced zero usable images
      13  One or more images failed to process during Stage
      14  -Uninstall requested but nothing to remove
      20  -Mode Apply but no staged payload found (Stage never ran)
      99  Unhandled exception

    PILOT CHECKS before wide release:
      * Confirm wscript.exe + local .vbs execution is permitted by your AppLocker / SRP / ASR
        posture. If VBS is blocked, set -UseShim:$false to fall back to direct powershell.exe
        (functional, but the logon flash returns).
      * Confirm a background actually appears in the Teams picker after a Teams restart.
#>

[CmdletBinding(DefaultParameterSetName = 'Install')]
param(
    [Parameter(ParameterSetName = 'Install')]
    [ValidateSet('Stage', 'Apply')]
    [string]$Mode = 'Stage',

    [Parameter(ParameterSetName = 'Install')]
    [string]$ZipUrl = '',

    [Parameter(ParameterSetName = 'Install')]
    [string]$ManifestUrl,

    [Parameter(ParameterSetName = 'Install')]
    [string[]]$ImageUrls,

    [Parameter(ParameterSetName = 'Install')]
    [string[]]$ImageNames,

    [Parameter(ParameterSetName = 'Install')]
    [string]$ImageFolder,

    [Parameter(ParameterSetName = 'Install')]
    [ValidateRange(0, 60)]
    [int]$LogonDelayMinutes = 10,

    [Parameter(ParameterSetName = 'Install')]
    [bool]$UseShim = $true,

    [Parameter(ParameterSetName = 'Install')]
    [switch]$Force,

    [Parameter(ParameterSetName = 'Uninstall', Mandatory = $true)]
    [switch]$Uninstall
)

# ------------------------------------------------------------------------------------------------
# Constants / paths
# ------------------------------------------------------------------------------------------------
$CompanyRoot     = 'C:\ProgramData\TeamsBG'
$LogDir          = Join-Path $CompanyRoot 'Logs'
$LogFile         = Join-Path $LogDir 'teamsbackground-logon.txt'
$RootDir         = Join-Path $CompanyRoot 'TeamsBackgroundsLogon'
$PayloadDir      = Join-Path $RootDir 'Payload'
$ManifestFile    = Join-Path $RootDir 'payload.json'
$SelfInstallPath = Join-Path $RootDir 'Deploy-TeamsBackgroundsLogon.ps1'
$ShimPath        = Join-Path $RootDir 'Apply-TeamsBackgrounds.vbs'
$WorkDir         = Join-Path $env:TEMP 'TeamsBG-TeamsBackgroundsLogon'
$ExtractDir      = Join-Path $WorkDir 'Extracted'

$TaskName        = 'TeamsBG-ApplyAtLogon'
$LegacyTaskName  = 'TeamsBG-Apply'

$UserRoot        = Join-Path $env:LOCALAPPDATA 'TeamsBG'
$UserLogDir      = Join-Path $UserRoot 'Logs'
$UserLogFile     = Join-Path $UserLogDir 'teamsbackground-logon-apply.txt'
$UserMarkerFile  = Join-Path $UserRoot 'teams-bg-logon-applied.json'

# New Teams per-user cache layout
$TeamsWebViewRel = 'Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\EBWebView'
$TeamsUploadsRel = 'Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\Backgrounds\Uploads'

$MainW  = 1920
$MainH  = 1080
$ThumbW = 280
$ThumbH = 158

$script:ActiveLogFile = $LogFile

# ------------------------------------------------------------------------------------------------
# Logging / ACLs
# ------------------------------------------------------------------------------------------------
function Repair-AclPermissions {
    param([Parameter(Mandatory)][string]$Path)
    try {
        if (Test-Path -LiteralPath $Path) {
            & takeown.exe /F $Path /A /R /D Y 2>&1 | Out-Null
            & icacls.exe $Path /grant '*S-1-5-32-544:(OI)(CI)F' /grant '*S-1-5-18:(OI)(CI)F' /grant '*S-1-5-32-545:(OI)(CI)RX' /T /C /Q 2>&1 | Out-Null
        }
    }
    catch { }
}

function Test-IsSystem {
    return ([Security.Principal.WindowsIdentity]::GetCurrent().User.Value -eq 'S-1-5-18')
}

function Set-LogContext {
    param([Parameter(Mandatory)][ValidateSet('Machine', 'User')][string]$Context)
    if ($Context -eq 'User') {
        foreach ($d in @($UserRoot, $UserLogDir)) {
            if (-not (Test-Path -LiteralPath $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
        }
        $script:ActiveLogFile = $UserLogFile
    }
    else {
        foreach ($d in @($LogDir, $RootDir, $PayloadDir, $WorkDir)) {
            if (-not (Test-Path -LiteralPath $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
        }
        Repair-AclPermissions -Path $RootDir
        $script:ActiveLogFile = $LogFile
    }
}

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR')][string]$Level = 'INFO'
    )
    $line = '[{0}] [{1}] [{2}] {3}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Mode, $Message
    try {
        Add-Content -LiteralPath $script:ActiveLogFile -Value $line -Encoding UTF8 -ErrorAction Stop
    }
    catch { }
    if ($Level -eq 'ERROR') { Write-Host $line -ForegroundColor Red }
    elseif ($Level -eq 'WARN') { Write-Host $line -ForegroundColor Yellow }
    else { Write-Host $line }
}

# ------------------------------------------------------------------------------------------------
# Helpers
# ------------------------------------------------------------------------------------------------
function Get-StableHash {
    param([Parameter(Mandatory)][string]$Text)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try   { $hash = $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Text)) }
    finally { $sha.Dispose() }
    $hex = -join ($hash | ForEach-Object { $_.ToString('x2') })
    return $hex.Substring(0, 16)
}

function Get-FileContentHash {
    param([Parameter(Mandatory)][string]$Path)
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.Substring(0, 16)
}

function New-DeterministicGuid {
    param([Parameter(Mandatory)][string]$Seed)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try   { $bytes = $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Seed)) }
    finally { $sha.Dispose() }
    $hex = (-join ($bytes | ForEach-Object { $_.ToString('x2') })).Substring(0, 32)
    return ([guid]$hex).ToString()
}

function Invoke-DownloadWithRetry {
    param(
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][string]$OutFile,
        [int]$Retries = 3,
        [int]$TimeoutSec = 120
    )
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    for ($i = 1; $i -le $Retries; $i++) {
        try {
            Write-Log "Downloading (attempt $i/$Retries): $Uri"
            Invoke-WebRequest -Uri $Uri -OutFile $OutFile -UseBasicParsing -TimeoutSec $TimeoutSec
            if ((Test-Path -LiteralPath $OutFile) -and (Get-Item -LiteralPath $OutFile).Length -gt 0) { return $true }
        }
        catch {
            Write-Log "Download attempt $i failed: $_" 'WARN'
            Start-Sleep -Seconds 5
        }
    }
    return $false
}

function Convert-CoverImage {
    param(
        [Parameter(Mandatory)][string]$SourcePath,
        [Parameter(Mandatory)][string]$DestPath,
        [Parameter(Mandatory)][int]$TargetWidth,
        [Parameter(Mandatory)][int]$TargetHeight
    )
    Add-Type -AssemblyName System.Drawing -ErrorAction Stop
    $src = [System.Drawing.Image]::FromFile($SourcePath)
    try {
        $scale   = [Math]::Max($TargetWidth / $src.Width, $TargetHeight / $src.Height)
        $scaledW = [int][Math]::Ceiling($src.Width * $scale)
        $scaledH = [int][Math]::Ceiling($src.Height * $scale)
        $offsetX = [int](($scaledW - $TargetWidth) / 2)
        $offsetY = [int](($scaledH - $TargetHeight) / 2)

        $bmp = New-Object System.Drawing.Bitmap -ArgumentList @($TargetWidth, $TargetHeight)
        $gfx = [System.Drawing.Graphics]::FromImage($bmp)
        try {
            $gfx.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
            $gfx.SmoothingMode     = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
            $gfx.PixelOffsetMode   = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
            $destRect = New-Object System.Drawing.Rectangle -ArgumentList @((-$offsetX), (-$offsetY), $scaledW, $scaledH)
            $gfx.DrawImage($src, $destRect)
        }
        finally { $gfx.Dispose() }

        $jpegCodec = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.MimeType -eq 'image/jpeg' }
        $encParams = New-Object System.Drawing.Imaging.EncoderParameters -ArgumentList @(1)
        $encParams.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter -ArgumentList @([System.Drawing.Imaging.Encoder]::Quality, [long]90)
        $bmp.Save($DestPath, $jpegCodec, $encParams)
        $bmp.Dispose()
    }
    finally { $src.Dispose() }
}

function Get-ImageList {
    if ($ManifestUrl) {
        Write-Log "Downloading manifest from $ManifestUrl"
        try {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            $raw = Invoke-RestMethod -Uri $ManifestUrl -UseBasicParsing -TimeoutSec 30
        }
        catch { Write-Log "Failed to download manifest: $_" 'ERROR'; exit 10 }

        $list = @()
        foreach ($item in @($raw)) {
            if (-not $item.Url) { continue }
            $name = if ($item.Name) { $item.Name } else { [IO.Path]::GetFileNameWithoutExtension($item.Url) }
            $list += [pscustomobject]@{ Name = $name; Source = $item.Url; IsLocal = $false }
        }
        return , $list
    }
    elseif ($ImageUrls) {
        $list = @()
        for ($i = 0; $i -lt $ImageUrls.Count; $i++) {
            $name = if ($ImageNames -and $i -lt $ImageNames.Count) { $ImageNames[$i] } else { [IO.Path]::GetFileNameWithoutExtension($ImageUrls[$i]) }
            $list += [pscustomobject]@{ Name = $name; Source = $ImageUrls[$i]; IsLocal = $false }
        }
        return , $list
    }
    elseif ($ImageFolder) {
        if (-not (Test-Path -LiteralPath $ImageFolder)) { Write-Log "Specified -ImageFolder not found: $ImageFolder" 'ERROR'; exit 11 }
        Write-Log "Using local image folder: $ImageFolder"
        $files = Get-ChildItem -LiteralPath $ImageFolder -File -Recurse -Include '*.jpg', '*.jpeg', '*.png' -ErrorAction SilentlyContinue
        if (-not $files -or $files.Count -eq 0) { Write-Log "Folder '$ImageFolder' contained no images." 'ERROR'; exit 11 }
        $list = @(); foreach ($f in $files) { $list += [pscustomobject]@{ Name = $f.BaseName; Source = $f.FullName; IsLocal = $true } }
        return , $list
    }
    elseif ($ZipUrl) {
        Write-Log "Downloading images zip from $ZipUrl"
        $zipTemp = Join-Path $WorkDir 'images-source.zip'
        if (Test-Path -LiteralPath $ExtractDir) { Remove-Item -LiteralPath $ExtractDir -Recurse -Force -ErrorAction SilentlyContinue }
        New-Item -ItemType Directory -Path $ExtractDir -Force | Out-Null

        if (-not (Invoke-DownloadWithRetry -Uri $ZipUrl -OutFile $zipTemp -TimeoutSec 300)) {
            Write-Log "Failed to download zip from $ZipUrl after 3 attempts." 'ERROR'; exit 10
        }
        Write-Log "Zip downloaded ($([math]::Round((Get-Item -LiteralPath $zipTemp).Length / 1MB, 1)) MB). Extracting..."
        try { Expand-Archive -LiteralPath $zipTemp -DestinationPath $ExtractDir -Force }
        catch { Write-Log "Failed to extract zip: $_" 'ERROR'; exit 10 }
        Remove-Item -LiteralPath $zipTemp -Force -ErrorAction SilentlyContinue

        $files = Get-ChildItem -LiteralPath $ExtractDir -File -Recurse -Include '*.jpg', '*.jpeg', '*.png' -ErrorAction SilentlyContinue
        if (-not $files -or $files.Count -eq 0) { Write-Log 'Zip extracted but contained no images.' 'ERROR'; exit 11 }
        Write-Log "Extracted $($files.Count) image file(s) from zip."
        $list = @(); foreach ($f in $files) { $list += [pscustomobject]@{ Name = $f.BaseName; Source = $f.FullName; IsLocal = $true } }
        return , $list
    }
    else {
        $folder = Join-Path $PSScriptRoot 'Images'
        if (-not (Test-Path -LiteralPath $folder)) {
            Write-Log "No source supplied and no 'Images' folder next to the script ($folder)." 'ERROR'; exit 11
        }
        Write-Log "Using local image folder: $folder"
        $files = Get-ChildItem -LiteralPath $folder -File -Recurse -Include '*.jpg', '*.jpeg', '*.png' -ErrorAction SilentlyContinue
        if (-not $files -or $files.Count -eq 0) { Write-Log "Folder '$folder' contained no images." 'ERROR'; exit 11 }
        $list = @(); foreach ($f in $files) { $list += [pscustomobject]@{ Name = $f.BaseName; Source = $f.FullName; IsLocal = $true } }
        return , $list
    }
}

function Test-TeamsInitialized {
    param([Parameter(Mandatory)][string]$LocalAppData)
    $webview = Join-Path $LocalAppData $TeamsWebViewRel
    if (-not (Test-Path -LiteralPath $webview)) { return $false }
    $hasContent = Get-ChildItem -LiteralPath $webview -Force -ErrorAction SilentlyContinue | Select-Object -First 1
    return [bool]$hasContent
}

# ------------------------------------------------------------------------------------------------
# Windowless shim
# ------------------------------------------------------------------------------------------------
function Write-LauncherShim {
    try {
        $vbs = @"
Option Explicit
' Launches the Teams background Apply run with no visible window.
' Generated by Deploy-TeamsBackgroundsLogon.ps1 - do not edit by hand.
Dim shell, command
Set shell = CreateObject("WScript.Shell")
command = "powershell.exe -ExecutionPolicy Bypass -NoProfile -NonInteractive -File ""$SelfInstallPath"" -Mode Apply"
shell.Run command, 0, False
"@
        Set-Content -LiteralPath $ShimPath -Value $vbs -Encoding ASCII -Force
        Write-Log "Wrote launcher shim: $ShimPath"
        return $true
    }
    catch {
        Write-Log "Failed to write launcher shim: $_" 'ERROR'
        return $false
    }
}

# ------------------------------------------------------------------------------------------------
# Scheduled task - LOGON ONLY, bounded retry window
# ------------------------------------------------------------------------------------------------
function Register-LogonTask {
    try {
        $delay = "PT${LogonDelayMinutes}M"

        if ($UseShim -and (Test-Path -LiteralPath $ShimPath)) {
            $command   = 'wscript.exe'
            $arguments = "//B //Nologo `"$ShimPath`""
        }
        else {
            $command   = 'powershell.exe'
            $arguments = "-ExecutionPolicy Bypass -NoProfile -NonInteractive -WindowStyle Hidden -File `"$SelfInstallPath`" -Mode Apply"
            Write-Log 'Shim not in use - task will call powershell.exe directly and a brief window may flash at logon.' 'WARN'
        }
        $argumentsEscaped = [System.Security.SecurityElement]::Escape($arguments)

        $taskXml = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.3" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Description>Applies staged Teams background images in the logged-on user context. Runs at sign-in only, $LogonDelayMinutes minute(s) after logon, at most once per user per day. No wall-clock schedule.</Description>
    <URI>\$TaskName</URI>
  </RegistrationInfo>
  <Triggers>
    <LogonTrigger>
      <Enabled>true</Enabled>
      <Delay>$delay</Delay>
    </LogonTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <GroupId>S-1-5-32-545</GroupId>
      <RunLevel>LeastPrivilege</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>false</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>false</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT10M</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>$command</Command>
      <Arguments>$argumentsEscaped</Arguments>
    </Exec>
  </Actions>
</Task>
"@

        Register-ScheduledTask -TaskName $TaskName -Xml $taskXml -Force | Out-Null

        $verify = Export-ScheduledTask -TaskName $TaskName
        if ($verify -notmatch '<LogonTrigger>') {
            Write-Log "Task '$TaskName' registered but no <LogonTrigger> is present." 'ERROR'
            return $false
        }
        if ($verify -match '<TimeTrigger>') {
            Write-Log "Task '$TaskName' unexpectedly contains a <TimeTrigger> - it would run through the day." 'ERROR'
            return $false
        }
        if ($verify -match '<Repetition>') {
            Write-Log "Task '$TaskName' registered but contains a <Repetition> - it would run more than once per sign-in." 'ERROR'
            return $false
        }

        Write-Log "Registered scheduled task '$TaskName' (sign-in only, +$LogonDelayMinutes min, max once per user per day; action: $command)."
        return $true
    }
    catch {
        Write-Log "Failed to register scheduled task '$TaskName': $_" 'ERROR'
        return $false
    }
}

function Remove-LegacyTask {
    try {
        if (Get-ScheduledTask -TaskName $LegacyTaskName -ErrorAction SilentlyContinue) {
            Stop-ScheduledTask -TaskName $LegacyTaskName -ErrorAction SilentlyContinue
            Unregister-ScheduledTask -TaskName $LegacyTaskName -Confirm:$false
            Write-Log "Removed legacy scheduled task '$LegacyTaskName'."
        }
    }
    catch { Write-Log "Could not remove legacy task '$LegacyTaskName': $_" 'WARN' }
}

function Unregister-LogonTask {
    try {
        if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
            Stop-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
            Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
            Write-Log "Removed scheduled task '$TaskName'."
        }
    }
    catch { Write-Log "Failed to remove scheduled task '$TaskName': $_" 'WARN' }
}

# ------------------------------------------------------------------------------------------------
# STAGE (SYSTEM): build payload once, write shim, register logon task
# ------------------------------------------------------------------------------------------------
function Invoke-Stage {
    Set-LogContext -Context Machine
    Write-Log '=== Teams background STAGE (logon variant) starting ==='

    if (Test-Path -LiteralPath $PayloadDir) {
        Get-ChildItem -LiteralPath $PayloadDir -File -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
    }
    else {
        New-Item -ItemType Directory -Path $PayloadDir -Force | Out-Null
    }

    $images = Get-ImageList
    if (-not $images -or $images.Count -eq 0) { Write-Log 'Source produced zero usable images.' 'ERROR'; exit 11 }

    $manifest  = @()
    $failures  = 0
    $seenGuids = @{}

    foreach ($img in $images) {
        try {
            if ($img.IsLocal) { $cacheKey = Get-FileContentHash -Path $img.Source }
            else              { $cacheKey = Get-StableHash -Text $img.Source }

            $guid = New-DeterministicGuid -Seed $cacheKey

            if ($seenGuids.ContainsKey($guid)) {
                Write-Log "Skipping duplicate image '$($img.Name)' (identical content to an earlier one)."
                continue
            }
            $seenGuids[$guid] = $true

            $mainFile  = "$guid.jpg"
            $thumbFile = "${guid}_thumb.jpg"
            $mainOut   = Join-Path $PayloadDir $mainFile
            $thumbOut  = Join-Path $PayloadDir $thumbFile

            $srcTemp = Join-Path $WorkDir "$guid.src"
            if ($img.IsLocal) {
                Copy-Item -LiteralPath $img.Source -Destination $srcTemp -Force
            }
            else {
                Write-Log "Downloading '$($img.Name)' from $($img.Source)"
                if (-not (Invoke-DownloadWithRetry -Uri $img.Source -OutFile $srcTemp -TimeoutSec 60)) {
                    throw "Failed to download after retries: $($img.Source)"
                }
            }

            Write-Log "Building payload for '$($img.Name)' -> $mainFile (1920x1080 + thumb)"
            Convert-CoverImage -SourcePath $srcTemp -DestPath $mainOut  -TargetWidth $MainW  -TargetHeight $MainH
            Convert-CoverImage -SourcePath $srcTemp -DestPath $thumbOut -TargetWidth $ThumbW -TargetHeight $ThumbH
            Remove-Item -LiteralPath $srcTemp -Force -ErrorAction SilentlyContinue

            $manifest += [pscustomobject]@{
                Name      = $img.Name
                CacheKey  = $cacheKey
                Guid      = $guid
                MainFile  = $mainFile
                ThumbFile = $thumbFile
                StagedUtc = (Get-Date).ToUniversalTime().ToString('o')
            }
        }
        catch {
            Write-Log "FAILED staging '$($img.Name)': $_" 'ERROR'
            $failures++
        }
    }

    if (-not $manifest -or $manifest.Count -eq 0) { Write-Log 'Staging produced zero usable images.' 'ERROR'; exit 12 }

    ConvertTo-Json -InputObject @($manifest) -Depth 5 | Set-Content -LiteralPath $ManifestFile -Encoding UTF8
    Write-Log "Manifest written: $ManifestFile ($($manifest.Count) image(s))."

    try {
        Copy-Item -LiteralPath $PSCommandPath -Destination $SelfInstallPath -Force
        Write-Log "Copied script to $SelfInstallPath"
    }
    catch { Write-Log "Could not copy script to $SelfInstallPath : $_" 'ERROR' }

    if ($UseShim) { $null = Write-LauncherShim }

    Repair-AclPermissions -Path $RootDir

    Remove-LegacyTask
    $null = Register-LogonTask

    Remove-Item -LiteralPath $ExtractDir -Recurse -Force -ErrorAction SilentlyContinue

    try {
        Start-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
        Write-Log "Kicked '$TaskName' once for sessions that are already signed in."
    }
    catch { }

    if ($failures -gt 0) { Write-Log "Stage completed with $failures failure(s)." 'ERROR'; exit 13 }
    Write-Log '=== Teams background STAGE completed successfully ==='
    exit 0
}

# ------------------------------------------------------------------------------------------------
# APPLY (user context, invoked by the logon task)
# ------------------------------------------------------------------------------------------------
function Invoke-Apply {
    Set-LogContext -Context User
    Write-Log "=== Teams background APPLY starting (user $env:USERNAME) ==="

    if (Test-IsSystem) {
        Write-Log 'Apply is running as SYSTEM - LOCALAPPDATA is not a real user profile. Nothing to do.' 'WARN'
        exit 0
    }

    if (-not $Force) {
        $today = (Get-Date).ToString('yyyy-MM-dd')
        try {
            if (Test-Path -LiteralPath $UserMarkerFile) {
                $prev = Get-Content -LiteralPath $UserMarkerFile -Raw | ConvertFrom-Json
                if ($prev.LastSuccessDate -eq $today) {
                    Write-Log "Already completed a run today ($today) for this user - nothing to do."
                    exit 0
                }
            }
        }
        catch { Write-Log 'Daily marker unreadable - treating as not yet run today.' 'WARN' }
    }

    if (-not (Test-Path -LiteralPath $ManifestFile)) {
        Write-Log "No staged manifest at $ManifestFile - Stage hasn't run. Nothing to do." 'WARN'
        exit 20
    }

    $staged = @(Get-ChildItem -LiteralPath $PayloadDir -File -ErrorAction SilentlyContinue |
                Where-Object { $_.Extension -match '^\.(jpg|jpeg|png)$' })

    if ($staged.Count -eq 0) {
        Write-Log "Payload folder '$PayloadDir' contains no images - Stage did not complete." 'ERROR'
        exit 20
    }
    Write-Log "Payload holds $($staged.Count) file(s)."

    if (-not (Test-TeamsInitialized -LocalAppData $env:LOCALAPPDATA)) {
        Write-Log 'New Teams not yet launched/signed in for this user - today''s slot is NOT consumed; will try again at the next sign-in.'
        exit 0
    }

    $teamsRunning = [bool](Get-Process -Name 'ms-teams' -ErrorAction SilentlyContinue)
    if ($teamsRunning) { Write-Log 'Teams is currently RUNNING (copying anyway - convergence handles any wipe).' }
    else               { Write-Log 'Teams is currently closed.' }

    $uploads = Join-Path $env:LOCALAPPDATA $TeamsUploadsRel
    try {
        if (-not (Test-Path -LiteralPath $uploads)) {
            New-Item -ItemType Directory -Path $uploads -Force -ErrorAction Stop | Out-Null
            Write-Log 'Created Uploads folder.'
        }
    }
    catch {
        Write-Log "Could not create Uploads folder '$uploads': $_" 'ERROR'
        exit 0
    }

    $copiedMissing = 0; $copiedChanged = 0; $verifiedOk = 0; $failed = 0

    foreach ($src in $staged) {
        $dst    = Join-Path $uploads $src.Name
        $reason = $null

        if (-not (Test-Path -LiteralPath $dst)) {
            $reason = 'MISSING'
        }
        elseif ($Force) {
            $reason = 'FORCED'
        }
        else {
            $dstItem = Get-Item -LiteralPath $dst -ErrorAction SilentlyContinue
            if ($null -eq $dstItem -or $dstItem.Length -ne $src.Length) {
                $reason = 'SIZE-MISMATCH'
            }
            else {
                try {
                    $srcHash = (Get-FileHash -LiteralPath $src.FullName -Algorithm SHA256 -ErrorAction Stop).Hash
                    $dstHash = (Get-FileHash -LiteralPath $dst          -Algorithm SHA256 -ErrorAction Stop).Hash
                    if ($srcHash -ne $dstHash) { $reason = 'HASH-MISMATCH' }
                }
                catch { $reason = 'UNREADABLE' }
            }
        }

        if (-not $reason) { $verifiedOk++; continue }

        try {
            Copy-Item -LiteralPath $src.FullName -Destination $dst -Force -ErrorAction Stop
            if ($reason -eq 'MISSING') { $copiedMissing++ } else { $copiedChanged++ }
            Write-Log "Copied '$($src.Name)' (reason=$reason)."
        }
        catch {
            $failed++
            Write-Log "Failed to copy '$($src.Name)': $_" 'ERROR'
        }
    }

    if ($copiedMissing -eq 0 -and $copiedChanged -eq 0 -and $failed -eq 0) {
        Write-Log "Verified $verifiedOk file(s) already present and correct - no action needed."
    }
    else {
        Write-Log ("Converged: {0} restored (missing), {1} refreshed (changed), {2} already correct, {3} failed." -f `
            $copiedMissing, $copiedChanged, $verifiedOk, $failed)
        if ($copiedMissing -gt 0 -or $copiedChanged -gt 0) {
            Write-Log 'Backgrounds appear in the picker the next time Teams restarts. No user action required.'
        }
    }

    try {
        [pscustomobject]@{
            LastSuccessDate = (Get-Date).ToString('yyyy-MM-dd')
            LastRunUtc   = (Get-Date).ToUniversalTime().ToString('o')
            PayloadCount = $staged.Count
            Verified     = $verifiedOk
            Restored     = $copiedMissing
            Refreshed    = $copiedChanged
            Failed       = $failed
        } | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath $UserMarkerFile -Encoding UTF8
    }
    catch { }

    Write-Log '=== Teams background APPLY finished ==='
    exit 0
}

# ------------------------------------------------------------------------------------------------
# UNINSTALL
# ------------------------------------------------------------------------------------------------
function Invoke-Uninstall {
    Set-LogContext -Context Machine
    Write-Log '=== Teams background removal (logon variant) starting ==='

    Unregister-LogonTask

    $removedAnything = $false

    if (Test-Path -LiteralPath $PayloadDir) {
        $guidFiles = @()
        $guidFiles += @(Get-ChildItem -LiteralPath $PayloadDir -File -ErrorAction SilentlyContinue |
                        Select-Object -ExpandProperty Name)

        if (Test-Path -LiteralPath $ManifestFile) {
            try {
                $manifest = @(Get-Content -LiteralPath $ManifestFile -Raw | ConvertFrom-Json)
                $guidFiles += @($manifest | ForEach-Object { $_.MainFile; $_.ThumbFile })
            }
            catch { Write-Log "Manifest unreadable - relying on payload folder names only." 'WARN' }
        }

        $guidFiles = @($guidFiles | Where-Object { $_ -is [string] -and $_ } | Select-Object -Unique)

        $uploadDirs = @()
        if (Test-IsSystem) {
            $excludedSids = '^S-1-5-18$|^S-1-5-19$|^S-1-5-20$'
            Get-CimInstance -ClassName Win32_UserProfile -ErrorAction SilentlyContinue |
                Where-Object { $_.Special -eq $false -and $_.SID -notmatch $excludedSids -and $_.LocalPath } |
                ForEach-Object { $uploadDirs += (Join-Path (Join-Path $_.LocalPath 'AppData\Local') $TeamsUploadsRel) }
        }
        else {
            $uploadDirs += (Join-Path $env:LOCALAPPDATA $TeamsUploadsRel)
        }

        foreach ($dir in ($uploadDirs | Select-Object -Unique)) {
            foreach ($fn in $guidFiles) {
                $f = Join-Path $dir $fn
                if (Test-Path -LiteralPath $f) {
                    Remove-Item -LiteralPath $f -Force -ErrorAction SilentlyContinue
                    Write-Log "Removed $f"
                    $removedAnything = $true
                }
            }
        }
    }
    else {
        Write-Log 'No payload folder found - nothing tracked to remove from Uploads folders.' 'WARN'
    }

    Remove-Item -LiteralPath $RootDir        -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $WorkDir        -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $UserMarkerFile -Force -ErrorAction SilentlyContinue

    if (-not $removedAnything) { Write-Log 'Nothing was present to remove.' 'WARN'; exit 14 }
    Write-Log '=== Teams background removal completed ==='
    exit 0
}

# ------------------------------------------------------------------------------------------------
# Entry point
# ------------------------------------------------------------------------------------------------
try {
    if ($Uninstall)            { Invoke-Uninstall }
    elseif ($Mode -eq 'Apply') { Invoke-Apply }
    else                       { Invoke-Stage }
}
catch {
    $ctx = if ($Mode -eq 'Apply') { 'User' } else { 'Machine' }
    try { Set-LogContext -Context $ctx } catch { }
    Write-Log "UNHANDLED EXCEPTION: $_" 'ERROR'
    exit 99
}
