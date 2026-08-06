<#
.SYNOPSIS
    Deploys (or removes) custom Microsoft Teams (new client) background images, applying them to
    each user only once that user has actually launched and signed into new Teams — and keeping
    watch so that users who sign in later still get them. Intune/RMM-ready.

.DESCRIPTION
    Two-mode Stage + Apply architecture:

      -Mode Stage  (default — run once by Intune device-context / RMM as SYSTEM)
          1. Resolves source images (-ManifestUrl > -ImageUrls > -ImageFolder > -ZipUrl > auto Images folder).
          2. Converts each to a GUID-named 1920x1080 JPEG + 280x158 thumbnail.
          3. Writes finished pairs to a machine-wide payload folder:
               C:\ProgramData\TeamsBG\Payload
          4. Copies this script to a stable path, registers a scheduled task
             (ApplyTeamsBackgrounds) that runs -Mode Apply in the user's context
             at logon and every N minutes, then kicks the task immediately.

      -Mode Apply  (run by the scheduled task in the logged-on user's context)
          1. Checks whether new Teams has been initialised for this user
             (EBWebView profile tree exists under LocalCache\Microsoft\MSTeams).
             If not, exits 0 — the task retries on the next tick.
          2. Copies pre-built pairs from the machine payload into the user's
               %LocalAppData%\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\Backgrounds\Uploads
             On every run it compares by size then SHA256 and restores anything missing or changed
             (convergence — no "already done" flag that can claim success while files are absent).

    -Uninstall removes the task, the machine payload, and every GUID-named file this tool
    deployed from every user profile, without touching the user's own uploads.

    Logs:
      Machine (Stage/Uninstall): C:\ProgramData\TeamsBG\Logs\teamsbackground.txt
      Per-user (Apply):          %LocalAppData%\TeamsBG\Logs\teamsbackground-apply.txt

.PARAMETER Mode
    Stage (default) or Apply. You invoke Stage; the scheduled task invokes Apply.

.PARAMETER ZipUrl
    Public URL to a .zip of images. Downloaded, extracted, every .jpg/.jpeg/.png used.
    Leave empty to use a local Images\ folder alongside the script.

.PARAMETER ManifestUrl
    Public URL to a JSON array: [ { "Name": "...", "Url": "https://.../image.jpg" }, ... ]

.PARAMETER ImageUrls
    Direct list of public image URLs.

.PARAMETER ImageNames
    Optional friendly names matched positionally to -ImageUrls.

.PARAMETER ImageFolder
    Explicit local folder of images (use when images are packaged with the Intune app).

.PARAMETER ApplyIntervalMinutes
    How often the per-user task re-checks. Default 30.

.PARAMETER Force
    Re-stage / re-apply even when content is unchanged.

.PARAMETER Uninstall
    Remove the task, payload, and all deployed files; clear state.

.EXAMPLE
    # Intune device-context / RMM (SYSTEM). Stages payload + registers the apply task.
    powershell.exe -ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File Deploy-TeamsBackgrounds.ps1

.EXAMPLE
    # Point at a different zip without repackaging the Intune app.
    powershell.exe -ExecutionPolicy Bypass -NoProfile -File Deploy-TeamsBackgrounds.ps1 -ZipUrl "https://yourstorage.blob.core.windows.net/container/backgrounds-v2.zip"

.EXAMPLE
    # Uninstall
    powershell.exe -ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File Deploy-TeamsBackgrounds.ps1 -Uninstall

.NOTES
    Exit codes:
      0   Success (Stage completed; or Apply completed / Teams not yet signed in)
      10  Download or zip-extraction failed
      11  No images supplied / source produced zero usable entries
      12  Payload staging produced zero usable images
      13  One or more images failed to process during Stage
      14  -Uninstall requested but nothing to remove
      20  -Mode Apply but no staged payload/manifest found (Stage has not run)
      99  Unhandled exception
#>

[CmdletBinding(DefaultParameterSetName = 'Install')]
param(
    [Parameter(ParameterSetName = 'Install')]
    [ValidateSet('Stage', 'Apply')]
    [string]$Mode = 'Stage',

    [Parameter(ParameterSetName = 'Install')]
    [string]$ZipUrl = '',   # <-- Set your blob storage URL here, or leave empty and use an Images\ folder

    [Parameter(ParameterSetName = 'Install')]
    [string]$ManifestUrl,

    [Parameter(ParameterSetName = 'Install')]
    [string[]]$ImageUrls,

    [Parameter(ParameterSetName = 'Install')]
    [string[]]$ImageNames,

    [Parameter(ParameterSetName = 'Install')]
    [string]$ImageFolder,

    [Parameter(ParameterSetName = 'Install')]
    [ValidateRange(5, 1440)]
    [int]$ApplyIntervalMinutes = 30,

    [Parameter(ParameterSetName = 'Install')]
    [switch]$Force,

    [Parameter(ParameterSetName = 'Uninstall', Mandatory = $true)]
    [switch]$Uninstall
)

# ------------------------------------------------------------------------------------------------
# Constants / paths  (change TeamsBG prefix and ApplyTeamsBackgrounds task name to fit your org)
# ------------------------------------------------------------------------------------------------
$CompanyRoot     = 'C:\ProgramData\TeamsBG'
$LogDir          = Join-Path $CompanyRoot 'Logs'
$LogFile         = Join-Path $LogDir 'teamsbackground.txt'
$RootDir         = Join-Path $CompanyRoot 'TeamsBackgrounds'
$PayloadDir      = Join-Path $RootDir 'Payload'
$ManifestFile    = Join-Path $RootDir 'payload.json'
$SelfInstallPath = Join-Path $RootDir 'Deploy-TeamsBackgrounds.ps1'
$WorkDir         = Join-Path $env:TEMP 'TeamsBG-Work'
$ExtractDir      = Join-Path $WorkDir 'Extracted'

$TaskName        = 'ApplyTeamsBackgrounds'

$UserRoot        = Join-Path $env:LOCALAPPDATA 'TeamsBG'
$UserLogDir      = Join-Path $UserRoot 'Logs'
$UserLogFile     = Join-Path $UserLogDir 'teamsbackground-apply.txt'
$UserMarkerFile  = Join-Path $UserRoot 'teams-bg-applied.json'

$TeamsWebViewRel = 'Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\EBWebView'
$TeamsUploadsRel = 'Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\Backgrounds\Uploads'

$MainW  = 1920;  $MainH  = 1080
$ThumbW = 280;   $ThumbH = 158

$script:ActiveLogFile = $LogFile

# ------------------------------------------------------------------------------------------------
# Logging / ACLs
# ------------------------------------------------------------------------------------------------
function Repair-AclPermissions {
    param([Parameter(Mandatory)][string]$Path)
    try {
        if (Test-Path -LiteralPath $Path) {
            & takeown.exe /F $Path /A /R /D Y 2>&1 | Out-Null
            & icacls.exe $Path /grant '*S-1-5-32-544:(OI)(CI)F' /grant '*S-1-5-18:(OI)(CI)F' /grant '*S-1-5-32-545:(OI)(CI)M' /T /C /Q 2>&1 | Out-Null
        }
    } catch { }
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
    } else {
        foreach ($d in @($LogDir, $RootDir, $PayloadDir, $WorkDir)) {
            if (-not (Test-Path -LiteralPath $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
        }
        Repair-AclPermissions -Path $CompanyRoot
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
    } catch {
        try {
            if (Test-IsSystem) { Repair-AclPermissions -Path $script:ActiveLogFile }
            Add-Content -LiteralPath $script:ActiveLogFile -Value $line -Encoding UTF8
        } catch { }
    }
    if ($Level -eq 'ERROR') { Write-Host $line -ForegroundColor Red }
    elseif ($Level -eq 'WARN') { Write-Host $line -ForegroundColor Yellow }
    else { Write-Host $line }
}

# ------------------------------------------------------------------------------------------------
# Helpers
# ------------------------------------------------------------------------------------------------
function Get-StableHash {
    param([Parameter(Mandatory)][string]$Text)
    $sha   = [System.Security.Cryptography.SHA256]::Create()
    $bytes = [Text.Encoding]::UTF8.GetBytes($Text)
    $hash  = $sha.ComputeHash($bytes)
    $hex   = -join ($hash | ForEach-Object { $_.ToString('x2') })
    return $hex.Substring(0, 16)
}

function Get-FileContentHash {
    param([Parameter(Mandatory)][string]$Path)
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.Substring(0, 16)
}

function New-DeterministicGuid {
    # Derives a stable, valid GUID from a seed string. Same seed always produces the same GUID
    # with no stored state — prevents filename accumulation across re-runs.
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
        } catch {
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
        } finally { $gfx.Dispose() }

        $jpegCodec = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.MimeType -eq 'image/jpeg' }
        $encParams = New-Object System.Drawing.Imaging.EncoderParameters -ArgumentList @(1)
        $encParams.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter -ArgumentList @([System.Drawing.Imaging.Encoder]::Quality, [long]90)
        $bmp.Save($DestPath, $jpegCodec, $encParams)
        $bmp.Dispose()
    } finally { $src.Dispose() }
}

function Get-ImageList {
    if ($ManifestUrl) {
        Write-Log "Downloading manifest from $ManifestUrl"
        try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; $raw = Invoke-RestMethod -Uri $ManifestUrl -UseBasicParsing -TimeoutSec 30 }
        catch { Write-Log "Failed to download manifest: $_" 'ERROR'; exit 10 }
        $list = @()
        foreach ($item in @($raw)) {
            if (-not $item.Url) { continue }
            $name = if ($item.Name) { $item.Name } else { [IO.Path]::GetFileNameWithoutExtension($item.Url) }
            $list += [pscustomobject]@{ Name = $name; Source = $item.Url; IsLocal = $false }
        }
        return , $list
    } elseif ($ImageUrls) {
        $list = @()
        for ($i = 0; $i -lt $ImageUrls.Count; $i++) {
            $name = if ($ImageNames -and $i -lt $ImageNames.Count) { $ImageNames[$i] } else { [IO.Path]::GetFileNameWithoutExtension($ImageUrls[$i]) }
            $list += [pscustomobject]@{ Name = $name; Source = $ImageUrls[$i]; IsLocal = $false }
        }
        return , $list
    } elseif ($ImageFolder) {
        if (-not (Test-Path -LiteralPath $ImageFolder)) { Write-Log "Specified -ImageFolder not found: $ImageFolder" 'ERROR'; exit 11 }
        $files = Get-ChildItem -LiteralPath $ImageFolder -File -Recurse -Include '*.jpg','*.jpeg','*.png' -ErrorAction SilentlyContinue
        if (-not $files -or $files.Count -eq 0) { Write-Log "Folder '$ImageFolder' contained no images." 'ERROR'; exit 11 }
        $list = @(); foreach ($f in $files) { $list += [pscustomobject]@{ Name = $f.BaseName; Source = $f.FullName; IsLocal = $true } }
        return , $list
    } elseif ($ZipUrl) {
        Write-Log "Downloading images zip from $ZipUrl"
        $zipTemp = Join-Path $WorkDir 'images-source.zip'
        if (Test-Path -LiteralPath $ExtractDir) { Remove-Item -LiteralPath $ExtractDir -Recurse -Force -ErrorAction SilentlyContinue }
        New-Item -ItemType Directory -Path $ExtractDir -Force | Out-Null
        if (-not (Invoke-DownloadWithRetry -Uri $ZipUrl -OutFile $zipTemp -TimeoutSec 300)) { Write-Log "Failed to download zip after 3 attempts." 'ERROR'; exit 10 }
        Write-Log "Zip downloaded. Extracting..."
        try { Expand-Archive -LiteralPath $zipTemp -DestinationPath $ExtractDir -Force }
        catch { Write-Log "Failed to extract zip: $_" 'ERROR'; exit 10 }
        Remove-Item -LiteralPath $zipTemp -Force -ErrorAction SilentlyContinue
        $files = Get-ChildItem -LiteralPath $ExtractDir -File -Recurse -Include '*.jpg','*.jpeg','*.png' -ErrorAction SilentlyContinue
        if (-not $files -or $files.Count -eq 0) { Write-Log 'Zip extracted but contained no images.' 'ERROR'; exit 11 }
        Write-Log "Extracted $($files.Count) image file(s)."
        $list = @(); foreach ($f in $files) { $list += [pscustomobject]@{ Name = $f.BaseName; Source = $f.FullName; IsLocal = $true } }
        return , $list
    } else {
        $folder = Join-Path $PSScriptRoot 'Images'
        if (-not (Test-Path -LiteralPath $folder)) { Write-Log "No source supplied and no 'Images' folder next to the script." 'ERROR'; exit 11 }
        $files = Get-ChildItem -LiteralPath $folder -File -Recurse -Include '*.jpg','*.jpeg','*.png' -ErrorAction SilentlyContinue
        if (-not $files -or $files.Count -eq 0) { Write-Log "Folder '$folder' contained no images." 'ERROR'; exit 11 }
        $list = @(); foreach ($f in $files) { $list += [pscustomobject]@{ Name = $f.BaseName; Source = $f.FullName; IsLocal = $true } }
        return , $list
    }
}

# ------------------------------------------------------------------------------------------------
# New Teams detection
# ------------------------------------------------------------------------------------------------
function Test-TeamsInitialized {
    # Returns true only once new Teams has actually launched for this user. The WebView2 EBWebView
    # profile tree is created on first launch. A machine-level script that merely created the
    # Backgrounds\Uploads folder does NOT create EBWebView, so this cannot be faked by the
    # folder-creation side effect — which is why we key off EBWebView, not the bare MSTeams folder.
    param([Parameter(Mandatory)][string]$LocalAppData)
    $webview = Join-Path $LocalAppData $TeamsWebViewRel
    if (-not (Test-Path -LiteralPath $webview)) { return $false }
    $hasContent = Get-ChildItem -LiteralPath $webview -Force -ErrorAction SilentlyContinue | Select-Object -First 1
    return [bool]$hasContent
}

# ------------------------------------------------------------------------------------------------
# Scheduled task registration
# ------------------------------------------------------------------------------------------------
function Register-ApplyTask {
    # Registers two INDEPENDENT triggers via raw task XML:
    #   <TimeTrigger>  — owns the repetition, starts immediately, ticks on the wall clock
    #   <LogonTrigger> — catches future sign-ins promptly
    #
    # Why raw XML? A <Repetition> nested inside a <LogonTrigger> (the PowerShell API default) is
    # a modifier on that trigger, not a standalone schedule. It only counts from the next logon
    # event, so on any device where Stage runs while a user is already signed in the task never
    # runs until the next reboot. Raw XML is the only registration path that correctly separates
    # the two triggers across all Windows builds.
    try {
        $startBoundary = (Get-Date).AddMinutes(-1).ToString('yyyy-MM-ddTHH:mm:ss')
        $interval      = "PT${ApplyIntervalMinutes}M"
        $arguments     = '-ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File "{0}" -Mode Apply' -f $SelfInstallPath
        $argumentsEscaped = [System.Security.SecurityElement]::Escape($arguments)

        $taskXml = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.3" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Description>Applies staged Teams background images in the logged-on user context, at logon and every $ApplyIntervalMinutes minutes. Converges on every run.</Description>
    <URI>$TaskName</URI>
  </RegistrationInfo>
  <Triggers>
    <LogonTrigger><Enabled>true</Enabled></LogonTrigger>
    <TimeTrigger>
      <StartBoundary>$startBoundary</StartBoundary>
      <Enabled>true</Enabled>
      <Repetition>
        <Interval>$interval</Interval>
        <StopAtDurationEnd>false</StopAtDurationEnd>
      </Repetition>
    </TimeTrigger>
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
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <ExecutionTimeLimit>PT15M</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>powershell.exe</Command>
      <Arguments>$argumentsEscaped</Arguments>
    </Exec>
  </Actions>
</Task>
"@
        Register-ScheduledTask -TaskName $TaskName -Xml $taskXml -Force | Out-Null

        # Verify — registration can appear to succeed while the trigger silently failed.
        $verify = Export-ScheduledTask -TaskName $TaskName
        if ($verify -notmatch '<TimeTrigger>') {
            Write-Log "Task registered but no <TimeTrigger> found — it would only run at logon." 'ERROR'
            return $false
        }
        if ($verify -notmatch [regex]::Escape($interval)) {
            Write-Log "Task registered but the $interval repetition did not persist." 'ERROR'
            return $false
        }

        $next = (Get-ScheduledTaskInfo -TaskName $TaskName -ErrorAction SilentlyContinue).NextRunTime
        if ($next) { Write-Log "Registered task '$TaskName' (logon + every $ApplyIntervalMinutes min, user context). Next run: $next" }
        else { Write-Log "Task '$TaskName' registered but NextRunTime is empty — investigate on this device." 'WARN' }
        return $true
    } catch {
        Write-Log "Failed to register task '$TaskName': $_" 'ERROR'
        return $false
    }
}

function Unregister-ApplyTask {
    try {
        if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
            Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
            Write-Log "Removed task '$TaskName'."
        }
    } catch { Write-Log "Failed to remove task '$TaskName': $_" 'WARN' }
}

# ------------------------------------------------------------------------------------------------
# STAGE
# ------------------------------------------------------------------------------------------------
function Invoke-Stage {
    Set-LogContext -Context Machine
    Write-Log '=== Teams background STAGE starting ==='

    if (Test-Path -LiteralPath $PayloadDir) {
        Get-ChildItem -LiteralPath $PayloadDir -File -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
    } else {
        New-Item -ItemType Directory -Path $PayloadDir -Force | Out-Null
    }

    $images = Get-ImageList
    if (-not $images -or $images.Count -eq 0) { Write-Log 'Source produced zero usable images.' 'ERROR'; exit 11 }

    $manifest  = @()
    $failures  = 0
    $seenGuids = @{}

    foreach ($img in $images) {
        try {
            $cacheKey = if ($img.IsLocal) { Get-FileContentHash -Path $img.Source } else { Get-StableHash -Text $img.Source }
            $guid = New-DeterministicGuid -Seed $cacheKey

            if ($seenGuids.ContainsKey($guid)) {
                Write-Log "Skipping duplicate '$($img.Name)' (identical content to an earlier image)."
                continue
            }
            $seenGuids[$guid] = $true

            $mainFile  = "$guid.jpg"
            $thumbFile = "${guid}_thumb.jpg"
            $mainOut   = Join-Path $PayloadDir $mainFile
            $thumbOut  = Join-Path $PayloadDir $thumbFile
            $srcTemp   = Join-Path $WorkDir "$guid.src"

            if ($img.IsLocal) {
                Copy-Item -LiteralPath $img.Source -Destination $srcTemp -Force
            } else {
                Write-Log "Downloading '$($img.Name)' from $($img.Source)"
                if (-not (Invoke-DownloadWithRetry -Uri $img.Source -OutFile $srcTemp -TimeoutSec 60)) {
                    throw "Failed to download: $($img.Source)"
                }
            }

            Write-Log "Building payload for '$($img.Name)' -> $mainFile"
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
        } catch {
            Write-Log "FAILED staging '$($img.Name)': $_" 'ERROR'
            $failures++
        }
    }

    if (-not $manifest -or $manifest.Count -eq 0) { Write-Log 'Staging produced zero usable images.' 'ERROR'; exit 12 }

    ConvertTo-Json -InputObject @($manifest) -Depth 5 | Set-Content -LiteralPath $ManifestFile -Encoding UTF8
    Write-Log "Manifest written: $($manifest.Count) image(s)."

    try {
        Copy-Item -LiteralPath $PSCommandPath -Destination $SelfInstallPath -Force
        Write-Log "Script copied to $SelfInstallPath"
    } catch { Write-Log "Could not copy script to $SelfInstallPath : $_" 'ERROR' }

    Repair-AclPermissions -Path $RootDir
    $null = Register-ApplyTask
    Remove-Item -LiteralPath $ExtractDir -Recurse -Force -ErrorAction SilentlyContinue

    try { Start-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue; Write-Log "Kicked task for current sessions." } catch { }

    if ($failures -gt 0) { Write-Log "Stage completed with $failures failure(s)." 'ERROR'; exit 13 }
    Write-Log '=== Teams background STAGE completed successfully ==='
    exit 0
}

# ------------------------------------------------------------------------------------------------
# APPLY
# ------------------------------------------------------------------------------------------------
function Invoke-Apply {
    Set-LogContext -Context User
    Write-Log "=== Teams background APPLY starting (user $env:USERNAME) ==="

    if (Test-IsSystem) {
        Write-Log 'Apply running as SYSTEM — LOCALAPPDATA is not a real user profile. Nothing to do.' 'WARN'
        exit 0
    }

    if (-not (Test-Path -LiteralPath $ManifestFile)) {
        Write-Log "No staged manifest at $ManifestFile — Stage has not run." 'WARN'
        exit 20
    }

    $staged = @(Get-ChildItem -LiteralPath $PayloadDir -File -ErrorAction SilentlyContinue |
                Where-Object { $_.Extension -match '^.(jpg|jpeg|png)$' })

    if ($staged.Count -eq 0) { Write-Log "Payload folder contains no images — Stage did not complete." 'ERROR'; exit 20 }
    Write-Log "Payload holds $($staged.Count) file(s)."

    if (-not (Test-TeamsInitialized -LocalAppData $env:LOCALAPPDATA)) {
        Write-Log 'New Teams not yet launched/signed in for this user - will re-check on the next run.'
        exit 0
    }

    $teamsRunning = [bool](Get-Process -Name 'ms-teams' -ErrorAction SilentlyContinue)
    if ($teamsRunning) { Write-Log 'Teams is currently RUNNING (copying anyway — convergence handles any wipe).' }
    else               { Write-Log 'Teams is currently closed.' }

    $uploads = Join-Path $env:LOCALAPPDATA $TeamsUploadsRel
    try {
        if (-not (Test-Path -LiteralPath $uploads)) {
            New-Item -ItemType Directory -Path $uploads -Force -ErrorAction Stop | Out-Null
            Write-Log 'Created Uploads folder.'
        }
    } catch { Write-Log "Could not create Uploads folder '$uploads': $_" 'ERROR'; exit 0 }

    $copiedMissing = 0; $copiedChanged = 0; $verifiedOk = 0; $failed = 0

    foreach ($src in $staged) {
        $dst    = Join-Path $uploads $src.Name
        $reason = $null

        if (-not (Test-Path -LiteralPath $dst)) {
            $reason = 'MISSING'
        } elseif ($Force) {
            $reason = 'FORCED'
        } else {
            $dstItem = Get-Item -LiteralPath $dst -ErrorAction SilentlyContinue
            if ($null -eq $dstItem -or $dstItem.Length -ne $src.Length) {
                $reason = 'SIZE-MISMATCH'
            } else {
                try {
                    $srcHash = (Get-FileHash -LiteralPath $src.FullName -Algorithm SHA256).Hash
                    $dstHash = (Get-FileHash -LiteralPath $dst          -Algorithm SHA256).Hash
                    if ($srcHash -ne $dstHash) { $reason = 'HASH-MISMATCH' }
                } catch { $reason = 'UNREADABLE' }
            }
        }

        if (-not $reason) { $verifiedOk++; continue }

        try {
            Copy-Item -LiteralPath $src.FullName -Destination $dst -Force -ErrorAction Stop
            if ($reason -eq 'MISSING') { $copiedMissing++ } else { $copiedChanged++ }
            Write-Log "Copied '$($src.Name)' (reason=$reason)."
        } catch {
            $failed++
            Write-Log "Failed to copy '$($src.Name)': $_" 'ERROR'
        }
    }

    if ($copiedMissing -eq 0 -and $copiedChanged -eq 0 -and $failed -eq 0) {
        Write-Log "Verified $verifiedOk file(s) already present and correct — no action needed."
    } else {
        $summary = "Converged: {0} restored (missing), {1} refreshed (changed), {2} already correct, {3} failed." -f $copiedMissing, $copiedChanged, $verifiedOk, $failed
        Write-Log $summary
        if ($copiedMissing -gt 0 -and $teamsRunning) {
            Write-Log 'NOTE: files were missing while Teams was running. If this recurs every tick, Teams is clearing the Uploads cache.' 'WARN'
        }
        if ($copiedMissing -gt 0 -or $copiedChanged -gt 0) {
            Write-Log 'Backgrounds appear in the picker the next time Teams restarts. No user action required.'
        }
    }

    try {
        [pscustomobject]@{
            LastRunUtc   = (Get-Date).ToUniversalTime().ToString('o')
            PayloadCount = $staged.Count
            Verified     = $verifiedOk
            Restored     = $copiedMissing
            Refreshed    = $copiedChanged
            Failed       = $failed
        } | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath $UserMarkerFile -Encoding UTF8
    } catch { }

    Write-Log '=== Teams background APPLY finished ==='
    exit 0
}

# ------------------------------------------------------------------------------------------------
# UNINSTALL
# ------------------------------------------------------------------------------------------------
function Invoke-Uninstall {
    Set-LogContext -Context Machine
    Write-Log '=== Teams background removal starting ==='
    Unregister-ApplyTask
    $removedAnything = $false

    if (Test-Path -LiteralPath $ManifestFile) {
        $guidFiles = @()
        if (Test-Path -LiteralPath $PayloadDir) {
            $guidFiles += @(Get-ChildItem -LiteralPath $PayloadDir -File | Select-Object -ExpandProperty Name)
        }
        $manifest   = @(Get-Content -LiteralPath $ManifestFile -Raw | ConvertFrom-Json)
        $guidFiles += @($manifest | ForEach-Object { $_.MainFile; $_.ThumbFile })
        $guidFiles   = @($guidFiles | Where-Object { $_ -is [string] -and $_ } | Select-Object -Unique)

        $uploadDirs = @()
        if (Test-IsSystem) {
            Get-CimInstance -ClassName Win32_UserProfile -ErrorAction SilentlyContinue |
                Where-Object { -not $_.Special -and $_.LocalPath } |
                ForEach-Object { $uploadDirs += Join-Path $_.LocalPath $TeamsUploadsRel }
        } else {
            $uploadDirs += Join-Path $env:LOCALAPPDATA $TeamsUploadsRel
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
    } else {
        Write-Log 'No manifest found — nothing tracked to remove.' 'WARN'
    }

    Remove-Item -LiteralPath $RootDir         -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $WorkDir         -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $UserMarkerFile  -Force   -ErrorAction SilentlyContinue

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
} catch {
    $ctx = if ($Mode -eq 'Apply') { 'User' } else { 'Machine' }
    try { Set-LogContext -Context $ctx } catch { }
    Write-Log "UNHANDLED EXCEPTION: $_" 'ERROR'
    exit 99
}