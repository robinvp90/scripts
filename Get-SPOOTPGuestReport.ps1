<#
.SYNOPSIS
    Reports SharePoint-only (SPO OTP) guest users across SharePoint and OneDrive sites and
    flags which of them have no matching Microsoft Entra B2B guest account.

.DESCRIPTION
    Context: Microsoft is moving external user authentication in SharePoint and OneDrive
    from SharePoint Online to Microsoft Entra B2B (Message Center MC1243549).
    One-time passcode itself is NOT being retired - Entra B2B uses OTP by default. What
    goes away is the SharePoint-only guest: an external user who was shared content via
    SPO OTP and never got an Entra B2B guest account. Once the retirement reaches your
    tenant, those users get access denied on previously shared links.

    A guest who already has an Entra B2B account keeps access to everything previously
    shared with them, so the fix is to find the ones who don't. Microsoft's documented
    way to do that is the site-level external sharing report, one site at a time. This
    script does it across every SharePoint site and OneDrive in the tenant.

    Detection logic:
      * SPO-only OTP guest  -> site login name contains 'urn:spo:guest' (or encoded 'spo%3aguest')
      * Entra B2B guest     -> site login name contains '#EXT#'

    SETUP: fully automated, one sign-in, no manual steps.
    Register-PnPEntraIDApp is NOT used - it runs its own sign-in and consent flows that
    prompt twice, print device codes unpredictably and hang. Instead this script does the
    whole setup itself against Graph REST with a single device-code sign-in:
        1. verifies the tenant you signed into is the tenant you asked for
        2. finds an existing app registration by name, or creates one
        3. reuses a matching local certificate, or creates and uploads one
        4. creates the service principal
        5. grants tenant-wide admin consent via appRoleAssignments
        6. caches everything, then retries the connection while consent propagates
    Re-running is safe: every step is skipped if already done.

    On later runs, if a saved setup exists for the tenant, the script shows it and asks
    whether to use it or run setup again. -UseCachedSetup skips the question and uses the
    saved setup; -Reset skips it and runs setup again.

    PERMISSIONS (application, admin-consented):
      * Microsoft Graph  User.Read.All          - read guest accounts
      * SharePoint       Sites.FullControl.All  - app-only access to the SharePoint
                                                  tenant-admin APIs, which tenant-wide
                                                  site enumeration needs
    Sites.FullControl.All is powerful. The report itself only reads, but delete the app
    registration when your audit is finished - see CLEANUP below.

    KNOWN LIMITATION:
    Only guests who have signed in to the site at least once appear in the site User
    Information List. A recipient sent a link who never opened it will NOT appear. Use the
    site-level external sharing report in the SharePoint admin centre as a complement.

    MODULE NOTE:
    PnP.PowerShell ships its own Microsoft.Graph.Core 1.25.1 assembly, which breaks the
    Microsoft.Graph SDK in the same session (pnp/powershell issue 3395). This script uses
    PnP only, and talks to Graph over raw REST or Invoke-PnPGraphMethod.

    DEVICE CODE FLOW:
    Setup signs in with the OAuth device code flow. If a Conditional Access policy blocks
    device code flow in your tenant, create the app registration yourself and pass
    -ClientId and -CertificateThumbprint instead.

    PRIVACY:
    The CSV contains external users' names and email addresses. Treat it as confidential.
    The cached SPO-OTP-Audit.<tenant>.config.json file identifies your tenant and app
    registration; keep it out of source control.

    CLEANUP (when the audit is done):
      1. Entra admin centre > App registrations > delete the app (its name is printed at
         start-up, by default SPO-OTP-Audit-<user>-<computer>).
      2. certmgr.msc > Personal > Certificates > delete the certificate with the same name.
      3. Delete the SPO-OTP-Audit.<tenant>.config.json file next to the script.

.NOTES
    Author  : Robin Poulose - https://robztech.com
    Version : 3.10
    Setup   : CREATES an Entra app, certificate and consent (idempotent).
    Report  : READ ONLY.

    Required: PowerShell 7.2+ on Windows, PnP.PowerShell, and a Global Administrator
    or Privileged Role Administrator for the one-time setup sign-in.

    Provided as-is, without warranty. Review it and test in a non-production tenant first.
    MIT License.

.EXAMPLE
    .\Get-SPOOTPGuestReport.ps1

.EXAMPLE
    .\Get-SPOOTPGuestReport.ps1 -TenantName contoso -TenantDomain contoso.onmicrosoft.com

.EXAMPLE
    .\Get-SPOOTPGuestReport.ps1 -SiteUrlFilter '*/sites/Finance*' -IncludeOneDrive:$false

.EXAMPLE
    .\Get-SPOOTPGuestReport.ps1 -IncludeSharedItems

    Also lists the files and folders each external user holds a sharing link to. Slower.

.EXAMPLE
    .\Get-SPOOTPGuestReport.ps1 -TenantName contoso -TenantDomain contoso.onmicrosoft.com -UseCachedSetup

    Unattended run: uses the saved setup without asking.

.EXAMPLE
    .\Get-SPOOTPGuestReport.ps1 -Reset

    Runs setup again without asking, re-checking the app, certificate and consent.
#>

[CmdletBinding()]
param(
    [string]$TenantName,
    [string]$TenantDomain,
    [string]$ClientId,
    [string]$CertificateThumbprint,

    # Name for the app registration. Default includes user and machine because the
    # certificate lives in one user's store on one machine.
    [string]$AppName,
    [string]$AppNameSuffix,

    # Re-run setup even if a saved, working configuration exists
    [switch]$Reset,

    # Use a saved setup without asking. Without this, the script asks whether to use
    # the saved setup or run setup again. Use it for unattended or scheduled runs.
    [switch]$UseCachedSetup,

    # Certificate lifetime when this script creates one
    [int]$CertYears = 2,

    [string]$ConfigPath,
    [bool]$IncludeOneDrive = $true,
    [string]$SiteUrlFilter,

    # Also resolve WHICH files/folders each external user holds a sharing link to.
    # Slower - adds one pass over each site's sharing-link groups.
    [switch]$IncludeSharedItems,

    # Cap on how many item paths are listed per user per site
    [int]$MaxItemsPerUser = 25,

    # Cap on how many items are paged per document library when resolving paths
    [int]$MaxItemsPerList = 20000,
    [string]$OutputPath = ".\SPO-OTP-GuestReport_$(Get-Date -Format 'yyyyMMdd-HHmm').csv"
)

$ErrorActionPreference = 'Stop'
$ScriptVersion = '3.10'

# Permissions this app needs, by resource app id. Least privilege: the report reads guest
# accounts and enumerates sites; it never invites or modifies users.
$RequiredPermissions = @{
    '00000003-0000-0000-c000-000000000000' = @('User.Read.All')          # Microsoft Graph
    '00000003-0000-0ff1-ce00-000000000000' = @('Sites.FullControl.All')  # SharePoint
}

if (-not $AppName) {
    if (-not $AppNameSuffix) {
        $who   = ($env:USERNAME     -replace '[^A-Za-z0-9\-]', '')
        $where = ($env:COMPUTERNAME -replace '[^A-Za-z0-9\-]', '')
        $AppNameSuffix = (@($who, $where) | Where-Object { $_ }) -join '-'
    }
    $AppName = if ($AppNameSuffix) { "SPO-OTP-Audit-$AppNameSuffix" } else { 'SPO-OTP-Audit' }
}

Write-Host ""
Write-Host "  SPO OTP Guest Report  v$ScriptVersion" -ForegroundColor White -BackgroundColor DarkBlue
Write-Host ("  PowerShell {0} ({1})" -f $PSVersionTable.PSVersion, $PSVersionTable.PSEdition) -ForegroundColor DarkGray
Write-Host ("  App name   {0}" -f $AppName) -ForegroundColor DarkGray
Write-Host ""

# ===========================================================================
# Preflight
# ===========================================================================
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "PowerShell 7.2 or later is required (PnP.PowerShell dropped 5.1 and the ISE)." -ForegroundColor Red
    Write-Host "  winget install --id Microsoft.PowerShell --source winget" -ForegroundColor Yellow
    return
}

if (Get-Module -Name 'Microsoft.Graph*') {
    Write-Host "Microsoft.Graph SDK modules are loaded and conflict with PnP.PowerShell." -ForegroundColor Red
    Write-Host "  Open a fresh PowerShell 7 window and run this script again." -ForegroundColor Yellow
    return
}

$pnp = Get-Module -ListAvailable -Name 'PnP.PowerShell' | Sort-Object Version -Descending | Select-Object -First 1
if (-not $pnp) {
    Write-Host "Missing module: PnP.PowerShell" -ForegroundColor Red
    Write-Host "  Install-Module PnP.PowerShell -Scope CurrentUser -Force" -ForegroundColor Yellow
    return
}
Import-Module PnP.PowerShell -ErrorAction Stop
Write-Host ("  OK  PnP.PowerShell v{0}" -f $pnp.Version) -ForegroundColor DarkGreen

if (-not (Get-Command New-SelfSignedCertificate -ErrorAction SilentlyContinue)) {
    Write-Host "New-SelfSignedCertificate is unavailable - setup needs Windows." -ForegroundColor Red
    Write-Host "Supply -ClientId and -CertificateThumbprint for an app created elsewhere." -ForegroundColor Yellow
    if (-not $ClientId) { return }
}
Write-Host ""

# ===========================================================================
# Input
# ===========================================================================
function Read-RequiredValue {
    param([string]$Label, [string]$Example, [string]$Pattern, [string]$Hint, [string]$Current)
    $value = $Current
    while ($true) {
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            $value = $value.Trim().Trim('"').Trim("'").TrimEnd(']', ')')
            if (-not $Pattern -or $value -match $Pattern) { return $value }
            Write-Host ("  Not a valid {0}. {1}" -f $Label, $Hint) -ForegroundColor Yellow
        }
        $value = Read-Host ("{0}  [example: {1}]" -f $Label, $Example)
    }
}

if ($TenantName) { $TenantName = ($TenantName -replace '^https?://', '' -split '\.')[0] -replace '-admin$', '' }
if ($CertificateThumbprint) { $CertificateThumbprint = $CertificateThumbprint -replace '[\s:]', '' }

$TenantName = Read-RequiredValue -Label 'SharePoint tenant short name' `
    -Example 'contoso   (the bit before .sharepoint.com)' `
    -Pattern '^[A-Za-z0-9][A-Za-z0-9\-]{1,62}$' `
    -Hint 'Short name only, no dots and no .sharepoint.com' -Current $TenantName

$TenantDomain = Read-RequiredValue -Label 'Tenant domain' `
    -Example 'contoso.onmicrosoft.com' `
    -Pattern '^[A-Za-z0-9\-\.]+\.[A-Za-z]{2,}$' `
    -Hint 'A full domain, e.g. contoso.onmicrosoft.com' -Current $TenantDomain

if (-not $ConfigPath) {
    $baseDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
    $ConfigPath = Join-Path $baseDir ("SPO-OTP-Audit.{0}.config.json" -f $TenantName)
}

# ===========================================================================
# Helpers
# ===========================================================================
function Save-AppConfig {
    param([string]$Path, [string]$AppName, [string]$TenantName, [string]$TenantDomain,
          [string]$ClientId, [string]$Thumbprint, [bool]$Consented = $false)
    if (-not $ClientId) { return }
    try {
        [pscustomobject]@{
            AppName = $AppName; TenantName = $TenantName; TenantDomain = $TenantDomain
            ClientId = $ClientId; CertificateThumbprint = $Thumbprint
            ConsentGranted = $Consented
            SavedUtc = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
        } | ConvertTo-Json | Set-Content -Path $Path -Encoding UTF8
    }
    catch { Write-Warning ("Could not cache settings: {0}" -f $_.Exception.Message) }
}

function Get-TenantIdFromToken {
    param([string]$AccessToken)
    try {
        $p = $AccessToken.Split('.')[1].Replace('-', '+').Replace('_', '/')
        while ($p.Length % 4) { $p += '=' }
        return ([Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($p)) | ConvertFrom-Json).tid
    } catch { return $null }
}

function Get-ExpectedTenantId {
    # Resolve the tenant GUID for the domain so a wrong-tenant sign-in is caught
    param([string]$Domain)
    try {
        $cfg = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$Domain/v2.0/.well-known/openid-configuration"
        if ($cfg.issuer -match '([0-9a-fA-F-]{36})') { return $Matches[1] }
    } catch { }
    return $null
}

function Get-GraphTokenByDeviceCode {
    param([string]$Tenant, [string[]]$Scopes)

    $publicClientId = '14d82eec-204b-4c2f-b7e8-296a70dab67e'   # Microsoft Graph PowerShell (public client)
    $scopeString = (($Scopes | ForEach-Object { "https://graph.microsoft.com/$_" }) -join ' ') + ' offline_access'

    $dc = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$Tenant/oauth2/v2.0/devicecode" `
        -Body @{ client_id = $publicClientId; scope = $scopeString } -ErrorAction Stop

    $copied = $false
    try { Set-Clipboard -Value $dc.user_code; $copied = $true } catch { }

    Write-Host ""
    Write-Host "  ======================================================" -ForegroundColor Yellow
    Write-Host "   SIGN-IN CODE - shown HERE, never in the browser" -ForegroundColor Yellow
    Write-Host "  ======================================================" -ForegroundColor Yellow
    Write-Host ""
    Write-Host ("     {0}" -f $dc.user_code) -ForegroundColor Black -BackgroundColor White
    Write-Host ""
    Write-Host ("   Enter it at : {0}" -f $dc.verification_uri) -ForegroundColor White
    if ($copied) { Write-Host "   Already copied - press Ctrl+V in the code box." -ForegroundColor Green }
    Write-Host "   Sign in as a Global Administrator or Privileged Role Administrator of $Tenant." -ForegroundColor DarkGray
    Write-Host "   This is the ONLY sign-in. There is no second code." -ForegroundColor DarkGray
    Write-Host ""
    Read-Host "   Press Enter to open the sign-in page"
    try { Start-Process $dc.verification_uri | Out-Null } catch { }
    Write-Host "   Waiting for sign-in..." -ForegroundColor DarkGray

    $deadline = (Get-Date).AddSeconds([int]$dc.expires_in)
    $interval = [Math]::Max([int]$dc.interval, 3)

    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds $interval
        try {
            $tok = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$Tenant/oauth2/v2.0/token" `
                -Body @{
                    grant_type  = 'urn:ietf:params:oauth:grant-type:device_code'
                    client_id   = $publicClientId
                    device_code = $dc.device_code
                } -ErrorAction Stop
            Write-Host "   Signed in." -ForegroundColor Green
            return $tok.access_token
        }
        catch {
            $detail = ''
            try { $detail = ($_.ErrorDetails.Message | ConvertFrom-Json).error } catch { }
            switch ($detail) {
                'authorization_pending'  { continue }
                'slow_down'              { $interval += 5; continue }
                'authorization_declined' { throw 'Sign-in was declined.' }
                'expired_token'          { throw 'The code expired before sign-in completed.' }
                default { throw ("Token request failed: {0}" -f $(if ($detail) { $detail } else { $_.Exception.Message })) }
            }
        }
    }
    throw 'Timed out waiting for sign-in.'
}

function Invoke-Graph {
    param([hashtable]$Headers, [string]$Method = 'Get', [string]$Uri, $Body)
    # Not named $args: that is a PowerShell automatic variable.
    $request = @{ Headers = $Headers; Method = $Method; Uri = $Uri; ErrorAction = 'Stop' }
    if ($Body) { $request['Body'] = ($Body | ConvertTo-Json -Depth 10) }
    return Invoke-RestMethod @request
}

function Get-KeyCredentialThumbprint {
    # Thumbprint of a certificate registered on an app. Computed from the certificate's own
    # bytes ('key') when present, so it doesn't depend on how customKeyIdentifier is encoded.
    # Fallback: customKeyIdentifier, documented as a 40-character value defaulting to the
    # thumbprint - accepted as hex text, base64 of that text, or base64 of the raw 20 bytes.
    param($KeyCredential)
    if ($KeyCredential.key) {
        try {
            $cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new([Convert]::FromBase64String([string]$KeyCredential.key))
            return $cert.Thumbprint.ToUpper()
        } catch { }
    }
    $id = [string]$KeyCredential.customKeyIdentifier
    if (-not $id) { return $null }
    if ($id -match '^[0-9A-Fa-f]{40}$') { return $id.ToUpper() }
    try {
        $bytes = [Convert]::FromBase64String($id)
        if ($bytes.Length -eq 20) { return ([BitConverter]::ToString($bytes) -replace '-', '') }
        $text = [Text.Encoding]::ASCII.GetString($bytes)
        if ($text -match '^[0-9A-Fa-f]{40}$') { return $text.ToUpper() }
    } catch { }
    return $null
}

function Initialize-AppRegistration {
    <#
        Creates or reuses the app registration, certificate, service principal and
        admin consent. One sign-in, idempotent, no PnP registration cmdlets.
    #>
    param([string]$Tenant, [string]$AppName, [hashtable]$Permissions, [int]$CertYears)

    $base = 'https://graph.microsoft.com/v1.0'

    $expectedTid = Get-ExpectedTenantId -Domain $Tenant
    $token = Get-GraphTokenByDeviceCode -Tenant $Tenant `
        -Scopes @('Application.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All', 'Directory.Read.All')
    $headers = @{ Authorization = "Bearer $token"; 'Content-Type' = 'application/json' }

    $actualTid = Get-TenantIdFromToken -AccessToken $token
    Write-Host ("  Tenant: {0}" -f $actualTid) -ForegroundColor DarkGray
    if ($expectedTid -and $actualTid -and $expectedTid -ne $actualTid) {
        throw ("Wrong tenant. '{0}' is tenant {1} but you signed in to {2}. " -f $Tenant, $expectedTid, $actualTid) +
              "Sign in with an admin account of $Tenant, not another directory."
    }

    # --- resolve resource service principals and their app roles ---
    $resources = @{}
    foreach ($resAppId in $Permissions.Keys) {
        $sp = @((Invoke-Graph -Headers $headers -Uri "$base/servicePrincipals?`$filter=appId eq '$resAppId'").value)[0]
        if (-not $sp) { throw "Resource app $resAppId not found in this tenant." }
        $roles = @()
        foreach ($v in $Permissions[$resAppId]) {
            $r = @($sp.appRoles) | Where-Object { $_.value -eq $v -and $_.allowedMemberTypes -contains 'Application' } | Select-Object -First 1
            if (-not $r) { throw "Permission $v not found on $($sp.displayName)." }
            $roles += [pscustomobject]@{ Value = $v; Id = $r.id }
        }
        $resources[$resAppId] = [pscustomobject]@{ Sp = $sp; Roles = $roles }
    }

    # --- find or create the application ---
    $escaped = $AppName.Replace("'", "''")
    $app = @((Invoke-Graph -Headers $headers -Uri "$base/applications?`$filter=displayName eq '$escaped'").value)[0]
    $createdApp = $false

    if ($app) {
        Write-Host ("  App exists  : {0}" -f $app.appId) -ForegroundColor DarkGreen
    }
    else {
        $requiredResourceAccess = foreach ($resAppId in $resources.Keys) {
            @{
                resourceAppId  = $resAppId
                resourceAccess = @(foreach ($r in $resources[$resAppId].Roles) { @{ id = $r.Id; type = 'Role' } })
            }
        }
        $app = Invoke-Graph -Headers $headers -Method Post -Uri "$base/applications" -Body @{
            displayName            = $AppName
            signInAudience         = 'AzureADMyOrg'
            requiredResourceAccess = @($requiredResourceAccess)
        }
        $createdApp = $true
        Write-Host ("  App created : {0}" -f $app.appId) -ForegroundColor Green
    }

    # --- certificate: reuse a local one the app already trusts, else create and upload ---
    # Graph only returns each certificate's bytes ('key') for a single-app GET with $select.
    $appKeys = @($app.keyCredentials)
    if (-not $createdApp) {
        try { $appKeys = @((Invoke-Graph -Headers $headers -Uri "$base/applications/$($app.id)?`$select=keyCredentials").keyCredentials) } catch { }
    }
    $appCertThumbs = @($appKeys | ForEach-Object { Get-KeyCredentialThumbprint -KeyCredential $_ } | Where-Object { $_ })

    $localCert = Get-ChildItem Cert:\CurrentUser\My -ErrorAction SilentlyContinue |
                 Where-Object { $_.Subject -like "*CN=$AppName*" -and $_.NotAfter -gt (Get-Date).AddDays(30) -and
                                $appCertThumbs -contains $_.Thumbprint.ToUpper() } |
                 Sort-Object NotAfter -Descending | Select-Object -First 1

    if ($localCert) {
        Write-Host ("  Certificate : reusing {0}" -f $localCert.Thumbprint) -ForegroundColor DarkGreen
    }
    else {
        # Uploading below REPLACES the app's whole keyCredentials collection, so ask before
        # discarding certificates this machine doesn't hold - something else may use them.
        if (-not $createdApp -and $appKeys.Count -gt 0) {
            Write-Host ("  App '{0}' has certificate(s) that are not in this user's store:" -f $AppName) -ForegroundColor Yellow
            foreach ($k in $appKeys) {
                Write-Host ("     {0}  expires {1}" -f (Get-KeyCredentialThumbprint -KeyCredential $k), $k.endDateTime) -ForegroundColor Yellow
            }
            Write-Host "  Replacing them breaks anything else that signs in with them." -ForegroundColor Yellow
            $answer = "$(Read-Host '  Replace with a new certificate from this machine? [y/N]')".Trim()
            if ($answer -notmatch '^(y|yes)$') {
                throw "Stopped without changing the app. Run again with -AppName <new name> to create a separate app instead."
            }
        }

        Write-Host "  Certificate : creating a new one" -ForegroundColor Cyan
        $localCert = New-SelfSignedCertificate -Subject "CN=$AppName" `
            -CertStoreLocation 'Cert:\CurrentUser\My' -KeyExportPolicy Exportable `
            -KeySpec Signature -KeyAlgorithm RSA -KeyLength 2048 `
            -NotAfter (Get-Date).AddYears($CertYears) -ErrorAction Stop

        Invoke-Graph -Headers $headers -Method Patch -Uri "$base/applications/$($app.id)" -Body @{
            keyCredentials = @(@{
                type        = 'AsymmetricX509Cert'
                usage       = 'Verify'
                key         = [Convert]::ToBase64String($localCert.RawData)
                displayName = "CN=$AppName"
            })
        } | Out-Null
        Write-Host ("  Certificate : uploaded {0}" -f $localCert.Thumbprint) -ForegroundColor Green
    }

    # --- service principal ---
    $clientSp = @((Invoke-Graph -Headers $headers -Uri "$base/servicePrincipals?`$filter=appId eq '$($app.appId)'").value)[0]
    if (-not $clientSp) {
        $clientSp = $null
        for ($attempt = 1; $attempt -le 5 -and -not $clientSp; $attempt++) {
            try { $clientSp = Invoke-Graph -Headers $headers -Method Post -Uri "$base/servicePrincipals" -Body @{ appId = $app.appId } }
            catch { Start-Sleep -Seconds 5 }
        }
        if (-not $clientSp) { throw "Could not create the service principal for $($app.appId)." }
        Write-Host "  Service principal created" -ForegroundColor Green
    }

    # --- admin consent as app role assignments ---
    Write-Host "  Consent     :" -ForegroundColor Cyan
    $existing = @()
    try { $existing = @((Invoke-Graph -Headers $headers -Uri "$base/servicePrincipals/$($clientSp.id)/appRoleAssignments").value) } catch { }

    $failed = 0
    foreach ($resAppId in $resources.Keys) {
        $resSp = $resources[$resAppId].Sp
        foreach ($r in $resources[$resAppId].Roles) {
            if ($existing | Where-Object { $_.appRoleId -eq $r.Id -and $_.resourceId -eq $resSp.id }) {
                Write-Host ("     {0} : already granted" -f $r.Value) -ForegroundColor DarkGreen
                continue
            }
            try {
                Invoke-Graph -Headers $headers -Method Post `
                    -Uri "$base/servicePrincipals/$($clientSp.id)/appRoleAssignments" `
                    -Body @{ principalId = $clientSp.id; resourceId = $resSp.id; appRoleId = $r.Id } | Out-Null
                Write-Host ("     {0} : granted" -f $r.Value) -ForegroundColor Green
            }
            catch {
                Write-Host ("     {0} : FAILED - {1}" -f $r.Value, $_.Exception.Message) -ForegroundColor Red
                $failed++
            }
        }
    }
    if ($failed -gt 0) { throw "$failed permission(s) could not be granted. The signed-in account may lack Global Administrator or Privileged Role Administrator." }

    return [pscustomobject]@{ ClientId = $app.appId; Thumbprint = $localCert.Thumbprint; TenantId = $actualTid }
}

function Connect-WithRetry {
    # Consent and new certificates take a moment to propagate - retry instead of
    # telling the user to wait and run again.
    param([string]$Url, [string]$ClientId, [string]$Tenant, [string]$Thumbprint, [int]$Attempts = 6)
    for ($i = 1; $i -le $Attempts; $i++) {
        try {
            Connect-PnPOnline -Url $Url -ClientId $ClientId -Tenant $Tenant -Thumbprint $Thumbprint
            return $true
        }
        catch {
            if ($i -eq $Attempts) { throw }
            Write-Host ("  Attempt {0}/{1} failed, retrying in 20s (consent propagation)..." -f $i, $Attempts) -ForegroundColor DarkYellow
            Start-Sleep -Seconds 20
        }
    }
}

function Invoke-WithPropagationRetry {
    # A new app certificate or consent replicates across Entra over a minute or two, and
    # until it has, a token request can fail on one replica after succeeding on another:
    # SharePoint connects, then the first Graph call fails with AADSTS700027 ("certificate
    # not registered on application"). SharePoint is slowest of all to honour a new app's
    # Sites.FullControl.All grant, so its first tenant-admin call can return 401
    # Unauthorized even though Connect-PnPOnline succeeded - that cmdlet doesn't check
    # authorisation, the first real call does. A brand-new service principal can also be
    # unknown to Graph for a while: "The identity of the calling application could not be
    # established" (Authorization_IdentityNotFound).
    #
    # Each layer fails with a different message, so matching messages alone keeps missing
    # one. Right after setup the caller passes -AnyError: the app is known to be settling,
    # so every error is retried. On a run that reuses a saved setup, only the known
    # replication errors are retried, and a genuine failure still surfaces immediately.
    #
    # -BeforeRetry runs before every retry. Pass a reconnect: a token issued before consent
    # replicated has no roles and stays cached for its lifetime, so retrying on the same
    # connection fails forever.
    param([scriptblock]$Action, [string]$What, [int]$Attempts = 6, [int]$DelaySeconds = 20,
          [switch]$AnyError, [scriptblock]$BeforeRetry)
    for ($i = 1; $i -le $Attempts; $i++) {
        try { return & $Action }
        catch {
            $msg = $_.Exception.Message
            $transient = $AnyError -or ($msg -match 'AADSTS700027|AADSTS700016|Authorization_IdentityNotFound|identity of the calling application could not be established|Authorization_RequestDenied|Insufficient privileges|\(403\)|Forbidden|\(401\)|Unauthorized')
            if (-not $transient -or $i -eq $Attempts) { throw }
            Write-Host ("  {0}: attempt {1}/{2} failed, retrying in {3}s (Entra propagation)..." -f $What, $i, $Attempts, $DelaySeconds) -ForegroundColor DarkYellow
            Start-Sleep -Seconds $DelaySeconds
            if ($BeforeRetry) { try { & $BeforeRetry } catch { } }
        }
    }
}

function Test-IsExternalLogin {
    param([string]$LoginName)
    return ($LoginName -like '*urn:spo:guest*' -or $LoginName -like '*spo%3aguest*' -or $LoginName -like '*#EXT#*')
}

function Convert-ExtUpnToEmail {
    param([string]$Upn)
    if ([string]::IsNullOrWhiteSpace($Upn)) { return $null }
    if ($Upn -notlike '*#EXT#*') { return $Upn }
    $prefix = ($Upn -split '#EXT#')[0]
    $idx = $prefix.LastIndexOf('_')
    if ($idx -gt 0) { return ('{0}@{1}' -f $prefix.Substring(0, $idx), $prefix.Substring($idx + 1)) }
    return $prefix
}

function Get-EmailFromLoginName {
    param([string]$LoginName)
    if ([string]::IsNullOrWhiteSpace($LoginName)) { return $null }
    if ($LoginName -like '*#EXT#*') { return Convert-ExtUpnToEmail -Upn (($LoginName -split '\|')[-1]) }
    $m = [regex]::Match($LoginName, "[\w\.\-\+']+@[\w\.\-]+\.\w+")
    if ($m.Success) { return $m.Value }
    return $null
}

function Get-RestItems {
    # Invoke-PnPSPRestMethod response shape varies by version/accept header
    param($Response)
    if ($null -eq $Response) { return @() }
    if ($Response.PSObject.Properties.Name -contains 'value')   { return @($Response.value) }
    if ($Response.PSObject.Properties.Name -contains 'd' -and
        $Response.d.PSObject.Properties.Name -contains 'results') { return @($Response.d.results) }
    if ($Response.PSObject.Properties.Name -contains 'd')        { return @($Response.d) }
    return @($Response)
}

function Get-RestNextLink {
    param($Response)
    if ($null -eq $Response) { return $null }
    if ($Response.PSObject.Properties.Name -contains 'odata.nextLink')  { return $Response.'odata.nextLink' }
    if ($Response.PSObject.Properties.Name -contains '@odata.nextLink') { return $Response.'@odata.nextLink' }
    if ($Response.PSObject.Properties.Name -contains 'd' -and
        $Response.d.PSObject.Properties.Name -contains '__next')        { return $Response.d.__next }
    return $null
}

function Get-SiteExternalItemMap {
    <#
        Builds: lowercase email -> list of item paths, for the CURRENT site.

        External access to an item lands in one of two shapes, and this handles both:
          a) the guest is a DIRECT principal on an item with unique permissions
          b) the guest is a MEMBER of a SharingLinks.<guid> group that is a principal

        Scale: items are paged in bulk with HasUniqueRoleAssignments selected, so a
        large library costs a handful of calls, not one per item. Only the items that
        come back with unique permissions cost one extra call each.
    #>
    param([int]$MaxItemsPerList = 20000)

    $map = @{}

    # --- (b) group id -> external emails, one pass over site groups ---
    $groupExternals = @{}
    try {
        foreach ($g in @(Get-PnPGroup -ErrorAction Stop)) {
            $members = @()
            try { $members = @(Get-PnPProperty -ClientObject $g -Property Users -ErrorAction Stop) } catch { continue }
            $emails = @()
            foreach ($m in $members) {
                if (-not (Test-IsExternalLogin -LoginName $m.LoginName)) { continue }
                $e = $m.Email
                if ([string]::IsNullOrWhiteSpace($e)) { $e = Get-EmailFromLoginName -LoginName $m.LoginName }
                if ($e) { $emails += $e.ToLower() }
            }
            if ($emails.Count -gt 0) { $groupExternals[[string]$g.Id] = $emails }
        }
    }
    catch { }

    # --- (a) items with unique permissions ---
    $lists = @()
    try { $lists = @(Get-PnPList -ErrorAction Stop | Where-Object { $_.Hidden -eq $false -and $_.BaseTemplate -in @(101, 700) }) }
    catch { return $map }

    foreach ($list in $lists) {
        $listId = $list.Id
        $uniqueItems = [System.Collections.Generic.List[object]]::new()

        $url = "/_api/web/lists(guid'$listId')/items?`$select=Id,FileRef,HasUniqueRoleAssignments&`$top=5000"
        $seen = 0

        while ($url -and $seen -lt $MaxItemsPerList) {
            $resp = $null
            try { $resp = Invoke-PnPSPRestMethod -Url $url -Method Get -ErrorAction Stop }
            catch { break }

            foreach ($row in (Get-RestItems -Response $resp)) {
                $seen++
                if ($row.HasUniqueRoleAssignments -eq $true) {
                    $uniqueItems.Add([pscustomobject]@{ Id = $row.Id; Path = $row.FileRef })
                }
            }
            $url = Get-RestNextLink -Response $resp
            if ($url -and $url -match '^https?://[^/]+(/.*)$') { $url = $Matches[1] }
        }

        foreach ($item in $uniqueItems) {
            $raUrl = "/_api/web/lists(guid'$listId')/items($($item.Id))/roleassignments?`$expand=Member"
            $ras = $null
            try { $ras = Invoke-PnPSPRestMethod -Url $raUrl -Method Get -ErrorAction Stop }
            catch { continue }

            foreach ($ra in (Get-RestItems -Response $ras)) {
                $member = $ra.Member
                if (-not $member) { continue }

                $emails = @()

                if (Test-IsExternalLogin -LoginName $member.LoginName) {
                    # (a) direct external principal
                    $e = Get-EmailFromLoginName -LoginName $member.LoginName
                    if ($e) { $emails += $e.ToLower() }
                }
                elseif ($groupExternals.ContainsKey([string]$member.Id)) {
                    # (b) a group - attribute to its external members
                    $emails += $groupExternals[[string]$member.Id]
                }

                foreach ($e in ($emails | Select-Object -Unique)) {
                    if (-not $map.ContainsKey($e)) { $map[$e] = [System.Collections.Generic.List[string]]::new() }
                    if (-not $map[$e].Contains($item.Path)) { $map[$e].Add($item.Path) }
                }
            }
        }
    }
    return $map
}

# ===========================================================================
# Setup - cached config first, otherwise run it
# ===========================================================================
$fromCache = $false
if (-not $Reset -and -not $ClientId -and (Test-Path $ConfigPath)) {
    try {
        $cfg = Get-Content $ConfigPath -Raw | ConvertFrom-Json
        $ClientId = $cfg.ClientId
        $CertificateThumbprint = $cfg.CertificateThumbprint
        $fromCache = $true
    }
    catch { Write-Warning ("Could not read {0}: {1}" -f $ConfigPath, $_.Exception.Message) }
}

# A certificate that is missing or expired means setup has to run again
$haveCert = $null
if ($ClientId -and $CertificateThumbprint) {
    $haveCert = Get-ChildItem Cert:\CurrentUser\My -ErrorAction SilentlyContinue |
                Where-Object { $_.Thumbprint -eq $CertificateThumbprint -and $_.NotAfter -gt (Get-Date) } |
                Select-Object -First 1
    if (-not $haveCert) {
        Write-Host "Saved certificate is missing or expired - running setup again." -ForegroundColor Yellow
        $ClientId = $null; $CertificateThumbprint = $null; $fromCache = $false
    }
}

# A usable saved setup exists: ask rather than assume, unless told to use it
if ($fromCache -and -not $UseCachedSetup) {
    Write-Host ""
    Write-Host ("Found a saved setup for {0}:" -f $TenantName) -ForegroundColor Cyan
    Write-Host ("  App         : {0}" -f $cfg.AppName)
    Write-Host ("  Client ID   : {0}" -f $ClientId)
    Write-Host ("  Certificate : {0} (expires {1:yyyy-MM-dd})" -f $CertificateThumbprint, $haveCert.NotAfter)
    if ($cfg.SavedUtc) {
        # ConvertFrom-Json turns the ISO timestamp into a DateTime on PowerShell 7, which
        # would otherwise print in local format with no zone.
        $saved = if ($cfg.SavedUtc -is [datetime]) { $cfg.SavedUtc.ToUniversalTime().ToString('yyyy-MM-dd HH:mm') }
                 else { "$($cfg.SavedUtc)" -replace 'T', ' ' -replace ':\d{2}Z$', '' }
        Write-Host ("  Saved       : {0} UTC" -f $saved)
    }
    Write-Host ""
    Write-Host "  [U] Use this saved setup (default)"
    Write-Host "  [R] Run setup again - re-checks the app, certificate and consent, recreating anything missing"
    Write-Host "  [Q] Quit"

    while ($true) {
        $choice = "$(Read-Host '  Choose U, R or Q')".Trim().ToUpper()
        if ($choice -in @('', 'U')) {
            Write-Host ("Using saved setup: {0} ({1})" -f $cfg.AppName, $ClientId) -ForegroundColor DarkGreen
            break
        }
        if ($choice -eq 'R') {
            $Reset = $true
            $ClientId = $null; $CertificateThumbprint = $null
            break
        }
        if ($choice -eq 'Q') { return }
        Write-Host "  Enter U, R or Q." -ForegroundColor Yellow
    }
}
elseif ($fromCache) {
    Write-Host ("Using saved setup: {0} ({1})" -f $cfg.AppName, $ClientId) -ForegroundColor DarkGreen
}

$freshSetup = $false
if ($Reset -or -not $ClientId -or -not $CertificateThumbprint) {
    Write-Host ""
    Write-Host "SETUP - app registration, certificate and admin consent" -ForegroundColor Yellow
    Write-Host "One sign-in. Everything else is automatic." -ForegroundColor DarkGray

    $setup = Initialize-AppRegistration -Tenant $TenantDomain -AppName $AppName `
        -Permissions $RequiredPermissions -CertYears $CertYears

    $ClientId = $setup.ClientId
    $CertificateThumbprint = $setup.Thumbprint
    Save-AppConfig -Path $ConfigPath -AppName $AppName -TenantName $TenantName `
        -TenantDomain $TenantDomain -ClientId $ClientId -Thumbprint $CertificateThumbprint -Consented $true

    Write-Host ""
    Write-Host "Setup complete." -ForegroundColor Green
    Write-Host "A new app takes a few minutes to be recognised everywhere - first calls may retry." -ForegroundColor DarkGray
    $freshSetup = $true
}

# Right after setup, any failure is most likely the new app still replicating, so retry
# everything for up to ~5 minutes. With a saved setup, retry known replication errors only.
$retryAttempts = if ($freshSetup) { 15 } else { 6 }

Write-Host ""
Write-Host ("Tenant      : {0}.sharepoint.com" -f $TenantName) -ForegroundColor DarkCyan
Write-Host ("OneDrive    : {0}" -f $(if ($IncludeOneDrive) { 'included' } else { 'excluded' })) -ForegroundColor DarkCyan
if ($SiteUrlFilter) { Write-Host ("Site filter : {0}" -f $SiteUrlFilter) -ForegroundColor DarkCyan }
if ($IncludeSharedItems) { Write-Host "Shared items: resolving paths (slower)" -ForegroundColor DarkCyan }
Write-Host ""

# ===========================================================================
# 1. Connect
# ===========================================================================
$adminUrl = "https://$TenantName-admin.sharepoint.com"
Write-Host "Connecting to $adminUrl" -ForegroundColor Cyan

try {
    Connect-WithRetry -Url $adminUrl -ClientId $ClientId -Tenant $TenantDomain -Thumbprint $CertificateThumbprint | Out-Null
}
catch {
    Write-Host ""
    Write-Host ("Connect failed after retries: {0}" -f $_.Exception.Message) -ForegroundColor Red
    Write-Host "Run again with -Reset to rebuild the app registration and consent." -ForegroundColor Yellow
    return
}

# Disconnect first: a token issued before the grant replicated has no roles and is cached
# for its lifetime, so reconnecting without tearing the old connection down can hand back
# the same useless token and every retry then fails for the same reason.
$reconnect = {
    try { Disconnect-PnPOnline } catch { }
    Connect-PnPOnline -Url $adminUrl -ClientId $ClientId -Tenant $TenantDomain -Thumbprint $CertificateThumbprint
}

# ===========================================================================
# 2. Entra guest index
# ===========================================================================
Write-Host "Building Entra guest index..." -ForegroundColor Cyan
$entraGuestIndex = @{}
$guestCount = 0

try {
    $url = "v1.0/users?`$filter=userType eq 'Guest'&`$select=id,displayName,mail,userPrincipalName,externalUserState&`$top=999"
    while ($url) {
        $page = Invoke-WithPropagationRetry -What 'Reading guests' -Attempts $retryAttempts -AnyError:$freshSetup `
                    -BeforeRetry $reconnect -Action { Invoke-PnPGraphMethod -Url $url -Method Get }
        foreach ($g in @($page.value)) {
            $guestCount++
            $keys = @()
            if ($g.mail) { $keys += $g.mail }
            $derived = Convert-ExtUpnToEmail -Upn $g.userPrincipalName
            if ($derived) { $keys += $derived }
            foreach ($k in ($keys | Where-Object { $_ } | Select-Object -Unique)) { $entraGuestIndex[$k.ToLower()] = $g }
        }
        $url = $page.'@odata.nextLink'
    }
    Write-Host ("  {0} Entra guest accounts indexed" -f $guestCount) -ForegroundColor Green
}
catch {
    Write-Host ("Could not read guests from Graph: {0}" -f $_.Exception.Message) -ForegroundColor Red
    Disconnect-PnPOnline
    return
}

# ===========================================================================
# 3. Sites
# ===========================================================================
Write-Host "Enumerating sites..." -ForegroundColor Cyan
try {
    $sites = Invoke-WithPropagationRetry -What 'Enumerating sites' -Attempts $retryAttempts -AnyError:$freshSetup -BeforeRetry $reconnect -Action {
                 Get-PnPTenantSite -IncludeOneDriveSites:$IncludeOneDrive
             } | Where-Object { $_.Template -notlike 'RedirectSite*' }
}
catch {
    Write-Host ""
    Write-Host ("Could not list sites: {0}" -f $_.Exception.Message) -ForegroundColor Red
    if ($_.Exception.Message -match 'Unauthorized|\(401\)|\(403\)|Forbidden') {
        Write-Host "SharePoint has not accepted this app's Sites.FullControl.All grant yet. This is" -ForegroundColor Yellow
        Write-Host "normal for a newly created app and usually clears within 10-20 minutes." -ForegroundColor Yellow
        Write-Host "  1. Wait, then run this script again and choose [U] to reuse the same setup." -ForegroundColor Yellow
        Write-Host "  2. Do NOT re-run setup - a new certificate restarts the wait." -ForegroundColor Yellow
        Write-Host "  3. Still failing? Run Test-SPOOTPSetup.ps1 to see whether the token carries the role." -ForegroundColor Yellow
    }
    Disconnect-PnPOnline
    return
}
if ($SiteUrlFilter) { $sites = $sites | Where-Object { $_.Url -like $SiteUrlFilter } }
Write-Host ("  {0} sites in scope" -f @($sites).Count) -ForegroundColor Green

# ===========================================================================
# 4. Scan
# ===========================================================================
$results     = [System.Collections.Generic.List[object]]::new()
$failedSites = [System.Collections.Generic.List[object]]::new()
$i = 0

foreach ($site in $sites) {
    $i++
    Write-Progress -Activity 'Scanning sites for external users' `
        -Status ("{0} of {1} - {2}" -f $i, @($sites).Count, $site.Url) `
        -PercentComplete (($i / [math]::Max(@($sites).Count, 1)) * 100)

    try {
        Connect-PnPOnline -Url $site.Url -ClientId $ClientId -Tenant $TenantDomain `
            -Thumbprint $CertificateThumbprint

        $sharingMap = @{}
        if ($IncludeSharedItems) { $sharingMap = Get-SiteExternalItemMap -MaxItemsPerList $MaxItemsPerList }

        $externalUsers = Get-PnPUser | Where-Object { Test-IsExternalLogin -LoginName $_.LoginName }

        foreach ($u in $externalUsers) {
            $isSpoOtp = ($u.LoginName -like '*urn:spo:guest*' -or $u.LoginName -like '*spo%3aguest*')
            $type     = if ($isSpoOtp) { 'SPO-OTP' } else { 'Entra-B2B' }

            $email = $u.Email
            if ([string]::IsNullOrWhiteSpace($email)) { $email = Get-EmailFromLoginName -LoginName $u.LoginName }

            $entraMatch = $null
            if ($email) { $entraMatch = $entraGuestIndex[$email.ToLower()] }

            $sharedPaths = @()
            if ($IncludeSharedItems -and $email -and $sharingMap.ContainsKey($email.ToLower())) {
                $sharedPaths = @($sharingMap[$email.ToLower()])
            }

            $results.Add([pscustomobject]@{
                SiteUrl          = $site.Url
                SiteTitle        = $site.Title
                SiteType         = if ($site.Template -like 'SPSPERS*') { 'OneDrive' } else { 'SharePoint' }
                GuestType        = $type
                DisplayName      = $u.Title
                Email            = $email
                LoginName        = $u.LoginName
                HasEntraGuest    = if ($entraMatch) { 'Yes' } else { 'No' }
                EntraUserId      = $entraMatch.id
                EntraInviteState = $entraMatch.externalUserState
                AtRisk           = if ($type -eq 'SPO-OTP' -and -not $entraMatch) { 'YES - will lose access' } else { '' }
                SharedItemCount  = if ($IncludeSharedItems) { $sharedPaths.Count } else { '' }
                SharedItems      = if ($IncludeSharedItems) { ($sharedPaths | Select-Object -First $MaxItemsPerUser) -join ' | ' } else { '' }
            })
        }
    }
    catch {
        $failedSites.Add([pscustomobject]@{ SiteUrl = $site.Url; Error = $_.Exception.Message })
        Write-Warning ("Skipped {0}: {1}" -f $site.Url, $_.Exception.Message)
    }
}

Write-Progress -Activity 'Scanning sites for external users' -Completed

# ===========================================================================
# 5. Output
# ===========================================================================
$results | Sort-Object AtRisk -Descending | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

$atRisk       = @($results | Where-Object { $_.AtRisk })
$uniqueAtRisk = @($atRisk | Select-Object -ExpandProperty Email -Unique | Where-Object { $_ })

Write-Host ""
Write-Host "================ SUMMARY ================" -ForegroundColor Yellow
Write-Host ("Sites scanned          : {0}" -f @($sites).Count)
Write-Host ("Sites failed/skipped   : {0}" -f $failedSites.Count)
Write-Host ("External user records  : {0}" -f $results.Count)
Write-Host ("SPO-OTP guests         : {0}" -f @($results | Where-Object GuestType -eq 'SPO-OTP').Count)
Write-Host ("AT RISK records        : {0}" -f $atRisk.Count) -ForegroundColor Red
Write-Host ("AT RISK unique people  : {0}" -f $uniqueAtRisk.Count) -ForegroundColor Red
Write-Host ("Report                 : {0}" -f (Resolve-Path $OutputPath))
Write-Host "=========================================" -ForegroundColor Yellow

if ($failedSites.Count -gt 0) {
    $failPath = $OutputPath -replace '\.csv$', '_FailedSites.csv'
    $failedSites | Export-Csv -Path $failPath -NoTypeInformation -Encoding UTF8
    Write-Host ("Failed sites logged    : {0}" -f $failPath) -ForegroundColor DarkYellow
}

Write-Host ""
Write-Host "Reminder: guests who never signed in do not appear in the site user list." -ForegroundColor DarkCyan
Write-Host "Cross-check with the site-level external sharing report in the SharePoint admin centre." -ForegroundColor DarkCyan
Write-Host "The CSV contains external users' names and email addresses - handle it as confidential." -ForegroundColor DarkCyan
Write-Host "When the audit is finished, delete the app registration - see CLEANUP in Get-Help." -ForegroundColor DarkCyan

Disconnect-PnPOnline
