<#
.SYNOPSIS
    Network connectivity pre-flight for the PowerStacks Windows Inventory script.

.DESCRIPTION
    Tests every network endpoint the Intune Windows Inventory collector
    (Intune_Windows_Inventory.ps1) uses, one at a time, and reports exactly which
    ones fail and why. It does not collect or upload any inventory. It is meant to
    run on a device that is failing so you can tell, quickly, whether the problem is
    DNS, a firewall/proxy block, TLS interception, or proxy authentication under the
    SYSTEM account.

    For each endpoint it runs a layered test and records the first thing that breaks:
      1. DNS resolution of the host
      2. TCP connect to the host and port (443, or 80 for the Lenovo API)
      3. TLS handshake and an HTTP request (any HTTP response, even 401/403/404,
         proves the network path is open; only a timeout, connect failure, DNS
         failure, TLS failure, or a 407 proxy-auth challenge counts as a failure)
    It also reports the effective proxy (WinHTTP, which is what SYSTEM uses, plus the
    .NET system proxy), because the most common cause of "works for me but fails in
    Intune" is that the interactive user has a proxy and the SYSTEM account does not.

    Which endpoints are tested is driven by the settings below, so fill them in to
    match the inventory script you deployed (same $LogAPIMode, same $DceURI, etc.).
    Warranty endpoints are only tested when $CollectWarranty is $true, and by default
    only for this device's manufacturer.

.DEPLOYMENT
    Run as SYSTEM (64-bit). Two easy ways:
      - Intune > Devices > Scripts and remediations > Platform scripts: add this as a
        one-time script (Run as system = Yes, 64-bit = Yes). Read the results in the
        CMTrace-style log it writes to C:\Windows\Logs\.
      - Intune > Remediations: use this as the DETECTION script. It exits 0 when every
        required endpoint passes and 1 when any required endpoint fails, so a failing
        device shows up as "Issue detected." Leave the remediation script empty (or
        add your own follow-up). No admin data leaves the device.

    The full detail is always written to:
      C:\Windows\Logs\PowerStacks_Inventory_Connectivity_<timestamp>.log

.NOTES
    Author: John Marcum (PowerStacks)
    Compatible with Windows PowerShell 5.1 (the Intune script host). No modules required.
    This script does not upload inventory data. With valid $ClientId/$ClientSecret it
    also requests an access token and, when $DcrImmutableId is set, sends one empty
    probe to the real DCR stream to prove the ingestion path end to end (the empty body
    writes nothing). That is the only check that reproduces the collector's upload and
    catches a 404. Set $TestTokenAcquisition = $false or $TestIngestion = $false to stay
    purely connectivity-only.
#>

#region settings ---------------------------------------------------------------
# Copy these from the inventory script you deployed to this customer.

# "LogIngestionAPI" (current) or "DataCollectorAPI" (legacy). Match your collector.
$LogAPIMode = "LogIngestionAPI"

# ----- LogIngestionAPI mode -----
$TenantId        = "<Enter Your Tenant ID>"
$ClientId        = "<Enter Your Client ID>"
$ClientSecret    = "<Enter Your Client Secret>"
$DceURI          = "<Enter Your DCE Log Ingestion URL>"          # https://xxxx.region.ingest.monitor.azure.com
# DCR Immutable ID and stream, copied from the collector. Required for the end-to-end ingestion
# dry-run that reproduces the collector's real upload - the ONLY check that catches an upload 404.
$DcrImmutableId  = "<Enter Your DCR Immutable ID>"               # the DCR's IMMUTABLE ID (dcr-xxxxxxxx...), not its name or resource id
$StreamName      = "Custom-PowerStacksDeviceInventory_CL"        # the device stream the collector posts to first
# Optional Entra Token Broker (secretless) upload path. Leave the placeholder if unused.
$BrokerUrl       = "<Enter Your Broker URL>"

# ----- DataCollectorAPI (legacy) mode -----
$CustomerId      = "<Enter Your Log Analytics Workspace ID>"     # the workspace GUID

# ----- Which optional endpoint groups the collector uses -----
$CollectMicrosoft365      = $true    # tests the Office release/CDN endpoints
$CollectWarranty          = $false   # tests the vendor warranty API for this device
$TestAllWarrantyVendors   = $false   # $true to test Dell, HP, Lenovo and Getac regardless of make

# ----- Behavior -----
$TimeoutSeconds           = 15       # per-request timeout
$TestTokenAcquisition     = $true    # with a real ClientId/Secret, prove the sign-in path too
$TestIngestion            = $true    # POST an empty probe to the real DCR stream (reproduces the collector's upload; catches a 404)
#endregion settings ------------------------------------------------------------


#region init -------------------------------------------------------------------
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$Now    = Get-Date -Format "yyyy-MM-dd_HHmm"
$CMLog  = "C:\Windows\Logs\PowerStacks_Inventory_Connectivity_$Now.log"
$script:CMLog = $CMLog
$script:AccessToken = $null
$script:PreflightFailures = New-Object System.Collections.Generic.List[string]

function Write-CMTraceLog {
    param(
        [Parameter(Mandatory, Position = 0)][string]$Message,
        [ValidateSet(1, 2, 3)][int]$Type = 1,   # 1 Info, 2 Warning, 3 Error
        [string]$Component = "Test-InventoryConnectivity",
        [string]$Path = $script:CMLog
    )
    $time = (Get-Date -Format "HH:mm:ss.fff")
    $date = (Get-Date -Format "MM-dd-yyyy")
    $line = '<![LOG[{0}]LOG]!><time="{1}" date="{2}" component="{3}" context="" type="{4}" thread="{5}" file="">' -f `
        $Message, $time, $date, $Component, $Type, $PID
    Add-Content -Path $Path -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue
    switch ($Type) {
        2 { Write-Host $Message -ForegroundColor Yellow }
        3 { Write-Host $Message -ForegroundColor Red }
        default { Write-Host $Message }
    }
}

function Test-Configured {
    param([string]$Value)
    return ($Value -and $Value.Trim() -ne '' -and $Value -notmatch '^\s*<Enter')
}
#endregion init ----------------------------------------------------------------


#region probes -----------------------------------------------------------------
function Get-ProxyInfo {
    # WinHTTP is what the SYSTEM account uses. The .NET system proxy usually mirrors
    # the interactive user (WinINET); reporting both makes the difference visible.
    $winhttp = (netsh winhttp show proxy 2>$null | Out-String).Trim()
    return $winhttp
}

function Get-ProxyForUrl {
    param([string]$Url)
    try {
        $sysProxy = [System.Net.WebRequest]::GetSystemWebProxy()
        $target   = [uri]$Url
        $via      = $sysProxy.GetProxy($target)
        if ($via -and $via.AbsoluteUri.TrimEnd('/') -ne $target.AbsoluteUri.TrimEnd('/') -and $via.Host -ne $target.Host) {
            return $via.Host + ":" + $via.Port
        }
        return "direct"
    } catch { return "unknown" }
}

function Test-EndpointConnectivity {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Url,
        [bool]$Required = $true
    )

    $uri  = [uri]$Url
    $targetHost = $uri.Host
    $port = $uri.Port
    if ($port -le 0) { if ($uri.Scheme -eq 'https') { $port = 443 } else { $port = 80 } }

    $result = [ordered]@{
        Name        = $Name
        Url         = $Url
        Host        = $targetHost
        Port        = $port
        Required    = $Required
        Proxy       = (Get-ProxyForUrl -Url $Url)
        DnsOk       = $false
        ResolvedIPs = ""
        TcpOk       = $false
        HttpStatus  = $null
        Pass        = $false
        DurationMs  = 0
        Detail      = ""
    }

    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    # 1) DNS
    try {
        $ips = [System.Net.Dns]::GetHostAddresses($targetHost) | ForEach-Object { $_.IPAddressToString }
        $result.DnsOk = ($ips.Count -gt 0)
        $result.ResolvedIPs = ($ips -join ", ")
    } catch {
        $result.Detail = "DNS resolution failed. Check internal DNS / split-DNS for $targetHost."
        $sw.Stop(); $result.DurationMs = $sw.ElapsedMilliseconds
        return [pscustomobject]$result
    }

    # 2) TCP
    $client = New-Object System.Net.Sockets.TcpClient
    try {
        $iar = $client.BeginConnect($targetHost, $port, $null, $null)
        $ok  = $iar.AsyncWaitHandle.WaitOne([int]($TimeoutSeconds * 1000), $false)
        if ($ok -and $client.Connected) { $client.EndConnect($iar); $result.TcpOk = $true }
    } catch { } finally { $client.Close() }

    if (-not $result.TcpOk) {
        $result.Detail = "TCP connect to $targetHost`:$port timed out or was refused. Firewall or proxy is blocking the path."
        $sw.Stop(); $result.DurationMs = $sw.ElapsedMilliseconds
        return [pscustomobject]$result
    }

    # 3) TLS + HTTP (any HTTP status = path open)
    try {
        $resp = Invoke-WebRequest -Uri $Url -Method GET -TimeoutSec $TimeoutSeconds -UseBasicParsing -ErrorAction Stop
        $result.HttpStatus = [int]$resp.StatusCode
        $result.Pass = $true
        $result.Detail = "OK. Endpoint reachable (HTTP $($result.HttpStatus))."
    } catch [System.Net.WebException] {
        $r = $_.Exception.Response
        if ($r -and ($r -is [System.Net.HttpWebResponse])) {
            $result.HttpStatus = [int]$r.StatusCode
            if ($result.HttpStatus -eq 407) {
                $result.Pass = $false
                $result.Detail = "Proxy returned 407 (authentication required). The SYSTEM account cannot authenticate to the proxy."
            } else {
                # 400/401/403/404/405 etc. still prove the network path is open.
                $result.Pass = $true
                $result.Detail = "OK. Endpoint reachable (HTTP $($result.HttpStatus); an auth error at this stage is expected without credentials)."
            }
        } else {
            $result.Pass = $false
            $msg = $_.Exception.Message
            if ($msg -match 'SSL|TLS|secure channel|trust') {
                $result.Detail = "TLS handshake failed: $msg. Likely SSL inspection or an untrusted intercepting certificate."
            } else {
                $result.Detail = "No HTTP response: $msg."
            }
        }
    } catch {
        $result.Pass = $false
        $result.Detail = "Request failed: $($_.Exception.Message)"
    }

    $sw.Stop(); $result.DurationMs = $sw.ElapsedMilliseconds
    return [pscustomobject]$result
}
#endregion probes --------------------------------------------------------------


#region build endpoint list ----------------------------------------------------
$endpoints = New-Object System.Collections.Generic.List[object]

# --- Core ingestion path ---
if ($LogAPIMode -eq "LogIngestionAPI") {
    # Entra sign-in endpoint (token). Tenant-scoped OpenID metadata is a clean reachability + tenant check.
    if (Test-Configured $TenantId) {
        $endpoints.Add([pscustomobject]@{ Name = "Entra sign-in (token)"; Url = "https://login.microsoftonline.com/$TenantId/v2.0/.well-known/openid-configuration"; Required = $true })
    } else {
        $endpoints.Add([pscustomobject]@{ Name = "Entra sign-in (token)"; Url = "https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration"; Required = $true })
    }
    if (Test-Configured $DceURI) {
        $endpoints.Add([pscustomobject]@{ Name = "Data Collection Endpoint (ingest)"; Url = $DceURI; Required = $true })
    } else {
        Write-CMTraceLog "DceURI is not set. Skipping the ingestion endpoint test. Fill in `$DceURI to test it." -Type 2
    }
    if (Test-Configured $BrokerUrl) {
        $endpoints.Add([pscustomobject]@{ Name = "Entra Token Broker"; Url = $BrokerUrl; Required = $true })
    }
} elseif ($LogAPIMode -eq "DataCollectorAPI") {
    if (Test-Configured $CustomerId) {
        $endpoints.Add([pscustomobject]@{ Name = "Log Analytics ingestion (legacy)"; Url = "https://$CustomerId.ods.opinsights.azure.com"; Required = $true })
    } else {
        Write-CMTraceLog "CustomerId (workspace ID) is not set. Skipping the legacy ingestion test." -Type 2
    }
} else {
    Write-CMTraceLog "Unknown `$LogAPIMode '$LogAPIMode'. Expected 'LogIngestionAPI' or 'DataCollectorAPI'." -Type 3
}

# --- Microsoft 365 release data (only when the collector gathers it) ---
if ($CollectMicrosoft365) {
    $endpoints.Add([pscustomobject]@{ Name = "Microsoft 365 release data";  Url = "https://mrodevicemgr.officeapps.live.com/mrodevicemgrsvc/api/v2/C2RReleaseData"; Required = $false })
    $endpoints.Add([pscustomobject]@{ Name = "Office CDN release info";      Url = "https://clients.config.office.net/releases/v1.0/OfficeReleases";                Required = $false })
}

# --- Warranty APIs (only when the collector gathers warranty) ---
if ($CollectWarranty) {
    $mfr = ""
    try { $mfr = (Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop).Manufacturer } catch { }
    $mfr = "$mfr".ToLower()

    $vendors = @{
        Dell   = "https://apigtwb2c.us.dell.com"
        HP     = "https://warranty.api.hp.com"
        Lenovo = "http://supportapi.lenovo.com"   # Lenovo warranty API is HTTP (port 80)
        Getac  = "https://api.getac.us"
    }

    $selected = @()
    if ($TestAllWarrantyVendors) {
        $selected = $vendors.Keys
    } elseif ($mfr -match 'dell') {                 $selected = @("Dell")
    } elseif ($mfr -match 'hp|hewlett') {           $selected = @("HP")
    } elseif ($mfr -match 'lenovo') {               $selected = @("Lenovo")
    } elseif ($mfr -match 'getac') {                $selected = @("Getac")
    } else {
        Write-CMTraceLog "Manufacturer '$mfr' did not match a known warranty vendor. Testing all vendors." -Type 2
        $selected = $vendors.Keys
    }

    foreach ($v in $selected) {
        $endpoints.Add([pscustomobject]@{ Name = "$v warranty API"; Url = $vendors[$v]; Required = $false })
    }
}
#endregion build endpoint list -------------------------------------------------


#region run --------------------------------------------------------------------
Write-CMTraceLog "==================================================================="
Write-CMTraceLog "PowerStacks Inventory connectivity test"
Write-CMTraceLog "Device: $env:COMPUTERNAME   User context: $([Security.Principal.WindowsIdentity]::GetCurrent().Name)"
Write-CMTraceLog "LogAPIMode: $LogAPIMode   Endpoints to test: $($endpoints.Count)"
Write-CMTraceLog "WinHTTP proxy (SYSTEM uses this):"
foreach ($pl in (Get-ProxyInfo -split "`r?`n")) { if ($pl.Trim()) { Write-CMTraceLog "  $($pl.Trim())" } }
Write-CMTraceLog "-------------------------------------------------------------------"

$results = New-Object System.Collections.Generic.List[object]
foreach ($e in $endpoints) {
    Write-CMTraceLog "Testing: $($e.Name)  ->  $($e.Url)"
    $r = Test-EndpointConnectivity -Name $e.Name -Url $e.Url -Required $e.Required
    $results.Add($r)

    $status = if ($r.Pass) { "PASS" } else { "FAIL" }
    $type   = if ($r.Pass) { 1 } elseif ($r.Required) { 3 } else { 2 }
    Write-CMTraceLog ("  [{0}] {1}  (proxy: {2}, dns: {3}, tcp: {4}, http: {5}, {6} ms)" -f `
        $status, $r.Name, $r.Proxy, $r.DnsOk, $r.TcpOk, $r.HttpStatus, $r.DurationMs) -Type $type
    Write-CMTraceLog ("      {0}" -f $r.Detail) -Type $type
    if ($r.ResolvedIPs) { Write-CMTraceLog ("      resolved: {0}" -f $r.ResolvedIPs) }
}

# --- Optional: prove the sign-in path with real credentials (no data uploaded) ---
if ($TestTokenAcquisition -and $LogAPIMode -eq "LogIngestionAPI" -and
    (Test-Configured $TenantId) -and (Test-Configured $ClientId) -and (Test-Configured $ClientSecret)) {
    Write-CMTraceLog "-------------------------------------------------------------------"
    Write-CMTraceLog "Testing sign-in (client credentials -> access token for monitor.azure.com)"
    try {
        $body = @{
            client_id     = $ClientId
            client_secret = $ClientSecret
            scope         = "https://monitor.azure.com//.default"
            grant_type    = "client_credentials"
        }
        $tok = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
            -Body $body -ContentType "application/x-www-form-urlencoded" -TimeoutSec $TimeoutSeconds -ErrorAction Stop
        if ($tok.access_token) {
            $script:AccessToken = $tok.access_token
            Write-CMTraceLog "  [PASS] Acquired an access token. Sign-in and app credentials are good." -Type 1
        } else {
            Write-CMTraceLog "  [FAIL] Token endpoint responded but returned no access_token." -Type 3
            $script:PreflightFailures.Add("Sign-in: token endpoint returned no access_token.")
        }
    } catch {
        Write-CMTraceLog "  [FAIL] Token request failed: $($_.Exception.Message)" -Type 3
        Write-CMTraceLog "         (A network failure here points at the proxy/firewall; an AADSTS error points at the app registration or secret.)" -Type 2
        $script:PreflightFailures.Add("Sign-in: token request failed ($($_.Exception.Message)).")
    }
}

# --- Ingestion dry-run: reproduce the collector's actual upload (the only check that catches a 404) ---
if ($TestIngestion -and $LogAPIMode -eq "LogIngestionAPI" -and $script:AccessToken -and
    (Test-Configured $DceURI) -and (Test-Configured $DcrImmutableId) -and (Test-Configured $StreamName)) {
    Write-CMTraceLog "-------------------------------------------------------------------"
    Write-CMTraceLog "Ingestion dry-run: POST to the real DCR stream (the collector's actual upload path)"
    $ingestUri = "$($DceURI.TrimEnd('/'))/dataCollectionRules/$DcrImmutableId/streams/$StreamName" + "?api-version=2023-01-01"
    Write-CMTraceLog "  URL: $ingestUri"
    # An empty JSON array routes through the exact DCR + stream + permission checks the collector hits,
    # but writes no data. 404 = the DCR immutable id / stream name / DCE association is wrong (the
    # collector's error); 403 = missing 'Monitoring Metrics Publisher' role; 400/204 = the target
    # resolved correctly (empty body), so the ingestion configuration is good.
    try {
        $resp = Invoke-WebRequest -Uri $ingestUri -Method POST -Headers @{ Authorization = "Bearer $script:AccessToken" } `
            -ContentType "application/json" -Body "[]" -TimeoutSec $TimeoutSeconds -UseBasicParsing -ErrorAction Stop
        Write-CMTraceLog "  [PASS] Ingestion accepted (HTTP $([int]$resp.StatusCode)). The DCR, stream, and permissions are correct." -Type 1
    } catch [System.Net.WebException] {
        $r = $_.Exception.Response
        $code = if ($r -and ($r -is [System.Net.HttpWebResponse])) { [int]$r.StatusCode } else { $null }
        $bodyText = ""
        if ($r) { try { $rs = New-Object System.IO.StreamReader($r.GetResponseStream()); $bodyText = $rs.ReadToEnd(); $rs.Close() } catch {} }
        switch ($code) {
            400 { Write-CMTraceLog "  [PASS] The DCR and stream resolved (HTTP 400 for the empty probe body is expected). The ingestion path is configured correctly." -Type 1 }
            401 { Write-CMTraceLog "  [FAIL] HTTP 401. The access token was rejected (wrong tenant/authority or audience). The upload scope must be https://monitor.azure.com//.default." -Type 3
                  $script:PreflightFailures.Add("Ingestion: 401 (token rejected).") }
            403 { Write-CMTraceLog "  [FAIL] HTTP 403. Sign-in works but the app cannot write to this DCR. Grant the app the 'Monitoring Metrics Publisher' role on the DCR (or DCE), then wait a few minutes." -Type 3
                  $script:PreflightFailures.Add("Ingestion: 403 (app lacks Monitoring Metrics Publisher on the DCR).") }
            404 { Write-CMTraceLog "  [FAIL] HTTP 404. This is the collector's upload error. The DCR immutable id or stream does not resolve on this DCE. Verify: (1) `$DcrImmutableId is the DCR's IMMUTABLE ID (not its name or resource id); (2) the DCR declares a stream named '$StreamName'; (3) the DCR is associated with the DCE at `$DceURI (they must be from the same deployment)." -Type 3
                  $script:PreflightFailures.Add("Ingestion: 404 (DCR immutable id / stream / DCE association is wrong).") }
            407 { Write-CMTraceLog "  [FAIL] HTTP 407. The SYSTEM account cannot authenticate to the proxy." -Type 3
                  $script:PreflightFailures.Add("Ingestion: 407 (proxy authentication).") }
            default { Write-CMTraceLog "  [FAIL] HTTP $code on the ingestion POST. $($_.Exception.Message)" -Type 3
                  $script:PreflightFailures.Add("Ingestion: HTTP $code.") }
        }
        if ($bodyText) { Write-CMTraceLog "         Azure response: $bodyText" -Type 2 }
    } catch {
        Write-CMTraceLog "  [FAIL] Ingestion dry-run failed: $($_.Exception.Message)" -Type 3
        $script:PreflightFailures.Add("Ingestion: $($_.Exception.Message)")
    }
} elseif ($TestIngestion -and $LogAPIMode -eq "LogIngestionAPI" -and (Test-Configured $DceURI) -and -not (Test-Configured $DcrImmutableId)) {
    Write-CMTraceLog "-------------------------------------------------------------------"
    Write-CMTraceLog "Ingestion dry-run SKIPPED: set `$DcrImmutableId (and `$StreamName) to test the real upload path. Without it this test CANNOT catch an upload 404." -Type 2
}
#endregion run -----------------------------------------------------------------


#region summary ----------------------------------------------------------------
$failedRequired = @($results | Where-Object { $_.Required -and -not $_.Pass })
$failedOptional = @($results | Where-Object { -not $_.Required -and -not $_.Pass })
$passed         = @($results | Where-Object { $_.Pass })

$preflight = @($script:PreflightFailures)

Write-CMTraceLog "==================================================================="
Write-CMTraceLog ("SUMMARY: {0} passed, {1} required failure(s), {2} sign-in/ingestion failure(s), {3} optional failure(s)." -f `
    $passed.Count, $failedRequired.Count, $preflight.Count, $failedOptional.Count) -Type $(if ($failedRequired.Count -or $preflight.Count) { 3 } else { 1 })

if ($failedRequired.Count) {
    Write-CMTraceLog "Required endpoints that FAILED (these will break inventory upload):" -Type 3
    foreach ($f in $failedRequired) { Write-CMTraceLog ("  - {0} ({1}): {2}" -f $f.Name, $f.Host, $f.Detail) -Type 3 }
}
if ($preflight.Count) {
    Write-CMTraceLog "Sign-in / ingestion checks that FAILED (these break inventory upload even when every endpoint is reachable):" -Type 3
    foreach ($p in $preflight) { Write-CMTraceLog ("  - {0}" -f $p) -Type 3 }
}
if ($failedOptional.Count) {
    Write-CMTraceLog "Optional endpoints that failed (inventory still uploads; that data will be missing):" -Type 2
    foreach ($f in $failedOptional) { Write-CMTraceLog ("  - {0} ({1}): {2}" -f $f.Name, $f.Host, $f.Detail) -Type 2 }
}
Write-CMTraceLog "Full log: $CMLog"
Write-CMTraceLog "==================================================================="

# Concise machine-readable line for Intune's detection-output column.
if ($failedRequired.Count -eq 0 -and $preflight.Count -eq 0 -and $failedOptional.Count -eq 0) {
    Write-Output "PowerStacksInventoryConnectivity=OK (all checks passed)"
    exit 0
} elseif ($failedRequired.Count -eq 0 -and $preflight.Count -eq 0) {
    Write-Output ("PowerStacksInventoryConnectivity=WARN (optional failing: {0})" -f (($failedOptional | ForEach-Object { $_.Host }) -join ','))
    exit 0
} else {
    $reasons = @()
    if ($failedRequired.Count) { $reasons += (($failedRequired | ForEach-Object { $_.Host }) -join ',') }
    if ($preflight.Count)      { $reasons += 'ingestion/sign-in' }
    Write-Output ("PowerStacksInventoryConnectivity=FAIL ({0})" -f ($reasons -join '; '))
    exit 1
}
#endregion summary -------------------------------------------------------------
