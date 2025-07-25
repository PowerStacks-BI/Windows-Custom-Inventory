<#
.SYNOPSIS
    Comprehensive inventory and warranty collection script for Intune-managed Windows 10/11 devices.
    Gathers hardware, software, driver, monitor, disk, battery, Microsoft 365, and warranty data, 
    then uploads results to Azure Log Analytics for centralized reporting.

.DESCRIPTION
    This script is designed for deployment via Intune to Windows 10/11 endpoints. It collects detailed inventory data including:
      - Hardware specs (CPU, RAM, disks, chassis)
      - Battery health
      - Monitor (LCD) information including serial numbers and date of manufacture
      - Installed Win32 and UWP applications
      - Installed and available drivers
      - Microsoft 365 update channel and compliance
      - Device warranty status for Dell, HP, Lenovo, and Getac (via vendor APIs)
    Warranty lookups are cached locally to minimize API calls and can be forced to refresh as needed. 
    All collected data is compressed, base64-encoded, and securely uploaded to Azure Log Analytics.
    The script is modular, allowing granular control over which inventory types are collected via variables.

.APIKEYS
    To enable warranty lookups, you must obtain API credentials from each vendor:
      - Dell: Register for a Dell TechDirect account (https://techdirect.dell.com/), request API access, and generate a Client ID and Secret for the Warranty API.
      - Lenovo: Request access to the Lenovo Warranty API at https://supportapi.lenovo.com or through your Lenovo rep. You will receive a Client ID for authentication.
      - HP: Apply for HP’s Warranty API access at https://developer.hp.com/ or through your HP rep. After approval, create an application to obtain a Client ID and Secret.
      - Getac: Contact Getac support or your Getac account representative to request API access and credentials for warranty lookups.


.PARAMETER WarrantyMaxCacheAgeDays
    [int] Maximum number of days before cached warranty data is considered stale. Default: 180.

.PARAMETER WarrantyForceRefresh
    [switch] When set to $true, ignores cached warranty data and forces a fresh API lookup.

.PARAMETER $CollectDeviceInventory
    [bool] Set to $true to collect device hardware inventory. Default: $true.

.PARAMETER $CollectAppInventory
    [bool] Set to $true to collect installed application inventory. Default: $true.

.PARAMETER $CollectDriverInventory
    [bool] Set to $true to collect installed and available driver inventory. Default: $true.

.PARAMETER $CollectUWPInventory
    [bool] Set to $true to collect UWP (AppX) application inventory. Default: $false.

.PARAMETER $CollectMicrosoft365
    [bool] Set to $true to collect Microsoft 365 update and compliance data. Default: $true.

.PARAMETER $CollectWarranty
    [bool] Set to $true to collect device warranty information. Default: $false.

.PARAMETER $RemoveBuiltInMonitors
    [bool] Set to $true to exclude built-in monitors from monitor inventory. Default: $false.

.PARAMETER $InventoryDateFormat
    [string] Date format string for inventory timestamps. Default: "MM-dd HH:mm".

.PARAMETER $CustomerId
    [string] Azure Log Analytics Workspace ID.

.PARAMETER $SharedKey
    [string] Azure Log Analytics Primary Key.

.PARAMETER $WarrantyDellClientID, $WarrantyDellClientSecret
    [string] Dell API credentials for warranty lookup.

.PARAMETER $WarrantyLenovoClientID
    [string] Lenovo API credential for warranty lookup.

.PARAMETER $WarrantyHPClientID, $WarrantyHPClientSecret
    [string] HP API credentials for warranty lookup.

.PARAMETER $TimeStampField
    [string] Optional. Specifies the timestamp field for Log Analytics ingestion. Leave blank unless required.

.NOTES
    Author: John Marcum (PJM)
    Date: June 9, 2025
    Contact: https://x.com/MEM_MVP

.VERSION HISTORY
12 - June 9, 2025
- Added HP warranty support and warranty caching
- Changed warranty date fields to datetime
- Added $CollectUWPInventory toggle
- Fixed driver inventory bug with Get-Package provider
- Added OS install date

########### LEGAL DISCLAIMER ###########
    This script is provided "as is" without warranty of any kind, either express or implied. 
    Use at your own risk. Test thoroughly before deploying in production environments.
#>


#region initialize
# Enable TLS 1.2 support
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Replace with your Log Analytics Workspace ID
$CustomerId = "<Enter Your Log Analytics Workspace ID>"   

# Replace with your Primary Key
$SharedKey = "<Enter Your Log Analytics Workspace Primary Key>"

#Control if you want to collect Device, Win32 App, UWP App, and Driver Inventory. (True = Collect)
$CollectDeviceInventory = $true
$CollectAppInventory = $true
$CollectDriverInventory = $true
$CollectUWPInventory = $False # Set to true to collect UWP (modern app) inventory.

#Sub-Control under Device Inventory
$CollectMicrosoft365 = $true
$CollectWarranty = $False # Set to true to collect warranty data

#Warranty keys
$WarrantyDellClientID = "<Enter Your Dell Client ID>"
$WarrantyDellClientSecret = "<Enter Your Dell Client Secret>"
$WarrantyLenovoClientID = "<Enter Your Lenovo Client ID>"
$WarrantyHPClientID = "<Enter Your HP Client ID>"
$WarrantyHPClientSecret = "<Enter Your HP Client Secret>"  # Make note of expiration date!


# Warranty cache settings
[int]$WarrantyMaxCacheAgeDays = 180 # The max age of the .json file which caches warranty data. 
[switch]$WarrantyForceRefresh = $false # Set to true to ignore the json and pull data from the API.

# You can use an optional field to specify the timestamp from the data. If the time field is not specified, Azure Monitor assumes the time is the message ingestion time
# DO NOT DELETE THIS VARIABLE. Recommended keep this blank.
$TimeStampField = ""

#Control if you want to remove BuiltIn Monitors (true = Remove)
$RemoveBuiltInMonitors = $false

#Inventory Date Format (sample: "MM-dd HH:mm", "dd-MM HH:mm")
$InventoryDateFormat = "MM-dd HH:mm"

#endregion initialize

# Start Logging
$Now = Get-Date -Format "yyyy-MM-dd_HHmm"
$logPath = "C:\Windows\Logs\Intune_Inventory_$Now.log"
Start-Transcript -Path $logPath -Append


#region functions

# Function to get all Installed Application
function Get-InstalledApplications() {
    <#
.SYNOPSIS
    Retrieves installed Win32 applications for a specified user.
.DESCRIPTION
    Scans registry locations for installed Win32 applications, including 32-bit and 64-bit entries, and returns details such as name, version, publisher, and install date.
.PARAMETER UserSid
    The SID of the user whose HKU registry hive should be scanned for per-user applications.
.OUTPUTS
    PSCustomObject[] representing installed applications.
#>
    param(
        [string]$UserSid
    )
 
    New-PSDrive -PSProvider Registry -Name "HKU" -Root HKEY_USERS | Out-Null
    $regpath = @("HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*")
    $regpath += "HKU:\$UserSid\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
    if (-not ([IntPtr]::Size -eq 4)) {
        $regpath += "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
        $regpath += "HKU:\$UserSid\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    }
    $propertyNames = 'DisplayName', 'DisplayVersion', 'Publisher', 'UninstallString', 'InstallDate'
    $Apps = Get-ItemProperty $regpath -Name $propertyNames -ErrorAction SilentlyContinue | . { process { if ($_.DisplayName) { $_ } } } | Select-Object DisplayName, DisplayVersion, Publisher, UninstallString, InstallDate, PSPath | Sort-Object DisplayName
 
    # Convert InstallDate string to DateTime and format as DD/MM/YYYY, handling empty InstallDate
    foreach ($app in $Apps) {
        if (![string]::IsNullOrWhiteSpace($app.InstallDate)) {
            $parsedDate = [DateTime]::MinValue
            if ([DateTime]::TryParseExact($app.InstallDate, 'yyyyMMdd', [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$parsedDate)) {
                $app.InstallDate = $parsedDate.ToString('dd-MM-yyyy')
            }
            else {
                # Date parsing failed, handle accordingly (e.g., set to null or a default value)
                $app.InstallDate = $null
            }
        }
        else {
            # Empty InstallDate string, handle accordingly (e.g., set to null or a default value)
            $app.InstallDate = $null
        }
    }
 
    Remove-PSDrive -Name "HKU" | Out-Null
    Return $Apps
}

# Function to get deduplicated Appx Installed Applications (UWP)
Function Get-AppxInstalledApplications() {
    <#
.SYNOPSIS
    Retrieves deduplicated list of installed UWP (AppX) applications for all users.
.DESCRIPTION
    Handles known issues with Get-AppxPackage on Windows 11 24H2, loads required assemblies, and returns UWP app details including name, version, and publisher.
.OUTPUTS
    PSCustomObject[] representing installed UWP applications.
#>    
    
    # Fix for issue which breaks Get-AppxPackage in Win11 24H2
    # This is a known bug in 24H2. 
    # Remove the fix once MS fixes the issue. Until this UWP app inventory may or may not work in 24H2
    Add-Type -AssemblyName "System.EnterpriseServices"
    $publish = [System.EnterpriseServices.Internal.Publish]::new()

    $dlls = @(
        'System.Memory.dll',
        'System.Numerics.Vectors.dll',
        'System.Runtime.CompilerServices.Unsafe.dll',
        'System.Security.Principal.Windows.dll'
    )

    foreach ($dll in $dlls) {
        $dllPath = "$env:SystemRoot\\System32\\WindowsPowerShell\\v1.0\\$dll"
        $fileName = [System.IO.Path]::GetFileNameWithoutExtension($dll)

        $gacPath = "$env:windir\\Microsoft.NET\\assembly"
        $existsInGAC = Get-ChildItem -Recurse -Path $gacPath -Filter "$fileName.dll" -ErrorAction SilentlyContinue | Where-Object {
            $_.FullName -match [regex]::Escape($fileName)
        }

        if (-not $existsInGAC) {
            Write-Host "$dll not found in GAC. Installing..."
            $publish.GacInstall($dllPath)
        }
    }
    # End the fix


    # Get the apps
    try {
        $ErrorActionPreference = 'Stop'
        $appPackages = Get-AppxPackage -AllUsers
    }
    catch {
        Write-Warning "Failed to retrieve Appx packages: $($_.Exception.Message)"
        $appPackages = @() # or $null if you prefer
    }
    finally {
        $ErrorActionPreference = 'Continue' # Reset to default if needed
    }

    $uwpAppList = @()

    # Process only the installed apps
    foreach ($pkg in $appPackages) {
        if ($pkg.PackageUserInformation | Where-Object { $_.InstallState -eq 'Installed' }) {
            $publisher = $null
            try {
                $manifest = Get-AppxPackageManifest -Package $pkg.PackageFullName
                $publisher = $manifest.Package.Properties.PublisherDisplayName
            }
            catch {}

            $uwpApp = New-Object -TypeName PSObject
            $uwpApp | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $pkg.Name -Force
            $uwpApp | Add-Member -MemberType NoteProperty -Name "DisplayVersion" -Value $pkg.Version.ToString() -Force
            $uwpApp | Add-Member -MemberType NoteProperty -Name "Publisher" -Value $publisher -Force
            $uwpApp | Add-Member -MemberType NoteProperty -Name "AppType" -Value "UWP" -Force
            $uwpAppList += $uwpApp
        }
    }

    $dedupedUwpApps = $uwpAppList | Sort-Object DisplayName, DisplayVersion -Unique
    return $dedupedUwpApps
}



# Function to get Office update infomation
function Get-Microsoft365 {
    <#
.SYNOPSIS
    Retrieves Microsoft 365 (Office Click-to-Run) update and compliance information.
.DESCRIPTION
    Determines installed Office version, update channel, latest release, end of support, and other compliance details by querying registry and Microsoft APIs.
.OUTPUTS
    PSCustomObject with Office version, channel, release, and support information.
#>
    $IsC2R = Test-Path 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun'
    if (-not $IsC2R) { Write-Output "Not Click-to-Run Office"; return $null }

    try {
        $ConfigPath = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
        $OfficeVersion = [version](Get-ItemProperty -Path $ConfigPath -ErrorAction Stop | Select-Object -ExpandProperty VersionToReport)
        $OfficeProductIds = (Get-ItemProperty -Path $ConfigPath -ErrorAction Stop | Select-Object -ExpandProperty ProductReleaseIds)
        $OfficeVersionString = $OfficeVersion.ToString()
        Write-Output "Installed Version: $OfficeVersionString"
    }
    catch {
        Write-Output "Failed to read Office configuration: $_"
        return $null
    }

    $IsM365 = ($OfficeProductIds -like '*O365*') -or ($OfficeProductIds -like '*M365*')

    $Channels = @(
        @{ GUID = '492350f6-3a01-4f97-b9c0-c7c6ddf67d60'; Name = 'Monthly'; GPO = 'Current' }
        @{ GUID = '64256afe-f5d9-4f86-8936-8840a6a4f5be'; Name = 'Monthly (Preview)'; GPO = 'FirstReleaseCurrent' }
        @{ GUID = '55336b82-a18d-4dd6-b5f6-9e5095c314a6'; Name = 'Monthly Enterprise'; GPO = 'MonthlyEnterprise' }
        @{ GUID = '7ffbc6bf-bc32-4f92-8982-f9dd17fd3114'; Name = 'Semi-Annual'; GPO = 'Deferred' }
        @{ GUID = 'b8f9b850-328d-4355-9145-c59439a0c4cf'; Name = 'Semi-Annual (Preview)'; GPO = 'FirstReleaseDeferred' }
        @{ GUID = '5440fd1f-7ecb-4221-8110-145efaa6372f'; Name = 'Beta'; GPO = 'InsiderFast' }
    )

    $OfficeChannel = @{ Name = $null }
    $UpdateBranch = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\OfficeUpdate' -ErrorAction SilentlyContinue).UpdateBranch
    if ($UpdateBranch) {
        $OfficeChannel = $Channels | Where-Object { $_.GPO -eq $UpdateBranch }
        Write-Output "Update channel from GPO: $($UpdateBranch): $($OfficeChannel.Name)"
    }
    else {
        $CDNBaseUrl = (Get-ItemProperty -Path $ConfigPath -ErrorAction SilentlyContinue).CDNBaseUrl
        if ($CDNBaseUrl) {
            try {
                $Uri = [System.Uri]$CDNBaseUrl
                $Guid = $Uri.Segments[2].TrimEnd('/')
                $OfficeChannel = $Channels | Where-Object { $_.GUID -eq $Guid }
                Write-Output "Update channel from CDN GUID: $Guid → $($OfficeChannel.Name)"
            }
            catch {
                Write-Output "Failed to parse CDNBaseUrl for channel"
            }
        }
    }

    $ChannelPathMap = @{
        'Monthly'               = 'Monthly'
        'Monthly (Preview)'     = 'MonthlyPreview'
        'Monthly Enterprise'    = 'MonthlyEnterpriseChannel'
        'Semi-Annual'           = 'SAC'
        'Semi-Annual (Preview)' = 'SACT'
        'Beta'                  = 'Beta'
    }

    if ($OfficeProductIds -like '*2019Volume*') {
        $CDNChannel = 'LTSB'
        Write-Output "Legacy Office 2019 detected"
    }
    elseif ($OfficeProductIds -like '*2021Volume*') {
        $CDNChannel = 'LTSB2021'
        Write-Output "Legacy Office 2021 detected"
    }
    else {
        $CDNChannel = $ChannelPathMap[$OfficeChannel.Name]
    }

    Write-Output "CDN channel path: $CDNChannel"


    ### Defaults
    $LatestReleaseType = $null
    $LatestReleaseVersion = $null
    $EndOfSupportDate = $null
    $ReleaseDate = $null
    $ReleaseID = $null

    try {
        $C2RData = Invoke-RestMethod -Uri 'https://mrodevicemgr.officeapps.live.com/mrodevicemgrsvc/api/v2/C2RReleaseData' -Method GET -ErrorAction Stop
        $ReleaseMatch = $C2RData | Where-Object { $_.availableBuild -eq $OfficeVersionString }


        if (-not $ReleaseMatch) {
            $ReleaseMatch = $C2RData |
                Where-Object { $_.availableBuild -like "$($OfficeVersion.Major).$($OfficeVersion.Minor).*" } |
                Sort-Object availableBuild -Descending |
                Select-Object -First 1
            Write-Output "Fuzzy match used for release data"
        }

        # If ReleaseMatch is an array, pick the first match
        if ($ReleaseMatch -is [array]) {
            Write-Output "`n=== Raw C2R Match Array ==="
            $ReleaseMatch | ConvertTo-Json -Depth 5 | Out-String | Write-Output
            Write-Output "=== End Raw C2R Match Array ===`n"

            if ($ReleaseMatch.Count -gt 0) {
                Write-Output "Multiple matches found. Using first: $($ReleaseMatch[0].availableBuild)"
                $ReleaseMatch = $ReleaseMatch[0]
            }
            else {
                Write-Output "ReleaseMatch was empty array"
                $ReleaseMatch = $null
            }
        }

        if ($ReleaseMatch) {
            Write-Output "`n=== Raw Single C2R Match ==="
            $ReleaseMatch | ConvertTo-Json -Depth 5 | Out-String | Write-Output
            Write-Output "=== End Raw C2R Match ===`n"
            Write-Output "C2R Match found: $($ReleaseMatch.availableBuild)"

            $LatestReleaseVersion = "$($ReleaseMatch.availableBuild)"
            $LatestReleaseType = "$($ReleaseMatch.type)"
            $ReleaseDate = "$($ReleaseMatch.updatedTimeUtc)"
            $ReleaseID = ($ReleaseMatch.forkName -split '-')[0]

            if ($ReleaseMatch.endOfSupportDate -and $ReleaseMatch.endOfSupportDate -ne '0001-01-01T00:00:00Z') {
                $EndOfSupportDate = "$($ReleaseMatch.endOfSupportDate)"
            }
            else {
                Write-Output "EndOfSupportDate not found in C2R fallback needed"
            }
        }
        else {
            Write-Output "No C2R match found"
        }
    }
    catch {
        Write-Warning "C2R API failed: $_"
    }


    if ($CDNChannel) {
        try {
            $CDNUrl = "https://clients.config.office.net/releases/v1.0/LatestRelease/$CDNChannel"
            $CDNResp = Invoke-RestMethod -Uri $CDNUrl -Method GET -ErrorAction Stop

            if (-not $EndOfSupportDate -and $CDNResp.endOfSupportDate -ne '0001-01-01T00:00:00Z') {
                $EndOfSupportDate = $CDNResp.endOfSupportDate
                Write-Output "EndOfSupportDate pulled from CDN: $EndOfSupportDate"
            }

            if (-not $ReleaseDate -and $CDNResp.availabilityDate) {
                $ReleaseDate = $CDNResp.availabilityDate
                Write-Output "ReleaseDate pulled from CDN"
            }

            if (-not $LatestReleaseVersion -and $CDNResp.buildVersion.buildVersionString) {
                $LatestReleaseVersion = $CDNResp.buildVersion.buildVersionString
                Write-Output "ReleaseVersion pulled from CDN"
            }

            if (-not $ReleaseID -and $CDNResp.releaseVersion) {
                $ReleaseID = $CDNResp.releaseVersion
                Write-Output "ReleaseID pulled from CDN"
            }

            if (-not $LatestReleaseType -or $LatestReleaseType -eq 'Default') {
                $ReleaseTypes = @{ 1 = 'Feature Update'; 2 = 'Quality Update'; 3 = 'Security Update' }
                $LatestReleaseType = $ReleaseTypes[$CDNResp.releaseType]
                if (-not $LatestReleaseType -and $CDNResp.releaseType -ne $null) {
                    $LatestReleaseType = "$($CDNResp.releaseType)"  # fallback to raw value
                }
                Write-Output "ReleaseType pulled from CDN: $LatestReleaseType"
            }

        }
        catch {
            Write-Warning "CDN API failed: $_"
        }
    }

    Write-Output "FINAL: Installed=$OfficeVersionString | Channel=$($OfficeChannel.Name) | Release=$LatestReleaseVersion | Type=$LatestReleaseType | EoS=$EndOfSupportDate"

    return [pscustomobject]@{
        InstalledVersion     = $OfficeVersionString
        UpdateChannel        = $OfficeChannel.Name
        LatestReleaseType    = $LatestReleaseType
        LatestReleaseVersion = $LatestReleaseVersion
        EndOfSupportDate     = $EndOfSupportDate
        ReleaseDate          = $ReleaseDate
        ReleaseID            = $ReleaseID
    }
}

# Function to get Installed Drivers
<#
Feel free to edit the query collect more or less drivers. - PJM
#>
function Get-InstalledDrivers() {
    <#
.SYNOPSIS
    Retrieves installed and available driver information.
.DESCRIPTION
    Collects installed drivers from Win32_PnPSignedDriver and available driver updates from Windows Update, returning a unified list of driver details.
.OUTPUTS
    PSCustomObject[] representing installed and available drivers.
#>
    Write-Output "Begin getting installed drivers"
    # Get PnP signed drivers
    Write-Output "Get PnP signed drivers"
    $PNPSigned_Drivers = Get-CimInstance -ClassName Win32_PnPSignedDriver | Where-Object {
    ($_.Manufacturer -ne "Microsoft") -and 
    ($_.DriverProviderName -ne "Microsoft") -and 
    ($_.DeviceName -ne $null)
    } | Select-Object DeviceName, DriverVersion, DriverDate, DeviceClass, DeviceID, HardwareID, Manufacturer, InfName, Location, Description, DriverProviderName
    $PNPSigned_Drivers

    # Simulate installed driver packages based on PnP data
    Write-Output "Get installed drivers from PnP data"
    $InstalledDrivers = $PNPSigned_Drivers | Where-Object {
        $_.DriverProviderName -notlike "*Microsoft*" -and $_.DeviceName
    }
    $InstalledDrivers

    # Get optional updates
    Write-Output "Get optional updates"
    $updateSession = New-Object -ComObject Microsoft.Update.Session
    $updateSearcher = $updateSession.CreateUpdateSearcher()
    $searchResult = $updateSearcher.Search("IsInstalled=0 AND Type='Driver'")
    $OptionalWUList = @()

    Write-Output "Optional update count: $($searchResult.Updates.Count)"

    If ($searchResult.Updates.Count -gt 0) {
        For ($i = 0; $i -lt $searchResult.Updates.Count; $i++) {
            $update = $searchResult.Updates.Item($i)
            $OptionalWUList += [PSCustomObject]@{
                WUName             = $update.Title
                DriverName         = $update.DriverModel
                DriverVersion      = $null
                DriverReleaseDate  = if ($update.DriverVerDate) { 
                    $update.DriverVerDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffffffZ") 
                }
                else { 
                    $null 
                }
                DriverClass        = if ($update.DriverClass) { 
                    $update.DriverClass.ToUpper() 
                }
                else { 
                    $null 
                }
                DriverID           = $null
                DriverHardwareID   = $update.DriverHardwareID
                DriverManufacturer = $update.DriverManufacturer
                DriverInfName      = $null
                DriverLocation     = $null
                DriverDescription  = $update.Description
                DriverProvider     = $update.DriverProvider
                DriverPublishedOn  = if ($update.LastDeploymentChangeTime) {
                    $update.LastDeploymentChangeTime.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffffffZ")
                }
                else {
                    $null
                }
                DriverStatus       = "Optional"
            }
        }
    }

    # Link installed drivers
    $LinkedDrivers = foreach ($installedDriver in $InstalledDrivers) {
        $versionFromName = $installedDriver.DriverVersion
        $matchingDriver = $PNPSigned_Drivers | Where-Object {
            $_.DriverVersion -eq $versionFromName
        } | Select-Object -First 1

        if ($matchingDriver) {
            [PSCustomObject]@{
                WUName             = $null
                DriverName         = $matchingDriver.DeviceName
                DriverVersion      = $matchingDriver.DriverVersion
                DriverReleaseDate  = if ($matchingDriver.DriverDate) {
                    $matchingDriver.DriverDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffffffZ")
                }
                else {
                    $null
                }
                DriverClass        = $matchingDriver.DeviceClass
                DriverID           = $matchingDriver.DeviceID
                DriverHardwareID   = $matchingDriver.HardwareID
                DriverManufacturer = $matchingDriver.Manufacturer
                DriverInfName      = $matchingDriver.InfName
                DriverLocation     = $matchingDriver.Location
                DriverDescription  = $matchingDriver.Description
                DriverProvider     = $matchingDriver.DriverProviderName
                DriverPublishedOn  = $null
                DriverStatus       = "Installed"
            }
        }
    }

    # Add unmatched installed drivers
    $matchedVersions = $LinkedDrivers | Where-Object { $_.DriverVersion } | Select-Object -ExpandProperty DriverVersion
    Start-Sleep -Seconds 1
    $unmatchedDrivers = $PNPSigned_Drivers | Where-Object { $matchedVersions -notcontains $_.DriverVersion }

    # Combine all drivers
    $LinkedDrivers = @(
        $LinkedDrivers
        foreach ($driver in $unmatchedDrivers) {
            [PSCustomObject]@{
                WUName             = $null
                DriverName         = $driver.DeviceName
                DriverVersion      = $driver.DriverVersion
                DriverReleaseDate  = if ($driver.DriverDate) { 
                    $driver.DriverDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffffffZ") 
                }
                else { 
                    $null 
                }
                DriverClass        = $driver.DeviceClass
                DriverID           = $driver.DeviceID
                DriverHardwareID   = $driver.HardwareID
                DriverManufacturer = $driver.Manufacturer
                DriverInfName      = $driver.InfName
                DriverLocation     = $driver.Location
                DriverDescription  = $driver.Description
                DriverProvider     = $driver.DriverProviderName
                DriverPublishedOn  = $null
                DriverStatus       = "Installed"
            }
        }
    )

    # Add optional updates to the list
    foreach ($optionalDriver in $OptionalWUList) {
        $LinkedDrivers += [PSCustomObject]@{
            WUName             = $optionalDriver.WUName
            DriverName         = $optionalDriver.DriverName
            DriverVersion      = $optionalDriver.DriverVersion
            DriverReleaseDate  = $optionalDriver.DriverReleaseDate
            DriverClass        = $optionalDriver.DriverClass
            DriverID           = $optionalDriver.DriverID
            DriverHardwareID   = $optionalDriver.DriverHardwareID
            DriverManufacturer = $optionalDriver.DriverManufacturer
            DriverInfName      = $optionalDriver.DriverInfName
            DriverLocation     = $optionalDriver.DriverLocation
            DriverDescription  = $optionalDriver.DriverDescription
            DriverProvider     = $optionalDriver.DriverProvider
            DriverPublishedOn  = $optionalDriver.DriverPublishedOn
            DriverStatus       = $optionalDriver.DriverStatus
        }
    }

    Return $LinkedDrivers

}

# Function to get Dell Warranty
function Get-DellWarranty(
    <#
.SYNOPSIS
    Retrieves Dell warranty information using the Dell Warranty API.
.DESCRIPTION
    Authenticates with Dell’s API using provided credentials, retrieves warranty details for the specified device, and returns warranty information as a custom object.
.PARAMETER SourceDevice
    The Dell service tag (serial number) of the device.
.OUTPUTS
    PSCustomObject with Dell warranty details.
#>    
    [Parameter(Mandatory = $true)]$SourceDevice) {
    $AuthURI = "https://apigtwb2c.us.dell.com/auth/oauth/v2/token"

    try {
        Write-Output "[$SourceDevice] Checking Dell OAuth token validity..."
        if ($Global:TokenAge -lt (Get-Date).AddMinutes(-55)) {
            Write-Output "[$SourceDevice] Token expired or missing. Clearing existing token..."
            $global:Token = $null
        }

        if ($null -eq $global:Token) {
            Write-Output "[$SourceDevice] Requesting new Dell OAuth token..."
            $OAuth = "$WarrantyDellClientID`:$WarrantyDellClientSecret"
            $Bytes = [System.Text.Encoding]::ASCII.GetBytes($OAuth)
            $EncodedOAuth = [Convert]::ToBase64String($Bytes)
            $headersAuth = @{ "authorization" = "Basic $EncodedOAuth" }
            $Authbody = 'grant_type=client_credentials'

            $AuthResult = Invoke-RESTMethod -Method Post -Uri $AuthURI -Body $AuthBody -Headers $HeadersAuth
            $global:token = $AuthResult.access_token
            $Global:TokenAge = Get-Date
            Write-Output "[$SourceDevice] Dell token acquired successfully."
        }

        Write-Output "[$SourceDevice] Submitting Dell warranty request..."
        $headersReq = @{ "Authorization" = "Bearer $global:Token" }
        $ReqBody = @{ servicetags = $SourceDevice }

        $WarReq = Invoke-RestMethod -Uri "https://apigtwb2c.us.dell.com/PROD/sbil/eapi/v5/asset-entitlements" -Headers $headersReq -Body $ReqBody -Method Get -ContentType "application/json"

        if ($WarReq.entitlements.serviceleveldescription) {
            Write-Output "[$SourceDevice] Warranty data received from Dell."

            $WarObj = [PSCustomObject]@{
                'ServiceProvider'         = 'Dell'
                'ServiceModel'            = $WarReq.productLineDescription
                'ServiceTag'              = $SourceDevice
                'ServiceLevelDescription' = $WarReq.entitlements.serviceleveldescription -join "`n"
                'WarrantyStartDate'       = ($WarReq.entitlements.startdate | Sort-Object -Descending | Select-Object -Last 1)
                'WarrantyEndDate'         = ($WarReq.entitlements.enddate | Sort-Object | Select-Object -Last 1)
            }
        }
        else {
            Write-Output "[$SourceDevice] No service level description returned by Dell."
            $WarObj = [PSCustomObject]@{
                'ServiceProvider'         = 'Dell'
                'ServiceModel'            = $null
                'ServiceTag'              = $SourceDevice
                'ServiceLevelDescription' = 'Could not get warranty information'
                'WarrantyStartDate'       = $null
                'WarrantyEndDate'         = $null
            }
        }
    }
    catch {
        Write-Output "[$SourceDevice] ERROR during Dell warranty lookup: $($_.Exception.Message)"
        $WarObj = [PSCustomObject]@{
            'ServiceProvider'         = 'Dell'
            'ServiceModel'            = $null
            'ServiceTag'              = $SourceDevice
            'ServiceLevelDescription' = 'Could not get warranty information'
            'WarrantyStartDate'       = $null
            'WarrantyEndDate'         = $null
        }
    }

    return $WarObj
}



# Function to get Lenovo Warranty
function Get-LenovoWarranty(
    <#
.SYNOPSIS
    Retrieves Lenovo warranty information using the Lenovo Warranty API.
.DESCRIPTION
    Queries Lenovo’s API using the provided client ID and device serial number, returning warranty details as a custom object.
.PARAMETER SourceDevice
    The Lenovo serial number of the device.
.OUTPUTS
    PSCustomObject with Lenovo warranty details.
#>    
    [Parameter(Mandatory = $true)]$SourceDevice) {
    $headersReq = @{ "ClientID" = $WarrantyLenovoClientID }
    $WarReq = Invoke-RestMethod -Uri "http://supportapi.lenovo.com/V2.5/Warranty?Serial=$SourceDevice" -Headers $headersReq -Method Get -ContentType "application/json"
    
    try {
        $Warlist = $WarReq.Warranty | Where-Object { ($_.ID -eq "36Y") -or ($_.ID -eq "3EZ") -or ($_.ID -eq "12B") -or ($_.ID -eq "1EZ") }
        $WarObj = [PSCustomObject]@{
            'ServiceProvider'         = 'Lenovo'
            'ServiceModel'            = $WarReq.Product
            'ServiceTag'              = $SourceDevice
            'ServiceLevelDescription' = $Warlist.Name -join "`n"
            'WarrantyStartDate'       = ($Warlist.Start | sort-object -Descending | select-object -last 1)
            'WarrantyEndDate'         = ($Warlist.End | sort-object | select-object -last 1)
        }

    }
    catch {
        $WarObj = [PSCustomObject]@{
            'ServiceProvider'         = 'Lenovo'
            'ServiceModel'            = $null
            'ServiceTag'              = $SourceDevice
            'ServiceLevelDescription' = 'Could not get warranty information'
            'WarrantyStartDate'       = $null
            'WarrantyEndDate'         = $null
        }
    }
    return $WarObj  
}

# Function to get Getac Warranty
function Get-GetacWarranty(
    <#
.SYNOPSIS
    Retrieves Getac warranty information using the Getac API.
.DESCRIPTION
    Queries Getac’s API with the device serial number and returns warranty details as a custom object.
.PARAMETER SourceDevice
    The Getac serial number of the device.
.OUTPUTS
    PSCustomObject with Getac warranty details.
#>
    [Parameter(Mandatory = $true)]$SourceDevice) {
    $WarReq = Invoke-RestMethod -Uri https://api.getac.us/rma-manager/rma/verify-serial?serial=$SerialNumber -Method Get -ContentType "application/json"
    try {
        $WarObj = [PSCustomObject]@{
            'ServiceProvider'         = 'Getac'
            'ServiceModel'            = $WarReq.model
            'ServiceTag'              = $SourceDevice
            'ServiceLevelDescription' = $WarReq.warrantyType
            'WarrantyStartDate'       = $null
            'WarrantyEndDate'         = ($warreq.endDeviceWarranty | sort-object | select-object -last 1)
        }
    }
    catch {
        $WarObj = [PSCustomObject]@{
            'ServiceProvider'         = 'Getac'
            'ServiceModel'            = $null
            'ServiceTag'              = $SourceDevice
            'ServiceLevelDescription' = 'Could not get warranty information'
            'WarrantyStartDate'       = $null
            'WarrantyEndDate'         = $null
        }
    }
    return $WarObj
}

# Function to get HP Warranty
function Get-HPWarranty(
    <#
.SYNOPSIS
    Retrieves HP warranty information using the HP Warranty API.
.DESCRIPTION
    Authenticates with HP’s API using provided credentials, submits a batch job for the device serial number, and returns warranty details as a custom object.
.PARAMETER SourceDevice
    The HP serial number of the device.
.OUTPUTS
    PSCustomObject with HP warranty details.
#>
    [Parameter(Mandatory = $true)]$SourceDevice) {

    try {
        Write-Output "[$SourceDevice] Requesting HP warranty token..."
        $b64EncodedCred = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("$WarrantyHPClientID`:$WarrantyHPClientSecret"))
        $tokenURI = "https://warranty.api.hp.com/oauth/v1/token"
        $tokenHeaders = @{
            accept        = "application/json"
            authorization = "Basic $b64EncodedCred"
        }
        $tokenBody = "grant_type=client_credentials"

        $authResponse = Invoke-WebRequest -UseBasicParsing -Method POST -Uri $tokenURI -Headers $tokenHeaders -Body $tokenBody -ContentType "application/x-www-form-urlencoded"
        $accessToken = ($authResponse | ConvertFrom-Json).access_token
        Write-Output "[$SourceDevice] Successfully obtained access token."

        Write-Output "[$SourceDevice] Submitting HP batch job..."
        $queryBody = "[{""sn"":""$SourceDevice""}]"
        $queryURI = "https://warranty.api.hp.com/productwarranty/v2/jobs"
        $queryHeaders = @{
            accept        = "application/json"
            authorization = "Bearer $accessToken"
        }

        $jobResponse = Invoke-WebRequest -UseBasicParsing -Method POST -Uri $queryURI -Headers $queryHeaders -Body $queryBody -ContentType "application/json"
        $jobData = $jobResponse | ConvertFrom-Json
        $jobId = $jobData.jobId
        $estimatedTime = $jobData.estimatedTime
        Write-Output "[$SourceDevice] Batch job created. Job ID: $jobId. Estimated time: $estimatedTime seconds."

        Start-Sleep -Seconds $estimatedTime

        $JobStatusURI = "$queryURI/$jobId"
        $JobResultsURI = "$queryURI/$jobId/results"

        Write-Output "[$SourceDevice] Polling job status..."

        do {
            $JobStatus = Invoke-WebRequest -UseBasicParsing -Method GET -Uri $JobStatusURI -Headers $queryHeaders | ConvertFrom-Json
            Write-Output "[$SourceDevice] Job status: $($JobStatus.status)"
            if ($JobStatus.status -ne "completed") {
                Start-Sleep -Seconds 15
            }
        } while ($JobStatus.status -ne "completed")

        Write-Output "[$SourceDevice] Job completed. Retrieving results..."

        $result = Invoke-WebRequest -UseBasicParsing -Method GET -Uri $JobResultsURI -Headers $queryHeaders | ConvertFrom-Json

        if (-not $result -or $result.Count -eq 0) {
            Write-Output "[$SourceDevice] No data returned in results array."
            throw "Empty results"
        }

        $deviceData = $result[0]  # root array
        $product = $deviceData.product
        $offers = $deviceData.offers

        if ($offers) {
            Write-Output "[$SourceDevice] Warranty data retrieved successfully."

            $serviceDescriptions = $offers | ForEach-Object {
                "$($_.offerDescription) ($($_.serviceObligationLineItemStartDate) to $($_.serviceObligationLineItemEndDate))"
            }

            $startDate = ($offers.serviceObligationLineItemStartDate | Sort-Object -Descending | Select-Object -First 1)
            $endDate = ($offers.serviceObligationLineItemEndDate | Sort-Object | Select-Object -Last 1)

            $WarObj = [PSCustomObject]@{
                'ServiceProvider'         = 'HP'
                'ServiceModel'            = $product.productDescription
                'ServiceTag'              = $SourceDevice
                'ServiceLevelDescription' = $serviceDescriptions -join "`n"
                'WarrantyStartDate'       = $startDate
                'WarrantyEndDate'         = $endDate
            }
        }
        else {
            Write-Output "[$SourceDevice] No offers found in response."
            $WarObj = [PSCustomObject]@{
                'ServiceProvider'         = 'HP'
                'ServiceModel'            = $product.productDescription
                'ServiceTag'              = $SourceDevice
                'ServiceLevelDescription' = 'No warranty offers returned'
                'WarrantyStartDate'       = $null
                'WarrantyEndDate'         = $null
            }
        }
    }
    catch {
        Write-Output "[$SourceDevice] ERROR during HP warranty lookup: $($_.Exception.Message)"
        $WarObj = [PSCustomObject]@{
            'ServiceProvider'         = 'HP'
            'ServiceModel'            = $null
            'ServiceTag'              = $SourceDevice
            'ServiceLevelDescription' = 'Could not get warranty information'
            'WarrantyStartDate'       = $null
            'WarrantyEndDate'         = $null
        }
    }

    return $WarObj
}


# Unified Warranty Retriever
function Get-Warranty {
    <#
.SYNOPSIS
    Retrieves and caches warranty information for a device from the appropriate vendor API.
.DESCRIPTION
    Determines manufacturer, checks for cached warranty data, and if necessary, queries the correct API for warranty details. Supports Dell, Lenovo, HP, and Getac.
.PARAMETER SerialNumber
    The device serial number.
.PARAMETER Manufacturer
    The device manufacturer.
.OUTPUTS
    PSCustomObject with warranty details.
#>
    param(
        [Parameter(Mandatory = $true)][string]$SerialNumber,
        [Parameter(Mandatory = $true)][string]$Manufacturer
    )
    $CachePath = "C:\Windows\Warranty_$SerialNumber.json"
    try {
        if (-not $WarrantyForceRefresh -and (Test-Path $CachePath)) {
            $fileAge = (Get-Date) - (Get-Item $CachePath).LastWriteTime
            if ($fileAge.TotalDays -le $WarrantyMaxCacheAgeDays) {
                Write-Output "[$SerialNumber] Using cached warranty data from $CachePath"
                Write-Output "[$SerialNumber] Skipping API call, using saved JSON."
                $cached = Get-Content $CachePath -Raw | ConvertFrom-Json
                $cached.WarrantyStartDate = [datetime]$cached.WarrantyStartDate
                $cached.WarrantyEndDate = [datetime]$cached.WarrantyEndDate
                return $cached
            }
            else {
                Write-Output "[$SerialNumber] Cache expired ($([math]::Round($fileAge.TotalDays,1)) days old). Refreshing..."
            }
        }
    }
    catch {
        Write-Output "[$SerialNumber] Exception during cache read: $($_.Exception.Message)"
    }
    $normalizedMake = ($Manufacturer -replace '\s+', ' ').Trim().ToUpper()
    Write-Output "Entering warranty switch block with: [$normalizedMake]"
    $WarrantyData = $null
    switch -Regex ($normalizedMake) {
        "^DELL" { $WarrantyData = Get-DellWarranty -SourceDevice $SerialNumber; break }
        "^LENOVO|^IBM" { $WarrantyData = Get-LenovoWarranty -SourceDevice $SerialNumber; break }
        "^INSYDE" { $WarrantyData = Get-GetacWarranty -SourceDevice $SerialNumber; break }
        "^HP" { $WarrantyData = Get-HPWarranty -SourceDevice $SerialNumber; break }
        default { Write-Output "[$SerialNumber] Manufacturer '$Manufacturer' not supported for warranty lookup." }
    }
    if ($WarrantyData) {
        $WarrantyData | ConvertTo-Json -Depth 5 | Out-File -FilePath $CachePath -Encoding UTF8
        Write-Output "[$SerialNumber] Warranty data cached at $CachePath"
        return $WarrantyData
    }
    return $null
}


# Function to create the authorization signature
Function New-Signature (
    # Function to create the authorization signature
    <#
.SYNOPSIS
    Creates an authorization signature for Azure Log Analytics API.
.DESCRIPTION
    Generates a SharedKey signature for authenticating requests to Azure Log Analytics.
.PARAMETER customerId
    The Log Analytics Workspace ID.
.PARAMETER sharedKey
    The Log Analytics Primary Key.
.PARAMETER date
    The RFC1123 date string.
.PARAMETER contentLength
    The length of the request body.
.PARAMETER method
    The HTTP method (e.g., POST).
.PARAMETER contentType
    The content type (e.g., application/json).
.PARAMETER resource
    The API resource path.
.OUTPUTS
    String containing the authorization header value.
#>    
    $customerId, $sharedKey, $date, $contentLength, $method, $contentType, $resource) {
    $xHeaders = "x-ms-date:" + $date
    $stringToHash = $method + "`n" + $contentLength + "`n" + $contentType + "`n" + $xHeaders + "`n" + $resource

    $bytesToHash = [Text.Encoding]::UTF8.GetBytes($stringToHash)
    $keyBytes = [Convert]::FromBase64String($sharedKey)

    $sha256 = New-Object System.Security.Cryptography.HMACSHA256
    $sha256.Key = $keyBytes
    $calculatedHash = $sha256.ComputeHash($bytesToHash)
    $encodedHash = [Convert]::ToBase64String($calculatedHash)
    $authorization = 'SharedKey {0}:{1}' -f $customerId, $encodedHash
    return $authorization
}

# Function to create and post the request
Function Send-LogAnalyticsData(
    <#
    .SYNOPSIS
    Sends data to Azure Log Analytics.
    .DESCRIPTION
    Compresses and uploads JSON data to Azure Log Analytics using the provided credentials and log type.
    .PARAMETER customerId
    The Log Analytics Workspace ID.
    .PARAMETER sharedKey
    The Log Analytics Primary Key.
    .PARAMETER body
    The request body (JSON, as bytes).
    .PARAMETER logType
    The custom log type name.
    .OUTPUTS
    String with the HTTP status code and payload size.
    #>    
    $customerId, $sharedKey, $body, $logType) {
    $method = "POST"
    $contentType = "application/json"
    $resource = "/api/logs"
    $rfc1123date = [DateTime]::UtcNow.ToString("r")
    $contentLength = $body.Length
    $signature = New-Signature `
        -customerId $customerId `
        -sharedKey $sharedKey `
        -date $rfc1123date `
        -contentLength $contentLength `
        -method $method `
        -contentType $contentType `
        -resource $resource
    $uri = "https://" + $customerId + ".ods.opinsights.azure.com" + $resource + "?api-version=2016-04-01"

    #validate that payload data does not exceed limits
    if ($body.Length -gt (31.9 * 1024 * 1024)) {
        throw("Upload payload is too big and exceed the 32Mb limit for a single upload. Please reduce the payload size. Current payload size is: " + ($body.Length / 1024 / 1024).ToString("#.#") + "Mb")
    }

    $payloadsize = ("Upload payload size is " + ($body.Length / 1024).ToString("#.#") + "Kb ")

    $headers = @{
        "Authorization"        = $signature;
        "Log-Type"             = $logType;
        "x-ms-date"            = $rfc1123date;
        "time-generated-field" = $TimeStampField;
    }

    $response = Invoke-WebRequest -Uri $uri -Method $method -ContentType $contentType -Headers $headers -Body $body -UseBasicParsing
    $statusmessage = "$($response.StatusCode) : $($payloadsize)"
    return $statusmessage
}
function Start-PowerShellSysNative {
    <#
.SYNOPSIS
    Launches a 64-bit PowerShell process from a 32-bit process.
.DESCRIPTION
    Ensures that scripts requiring 64-bit PowerShell can be executed from a 32-bit context, passing any specified arguments.
.PARAMETER Arguments
    Optional arguments to pass to the new PowerShell process.
#>
    param (
        [parameter(Mandatory = $false, HelpMessage = "Specify arguments that will be passed to the sysnative PowerShell process.")]
        [ValidateNotNull()]
        [string]$Arguments
    )

    # Get the sysnative path for powershell.exe
    $SysNativePowerShell = Join-Path -Path ($PSHOME.ToLower().Replace("syswow64", "sysnative")) -ChildPath "powershell.exe"

    # Construct new ProcessStartInfo object to run scriptblock in fresh process
    $ProcessStartInfo = New-Object -TypeName System.Diagnostics.ProcessStartInfo
    $ProcessStartInfo.FileName = $SysNativePowerShell
    $ProcessStartInfo.Arguments = $Arguments
    $ProcessStartInfo.RedirectStandardOutput = $true
    $ProcessStartInfo.RedirectStandardError = $true
    $ProcessStartInfo.UseShellExecute = $false
    $ProcessStartInfo.WindowStyle = "Hidden"
    $ProcessStartInfo.CreateNoWindow = $true

    # Instatiate the new 64-bit process
    $Process = [System.Diagnostics.Process]::Start($ProcessStartInfo)

    # Read standard error output to determine if the 64-bit script process somehow failed
    $ErrorOutput = $Process.StandardError.ReadToEnd()
    if ($ErrorOutput) {
        Write-Error -Message $ErrorOutput
    }
}#endfunction
#endregion functions

#region script
#Get Common data for App and Device Inventory:
#Get Intune DeviceID and ManagedDeviceName
if (@(Get-ChildItem HKLM:SOFTWARE\Microsoft\Enrollments\ -Recurse | Where-Object { $_.PSChildName -eq 'MS DM Server' })) {
    $MSDMServerInfo = Get-ChildItem HKLM:SOFTWARE\Microsoft\Enrollments\ -Recurse | Where-Object { $_.PSChildName -eq 'MS DM Server' }
    $ManagedDeviceInfo = Get-ItemProperty -LiteralPath "Registry::$($MSDMServerInfo)"
}
$ManagedDeviceName = $ManagedDeviceInfo.EntDeviceName
$ManagedDeviceID = $ManagedDeviceInfo.EntDMID
#Get Computer Info
$ComputerInfo = Get-ComputerInfo
$ComputerName = $ComputerInfo.CsName
$ComputerManufacturer = $ComputerInfo.CsManufacturer
if ($ComputerManufacturer.ToUpper() -eq "LENOVO" -or $ComputerManufacturer.ToUpper() -eq "IBM") {
    $ComputerModel = (Get-CimInstance -ClassName Win32_ComputerSystemProduct -ErrorAction SilentlyContinue).Version
}
else {
    $ComputerModel = $ComputerInfo.CsModel
}

#region DEVICEINVENTORY
if ($CollectDeviceInventory) {
    #Set Name of Log
    $DeviceLog = "PowerStacksDeviceInventory"

    # Get Computer Inventory Information
    $ComputerLastBootUpTime = $ComputerInfo.OsLastBootUpTime.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffffffZ")
    $ComputerPhysicalMemory = $ComputerInfo.CsTotalPhysicalMemory
    $ComputerNumberOfProcessors = (Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue).NumberOfProcessors
    $ComputerCPU = Get-CimInstance win32_processor -ErrorAction SilentlyContinue | Select-Object Manufacturer, Name, MaxClockSpeed, NumberOfCores, NumberOfLogicalProcessors
    $ComputerProcessorManufacturer = $ComputerCPU.Manufacturer | Get-Unique
    $ComputerProcessorName = $ComputerCPU.Name | Get-Unique
    $ComputerProcessorMaxClockSpeed = $ComputerCPU.MaxClockSpeed | Get-Unique
    $ComputerNumberOfCores = $ComputerCPU.NumberOfCores | Get-Unique
    $ComputerNumberOfLogicalProcessors = $ComputerCPU.NumberOfLogicalProcessors | Get-Unique
    $ComputerOSInstallDate = $ComputerInfo.OsInstallDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffffffZ")
 

    $BatteryDesignedCapacity = (Get-WmiObject -Class "BatteryStaticData" -Namespace "ROOT\WMI" -ErrorAction SilentlyContinue).DesignedCapacity
    $BatteryFullChargedCapacity = (Get-WmiObject -Class "BatteryFullChargedCapacity" -Namespace "ROOT\WMI" -ErrorAction SilentlyContinue).FullChargedCapacity

    #Grab Built-in Monitors PNPDeviceID
    if ($RemoveBuiltInMonitors) {
        $BuiltInMonitors = Get-CimInstance Win32_DesktopMonitor | Select-Object PNPDeviceID -ErrorAction SilentlyContinue
    }
    else {
        $BuiltInMonitors = $null
    }

    #Grabs the Monitor objects from WMI
    $Monitors = Get-WmiObject -Namespace "root\WMI" -Class "WMIMonitorID" -ErrorAction SilentlyContinue

    #Creates an empty array to hold the data
    $MonitorArray = @()

    #Takes each monitor object found and runs the following code:
    foreach ($Monitor in $Monitors) {

        if (-Not($Monitor.InstanceName.Substring(0, $Monitor.InstanceName.LastIndexOf('_')) -in $BuiltInMonitors.PNPDeviceID)) {

            # Initialize variables with null by default
            $MonitorModel = $null
            $MonitorSerialNumber = $null
            $MonitorManufacturer = $null

            # Safely decode UserFriendlyName
            if ($Monitor.UserFriendlyName -ne $null) {
                $MonitorModel = ([System.Text.Encoding]::ASCII.GetString($Monitor.UserFriendlyName)).Replace("$([char]0x0000)", "")
            }

            # Safely decode SerialNumberID
            if ($Monitor.SerialNumberID -ne $null) {
                $MonitorSerialNumber = ([System.Text.Encoding]::ASCII.GetString($Monitor.SerialNumberID)).Replace("$([char]0x0000)", "")
            }

            # Safely decode ManufacturerName
            if ($Monitor.ManufacturerName -ne $null) {
                $MonitorManufacturer = ([System.Text.Encoding]::ASCII.GetString($Monitor.ManufacturerName)).Replace("$([char]0x0000)", "")
            }

            $MonitorWeekOfManufacture = $Monitor.WeekOfManufacture
            $MonitorYearOfManufacture = $Monitor.YearOfManufacture

            $tempmonitor = New-Object -TypeName PSObject
            $tempmonitor | Add-Member -MemberType NoteProperty -Name "Manufacturer" -Value "$MonitorManufacturer" -Force
            $tempmonitor | Add-Member -MemberType NoteProperty -Name "Model" -Value "$MonitorModel" -Force
            $tempmonitor | Add-Member -MemberType NoteProperty -Name "SerialNumber" -Value "$MonitorSerialNumber" -Force
            $tempmonitor | Add-Member -MemberType NoteProperty -Name "WeekOfManufacture" -Value "$MonitorWeekOfManufacture" -Force
            $tempmonitor | Add-Member -MemberType NoteProperty -Name "YearOfManufacture" -Value "$MonitorYearOfManufacture" -Force
            $MonitorArray += $tempmonitor
        }
    }
    [System.Collections.ArrayList]$MonitorArrayList = $MonitorArray

    # Obtain physical disk details
    $Disks = Get-PhysicalDisk | Where-Object { $_.BusType -match "NVMe|SATA|SAS|ATAPI|RAID" } | Select-Object -Property DeviceId, BusType, FirmwareVersion, HealthStatus, Manufacturer, Model, FriendlyName, SerialNumber, Size, MediaType

    #Creates an empty array to hold the data
    $DiskArray = @()

    foreach ($Disk in ($Disks | Sort-Object DeviceID)) {

        # Obtain disk health information from current disk
        $DiskHealth = Get-PhysicalDisk | Where-Object { $_.DeviceId -eq $Disk.DeviceID } | Get-StorageReliabilityCounter | Select-Object -Property Wear, ReadErrorsTotal, ReadErrorsUncorrected, WriteErrorsTotal, WriteErrorsUncorrected, Temperature, TemperatureMax

        # Obtain SMART failure information
        $DrivePNPDeviceID = (Get-WmiObject -Class Win32_DiskDrive | Where-Object { $_.Index -eq $Disk.DeviceID }).PNPDeviceID
        $DriveSMARTStatus = (Get-WmiObject -namespace root\wmi -class MSStorageDriver_FailurePredictStatus -ErrorAction SilentlyContinue | Select-Object PredictFailure, Reason) | Where-Object { $_.InstanceName -eq $DrivePNPDeviceID }

        # Create custom PSObject
        $tempdisk = new-object -TypeName PSObject

        # Create disk health state entry
        $tempdisk | Add-Member -MemberType NoteProperty -Name "Number" -Value $Disk.DeviceID
        $tempdisk | Add-Member -MemberType NoteProperty -Name "BusType" -Value $Disk.BusType
        $tempdisk | Add-Member -MemberType NoteProperty -Name "FirmwareVersion" -Value $Disk.FirmwareVersion
        $tempdisk | Add-Member -MemberType NoteProperty -Name "HealthStatus" -Value $Disk.HealthStatus
        $tempdisk | Add-Member -MemberType NoteProperty -Name "Manufacturer" -Value $Disk.Manufacturer
        $tempdisk | Add-Member -MemberType NoteProperty -Name "Model" -Value $Disk.Model
        $tempdisk | Add-Member -MemberType NoteProperty -Name "Name" -Value $Disk.FriendlyName
        $tempdisk | Add-Member -MemberType NoteProperty -Name "SerialNumber" -Value $Disk.SerialNumber
        $tempdisk | Add-Member -MemberType NoteProperty -Name "Size" -Value $Disk.Size
        $tempdisk | Add-Member -MemberType NoteProperty -Name "Type" -Value $Disk.MediaType
        $tempdisk | Add-Member -MemberType NoteProperty -Name "SMARTPredictFailure" -Value $DriveSMARTStatus.PredictFailure
        $tempdisk | Add-Member -MemberType NoteProperty -Name "SMARTReason" -Value $DriveSMARTStatus.Reason
        $tempdisk | Add-Member -MemberType NoteProperty -Name "Wear" -Value $([int]($DiskHealth.Wear))
        $tempdisk | Add-Member -MemberType NoteProperty -Name "ReadErrorsUncorrected" -Value $DiskHealth.ReadErrorsUncorrected
        $tempdisk | Add-Member -MemberType NoteProperty -Name "ReadErrorsTotal" -Value $DiskHealth.ReadErrorsTotal
        $tempdisk | Add-Member -MemberType NoteProperty -Name "WriteErrorsUncorrected" -Value $DiskHealth.WriteErrorsUncorrected
        $tempdisk | Add-Member -MemberType NoteProperty -Name "WriteErrorsTotal" -Value $DiskHealth.WriteErrorsTotal
        $tempdisk | Add-Member -MemberType NoteProperty -Name "Temperature" -Value $([int]($DiskHealth.Temperature))
        $tempdisk | Add-Member -MemberType NoteProperty -Name "TemperatureMax" -Value $([int]($DiskHealth.TemperatureMax))

        $DiskArray += $tempdisk
    }
    [System.Collections.ArrayList]$DiskArrayList = $DiskArray


    # Query Win32_SystemEnclosure
    $SystemEnclosures = Get-WmiObject -Class Win32_SystemEnclosure

    # Create an empty array to hold the data
    $ChassisArray = @()

    # Process each enclosure instance
    foreach ($Enclosure in $SystemEnclosures) {
        $ChassisTypeCodes = $Enclosure.ChassisTypes | ForEach-Object { [int]$_ }
        $SMBIOSAssetTag = $Enclosure.SMBIOSAssetTag

        # Process each ChassisTypeCode for this enclosure
        foreach ($ChassisTypeCode in $ChassisTypeCodes) {
            # Create custom PSObject
            $tempChassis = New-Object -TypeName PSObject
            $tempChassis | Add-Member -MemberType NoteProperty -Name "ChassisTypeCode" -Value $ChassisTypeCode
            $tempChassis | Add-Member -MemberType NoteProperty -Name "ChassisTag" -Value $SMBIOSAssetTag

            # Add to array
            $ChassisArray += $tempChassis
        }
    }
    [System.Collections.ArrayList]$ChassisArrayList = $ChassisArray

    
    # CollectMicrosoft365
    if ($CollectMicrosoft365) {
        # Get Microsoft 365
        Write-Output 'Get Microsoft 365'
        $Microsoft365Data = Get-Microsoft365

        #Creates an empty object to hold the data
        $Microsoft365 = New-Object -TypeName PSObject

        if ($Microsoft365Data) {   
            $Microsoft365 | Add-Member -MemberType NoteProperty -Name "InstalledVersion" -Value $Microsoft365Data.InstalledVersion
            $Microsoft365 | Add-Member -MemberType NoteProperty -Name "UpdateChannel" -Value $Microsoft365Data.UpdateChannel
            $Microsoft365 | Add-Member -MemberType NoteProperty -Name "LatestReleaseType" -Value $Microsoft365Data.LatestReleaseType
            $Microsoft365 | Add-Member -MemberType NoteProperty -Name "LatestReleaseVersion" -Value $Microsoft365Data.LatestReleaseVersion
            $Microsoft365 | Add-Member -MemberType NoteProperty -Name "EndOfSupportDate" -Value $Microsoft365Data.EndOfSupportDate
            $Microsoft365 | Add-Member -MemberType NoteProperty -Name "ReleaseDate" -Value $Microsoft365Data.ReleaseDate
            $Microsoft365 | Add-Member -MemberType NoteProperty -Name "ReleaseID" -Value $Microsoft365Data.ReleaseID
        }
    }

    # CollectWarranty
    if ($CollectWarranty) {
        Write-Output "Collect Warranty"
        $WarrantyBios = Get-WmiObject Win32_Bios
        $WarrantyMake = $WarrantyBios.Manufacturer
        $WarrantySerialNumber = $WarrantyBios.SerialNumber
        Write-Output "Warranty Make  : $WarrantyMake"
        Write-Output "Warranty Serial: $WarrantySerialNumber"
        $WarrantyData = Get-Warranty -SerialNumber $WarrantySerialNumber -Manufacturer $WarrantyMake
        if ($WarrantyData) {
            $Warranty = [PSCustomObject]@{
                'ServiceProvider'         = $WarrantyData.ServiceProvider
                'ServiceModel'            = $WarrantyData.ServiceModel
                'ServiceTag'              = $WarrantyData.ServiceTag
                'ServiceLevelDescription' = $WarrantyData.ServiceLevelDescription
                'WarrantyStartDate'       = ([datetime]::Parse($WarrantyData.WarrantyStartDate)).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffffffZ")
                'WarrantyEndDate'         = ([datetime]::Parse($WarrantyData.WarrantyEndDate)).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffffffZ")
            }

            $Warranty | Format-List         

        }
    }

    # Create JSON to Upload to Log Analytics
    $Inventory = New-Object System.Object
    $Inventory | Add-Member -MemberType NoteProperty -Name "Memory" -Value "$ComputerPhysicalMemory" -Force
    $Inventory | Add-Member -MemberType NoteProperty -Name "CPUManufacturer" -Value "$ComputerProcessorManufacturer" -Force
    $Inventory | Add-Member -MemberType NoteProperty -Name "CPUName" -Value "$ComputerProcessorName" -Force
    $Inventory | Add-Member -MemberType NoteProperty -Name "CPUMaxClockSpeed" -Value "$ComputerProcessorMaxClockSpeed" -Force
    $Inventory | Add-Member -MemberType NoteProperty -Name "CPUPhysical" -Value "$ComputerNumberOfProcessors" -Force
    $Inventory | Add-Member -MemberType NoteProperty -Name "CPUCores" -Value "$ComputerNumberOfCores" -Force
    $Inventory | Add-Member -MemberType NoteProperty -Name "CPULogical" -Value "$ComputerNumberOfLogicalProcessors" -Force
    $Inventory | Add-Member -MemberType NoteProperty -Name "BatteryDesignedCapacity" -Value "$BatteryDesignedCapacity" -Force
    $Inventory | Add-Member -MemberType NoteProperty -Name "BatteryFullChargedCapacity" -Value "$BatteryFullChargedCapacity" -Force
    $Inventory | Add-Member -MemberType NoteProperty -Name "Monitors" -Value $MonitorArrayList -Force
    $Inventory | Add-Member -MemberType NoteProperty -Name "LastBootTime" -Value "$ComputerLastBootUpTime" -Force
    $Inventory | Add-Member -MemberType NoteProperty -Name "PhysicalDisks" -Value $DiskArrayList -Force
    $Inventory | Add-Member -MemberType NoteProperty -Name "DeviceManufacturer" -Value "$ComputerManufacturer" -Force
    $Inventory | Add-Member -MemberType NoteProperty -Name "DeviceModel" -Value "$ComputerModel" -Force
    $Inventory | Add-Member -MemberType NoteProperty -Name "Chassis" -Value $ChassisArrayList -Force
    $Inventory | Add-Member -MemberType NoteProperty -Name "OSInstallDate" -Value "$ComputerOSInstallDate" -Force
    if ($CollectMicrosoft365) {
        $Inventory | Add-Member -MemberType NoteProperty -Name "Microsoft365" -Value $Microsoft365 -Force
    }
    if ($CollectWarranty) {
        $Inventory | Add-Member -MemberType NoteProperty -Name "Warranty" -Value $Warranty -Force
    }

    $DeviceDetailsJson = $Inventory | ConvertTo-Json

    $ms = New-Object System.IO.MemoryStream
    $cs = New-Object System.IO.Compression.GZipStream($ms, [System.IO.Compression.CompressionMode]::Compress)
    $sw = New-Object System.IO.StreamWriter($cs)
    $sw.Write($DeviceDetailsJson)
    $sw.Close();
    $DeviceDetailsJson = [System.Convert]::ToBase64String($ms.ToArray())

    $MainDevice = New-Object -TypeName PSObject
    $MainDevice | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value "$ComputerName" -Force
    $MainDevice | Add-Member -MemberType NoteProperty -Name "ManagedDeviceID" -Value "$ManagedDeviceID" -Force
    if ($CollectMicrosoft365) {
        $MainDevice | Add-Member -MemberType NoteProperty -Name "Microsoft365" -Value $true -Force
    }
    if ($CollectWarranty -and $Warranty -and $Warranty.PSObject.Properties.Count -gt 0) {
        Write-Output "Warranty property count: $($Warranty.PSObject.Properties.Count)"
        Write-Output "Warranty data contents:`n$($WarrantyData | Out-String)"
        Write-Output "Warranty contents:`n$($Warranty | Out-String)"
        $MainDevice | Add-Member -MemberType NoteProperty -Name "Warranty" -Value $true -Force
    }
    else {
        if (-not $CollectWarranty) {
            Write-Output "Warranty collection not enabled. Skipping warranty flag."
        }
        elseif (-not $Warranty) {
            Write-Output "Warranty object is null."
        }
        elseif ($Warranty.PSObject.Properties.Count -eq 0) {
            Write-Output "Warranty object is present but has no properties."
        }
        else {
            Write-Output "Warranty check did not meet conditions. Unexpected state."
        }

        $MainDevice | Add-Member -MemberType NoteProperty -Name "Warranty" -Value $false -Force
    }

    
    
    $DeviceDetailsJsonArr = $DeviceDetailsJson -split '(.{31744})'

    $i = 0

    foreach ($DeviceDetails in $DeviceDetailsJsonArr) {

        if ($DeviceDetails.Length -gt 0 ) {
            $i++
            $MainDevice | Add-Member -MemberType NoteProperty -Name ("DeviceDetails" + $i.ToString()) -Value $DeviceDetails -Force
        }

    }
    if ($DeviceDetailsJson.Length -gt (10 * 31 * 1024)) {
        throw("DeviceDetails is too big and exceed the 32kb limit per column for a single upload. Please increase number of columns (#10). Current payload size is: " + ($DeviceDetailsJson.Length / 1024).ToString("#.#") + "kb")
    }

    $DeviceJson = $MainDevice | ConvertTo-Json

    # Submit the data to the API endpoint
    $ResponseDeviceInventory = Send-LogAnalyticsData -customerId $customerId -sharedKey $sharedKey -body ([System.Text.Encoding]::UTF8.GetBytes($DeviceJson)) -logType $DeviceLog
}
# end region DEVICEINVENTORY

# region APPINVENTORY
if ($CollectAppInventory) {
    Write-Output "Collect App Inventory"
    $AppLog = "PowerStacksAppInventory"

    $CurrentLoggedOnUser = (Get-WmiObject -Class win32_computersystem).UserName
    if ($CurrentLoggedOnUser -eq $null) {
        $CurrentOwner = Get-CimInstance Win32_Process -Filter 'name = "explorer.exe"' | Invoke-CimMethod -MethodName getowner
        $CurrentLoggedOnUser = "$($CurrentOwner.Domain)\$($CurrentOwner.User)"
    }
    $AdObj = New-Object System.Security.Principal.NTAccount($CurrentLoggedOnUser)
    $strSID = $AdObj.Translate([System.Security.Principal.SecurityIdentifier])
    $UserSid = $strSID.Value

    $MyApps = Get-InstalledApplications -UserSid $UserSid
    $MyApps | ForEach-Object { $_ | Add-Member -NotePropertyName AppType -NotePropertyValue 'Win32' -Force }

    if ($CollectUWPInventory) {
        $MyAppsAppx = Get-AppxInstalledApplications # Due to limitations of Get-AppxPackage on AADJ devices we don't use the SID
        $MyApps += $MyAppsAppx
    }

    $UniqueApps = ($MyApps | Group-Object Displayname | Where-Object { $_.Count -eq 1 } ).Group
    $DuplicatedApps = ($MyApps | Group-Object Displayname | Where-Object { $_.Count -gt 1 } ).Group
    $NewestDuplicateApp = ($DuplicatedApps | Group-Object DisplayName) | ForEach-Object { $_.Group | Sort-Object { [version]$_.DisplayVersion } -Descending | Select-Object -First 1 }
    $CleanAppList = $UniqueApps + $NewestDuplicateApp | Sort-Object DisplayName
    Write-Output 'Clean app list:'
    Write-Output $CleanAppList
    
    $AppArray = @()
    foreach ($App in $CleanAppList) {
        $tempapp = New-Object -TypeName PSObject

        if ($null -ne $App.DisplayName) {
            $tempapp | Add-Member -MemberType NoteProperty -Name "AppName" -Value $App.DisplayName -Force
        }
        else {
            $tempapp | Add-Member -MemberType NoteProperty -Name "AppName" -Value "" -Force
        }

        if ($null -ne $App.DisplayVersion) {
            $tempapp | Add-Member -MemberType NoteProperty -Name "AppVersion" -Value $App.DisplayVersion -Force
        }
        else {
            $tempapp | Add-Member -MemberType NoteProperty -Name "AppVersion" -Value "" -Force
        }

        if ($null -ne $App.Publisher) {
            $tempapp | Add-Member -MemberType NoteProperty -Name "AppPublisher" -Value $App.Publisher -Force
        }
        else {
            $tempapp | Add-Member -MemberType NoteProperty -Name "AppPublisher" -Value "" -Force
        }

        if ($null -ne $App.AppType) {
            $tempapp | Add-Member -MemberType NoteProperty -Name "AppType" -Value $App.AppType -Force
        }
        else {
            $tempapp | Add-Member -MemberType NoteProperty -Name "AppType" -Value "Unknown" -Force
        }

        if ($App.PSObject.Properties.Name -contains "InstallDate" -and $App.InstallDate) {
            $tempapp | Add-Member -MemberType NoteProperty -Name "AppInstallDate" -Value $App.InstallDate -Force
        }
        else {
            $tempapp | Add-Member -MemberType NoteProperty -Name "AppInstallDate" -Value $null -Force
        }

        if ($App.PSObject.Properties.Name -contains "UninstallString" -and $App.UninstallString) {
            $tempapp | Add-Member -MemberType NoteProperty -Name "AppUninstallString" -Value $App.UninstallString -Force
        }
        else {
            $tempapp | Add-Member -MemberType NoteProperty -Name "AppUninstallString" -Value $null -Force
        }

        if ($App.PSObject.Properties.Name -contains "PSPath" -and $App.PSPath) {
            $tempapp | Add-Member -MemberType NoteProperty -Name "AppUninstallRegPath" -Value $App.PSPath.Split("::")[-1] -Force
        }
        else {
            $tempapp | Add-Member -MemberType NoteProperty -Name "AppUninstallRegPath" -Value $null -Force
        }

        $AppArray += $tempapp
    }

    $InstalledAppJson = $AppArray | ConvertTo-Json

    $ms = New-Object System.IO.MemoryStream
    $cs = New-Object System.IO.Compression.GZipStream($ms, [System.IO.Compression.CompressionMode]::Compress)
    $sw = New-Object System.IO.StreamWriter($cs)
    $sw.Write($InstalledAppJson)
    $sw.Close()
    $InstalledAppJson = [System.Convert]::ToBase64String($ms.ToArray())

    $MainApp = New-Object -TypeName PSObject
    $MainApp | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value "$ComputerName" -Force
    $MainApp | Add-Member -MemberType NoteProperty -Name "ManagedDeviceID" -Value "$ManagedDeviceID" -Force

    $InstalledAppJsonArr = $InstalledAppJson -split '(.{31744})'
    $i = 0
    foreach ($InstalledApp in $InstalledAppJsonArr) {
        if ($InstalledApp.Length -gt 0 ) {
            $i++
            $MainApp | Add-Member -MemberType NoteProperty -Name ("InstalledApps" + $i.ToString()) -Value $InstalledApp -Force
        }
    }

    if ($InstalledAppJson.Length -gt (10 * 31 * 1024)) {
        throw("InstallApp is too big and exceed the 32kb limit per column for a single upload. Please increase number of columns (#10). Current payload size is: " + ($InstalledAppJson.Length / 1024).ToString("#.#") + "kb")
    }

    $AppJson = $MainApp | ConvertTo-Json   
    $ResponseAppInventory = Send-LogAnalyticsData -customerId $customerId -sharedKey $sharedKey -body ([System.Text.Encoding]::UTF8.GetBytes($AppJson)) -logType $AppLog
}
# end region APPINVENTORY



# region DRIVERINVENTORY
if ($CollectDriverInventory) {
    Write-Output "Collect Driver Inventory"
    #Set Name of Log
    $DriverLog = "PowerStacksDriverInventory"

    #get drivers
    $Drivers = Get-InstalledDrivers

    $DriverArray = @()
    foreach ($Driver in $Drivers) {
        $tempdriver = New-Object -TypeName PSObject
        $tempdriver | Add-Member -MemberType NoteProperty -Name "WUName" -Value $Driver.WUName -Force
        $tempdriver | Add-Member -MemberType NoteProperty -Name "DriverName" -Value $Driver.DriverName -Force
        $tempdriver | Add-Member -MemberType NoteProperty -Name "DriverVersion" -Value $Driver.DriverVersion -Force
        $tempdriver | Add-Member -MemberType NoteProperty -Name "DriverReleaseDate" -Value $Driver.DriverReleaseDate -Force
        $tempdriver | Add-Member -MemberType NoteProperty -Name "DriverClass" -Value $Driver.DriverClass -Force
        $tempdriver | Add-Member -MemberType NoteProperty -Name "DriverID" -Value $Driver.DriverID -Force
        $tempdriver | Add-Member -MemberType NoteProperty -Name "DriverHardwareID" -Value $Driver.DriverHardwareID -Force
        $tempdriver | Add-Member -MemberType NoteProperty -Name "DriverManufacturer" -Value $Driver.DriverManufacturer -Force
        $tempdriver | Add-Member -MemberType NoteProperty -Name "DriverInfName" -Value $Driver.DriverInfName -Force
        $tempdriver | Add-Member -MemberType NoteProperty -Name "DriverLocation" -Value $Driver.DriverLocation -Force
        $tempdriver | Add-Member -MemberType NoteProperty -Name "DriverDescription" -Value $Driver.DriverDescription -Force
        $tempdriver | Add-Member -MemberType NoteProperty -Name "DriverProvider" -Value $Driver.DriverProvider -Force
        $tempdriver | Add-Member -MemberType NoteProperty -Name "DriverPublishedOn" -Value $Driver.DriverPublishedOn -Force
        $tempdriver | Add-Member -MemberType NoteProperty -Name "DriverStatus" -Value $Driver.DriverStatus -Force
        $DriverArray += $tempdriver
    }

    $ListedDriverJson = $DriverArray | ConvertTo-Json

    $ms = New-Object System.IO.MemoryStream
    $cs = New-Object System.IO.Compression.GZipStream($ms, [System.IO.Compression.CompressionMode]::Compress)
    $sw = New-Object System.IO.StreamWriter($cs)
    $sw.Write($ListedDriverJson)
    $sw.Close();
    $ListedDriverJson = [System.Convert]::ToBase64String($ms.ToArray())

    $MainDriver = New-Object -TypeName PSObject
    $MainDriver | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value "$ComputerName" -Force
    $MainDriver | Add-Member -MemberType NoteProperty -Name "ManagedDeviceID" -Value "$ManagedDeviceID" -Force

    $ListedDriverJsonArr = $ListedDriverJson -split '(.{31744})'

    $i = 0

    foreach ($ListedDriver in $ListedDriverJsonArr) {

        if ($ListedDriver.Length -gt 0 ) {
            $i++
            $MainDriver | Add-Member -MemberType NoteProperty -Name ("ListedDrivers" + $i.ToString()) -Value $ListedDriver -Force
        }

    }
    if ($ListedDriverJson.Length -gt (10 * 31 * 1024)) {
        throw("Driver is too big and exceed the 32kb limit per column for a single upload. Please increase number of columns (#10). Current payload size is: " + ($ListedDriverJson.Length / 1024).ToString("#.#") + "kb")
    }

    $DriverJson = $MainDriver | ConvertTo-Json

    # Submit the data to the API endpoint
    $ResponseDriverInventory = Send-LogAnalyticsData -customerId $customerId -sharedKey $sharedKey -body ([System.Text.Encoding]::UTF8.GetBytes($DriverJson)) -logType $DriverLog
}
#endregion DRIVERINVENTORY

# Report back status
Write-Output "Report back status"
$date = (Get-Date).ToUniversalTime().ToString($InventoryDateFormat)
$OutputMessage = "InventoryDate: $date "

if ($CollectDeviceInventory) {
    if ($ResponseDeviceInventory -match "200 :") {

        $OutputMessage = $OutPutMessage + "DeviceInventory: OK " + $ResponseDeviceInventory
    }
    else {
        $OutputMessage = $OutPutMessage + "DeviceInventory: Fail "
    }
}
if ($CollectAppInventory) {
    if ($ResponseAppInventory -match "200 :") {

        $OutputMessage = $OutPutMessage + " AppInventory: OK " + $ResponseAppInventory
    }
    else {
        $OutputMessage = $OutPutMessage + " AppInventory: Fail "
    }
}
if ($CollectDriverInventory) {
    if ($ResponseDriverInventory -match "200 :") {

        $OutputMessage = $OutPutMessage + " DriverInventory: OK " + $ResponseDriverInventory
    }
    else {
        $OutputMessage = $OutPutMessage + " DriverInventory: Fail "
    }
}

Stop-Transcript

Write-Output $OutputMessage
#endregion script

