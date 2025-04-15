<#
.SYNOPSIS
This script collects application and device inventory data from Windows machines and sends the data to Azure Log Analytics.

.DESCRIPTION
The script gathers detailed information about installed applications and hardware inventory, including device information, memory, CPU, monitor details, and physical disk details. The data is then compressed, encoded in Base64, and sent to an Azure Log Analytics workspace for further analysis and reporting.

.PARAMETER CustomerId
The Log Analytics Workspace ID.

.PARAMETER SharedKey
The Primary Key for the Log Analytics Workspace.

.PARAMETER CollectDeviceInventory
Boolean parameter to specify whether to collect device inventory. Default is $true.

.PARAMETER CollectAppInventory
Boolean parameter to specify whether to collect application inventory. Default is $true.

.PARAMETER TimeStampField
Optional field to specify the timestamp from the data. If not specified, Azure Monitor assumes the ingestion time as the timestamp.

.PARAMETER RemoveBuiltInMonitors
Boolean parameter to specify whether to remove built-in monitors from the inventory. Default is $true.

.PARAMETER InventoryDateFormat
Format string for the inventory date. Default is "MM-dd HH:mm".

.EXAMPLE
.\InventoryCollector.ps1 -CustomerId "<YourWorkspaceID>" -SharedKey "<YourPrimaryKey>"

.NOTES
The script requires PowerShell 5.1 or later and the Azure Log Analytics workspace credentials.

Script Name: InventoryCollector.ps1
Date: 4/12/2025
Version: 5.0

# LEGAL DISCLAIMER
# This script is provided "as is" without any warranty of any kind, either express or implied, including but not limited to the implied warranties of merchantability, fitness for a particular purpose, or non-infringement. The entire risk as to the quality and performance of the script is with you.
# In no event shall the authors or copyright holders be liable for any claim, damages, or other liability, whether in an action of contract, tort, or otherwise, arising from, out of, or in connection with the script or the use or other dealings in the script.
# You should never run any script from the Internet without understanding its contents and effects. It is highly recommended that you thoroughly test the script in a safe environment before running it in production.
#>


#region initialize
# Enable TLS 1.2 support
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Replace with your Log Analytics Workspace ID
$CustomerId = ""  

# Replace with your Primary Key
$SharedKey = ""


#Control if you want to collect Device, App, and Driver Inventory or both (True = Collect)
$CollectDeviceInventory = $true
$CollectAppInventory = $true
$CollectDriverInventory = $true

#Sub-Control under Device Inventory
$CollectMicrosoft365 = $true
$CollectWarranty = $true

#Warranty key
#Warranty keys
$WarrantyDellClientID = ""
$WarrantyDellClientSecret = ""
$WarrantyLenovoClientID = $null

# You can use an optional field to specify the timestamp from the data. If the time field is not specified, Azure Monitor assumes the time is the message ingestion time
# DO NOT DELETE THIS VARIABLE. Recommened keep this blank.
$TimeStampField = ""

#Control if you want to remove BuiltIn Monitors (true = Remove)
$RemoveBuiltInMonitors = $false

#Inventory Date Format (sample: "MM-dd HH:mm", "dd-MM HH:mm")
$InventoryDateFormat = "MM-dd HH:mm"

#endregion initialize

#region functions

# Function to get all Installed Application
function Get-InstalledApplications() {
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

# Function to get Microsoft 365
function Get-Microsoft365() {
    ### Check for Click-to-Run Office
    $IsC2R = Test-Path 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun'
    if ($IsC2R) {
        try {
            $OfficeVersion = [version](Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction Stop | Select-Object -ExpandProperty VersionToReport)
            $OfficeProductIds = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction Stop | Select-Object -ExpandProperty ProductReleaseIds)
        }
        catch {
            Write-Output "Failed to retrieve Office version or product IDs: $_"
            $OfficeVersion = $null
            $OfficeProductIds = $null
        }
    }
    else {
        #Write-Output "No Click-to-Run Office detected. Setting default values."
        $OfficeVersion = $null
        $OfficeProductIds = $null
    }

    ### Determine if it’s Microsoft 365
    $IsM365 = ($OfficeProductIds -like '*O365*') -or ($OfficeProductIds -like '*M365*')

    ### Define update channels (corrected syntax)
    $Channels = @(
        @{ GUID = '492350f6-3a01-4f97-b9c0-c7c6ddf67d60'; PathPart = 'Monthly'; GPO = 'Current'; ID = 'Current'; Name = 'Monthly' }
        @{ GUID = '64256afe-f5d9-4f86-8936-8840a6a4f5be'; PathPart = 'MonthlyPreview'; GPO = 'FirstReleaseCurrent'; ID = 'CurrentPreview'; Name = 'Monthly (Preview)'; AlternateNames = @('InsiderSlow', 'FirstReleaseCurrent', 'Insiders') }
        @{ GUID = '55336b82-a18d-4dd6-b5f6-9e5095c314a6'; PathPart = 'MonthlyEnterpriseChannel'; GPO = 'MonthlyEnterprise'; ID = 'MonthlyEnterprise'; Name = 'MEC' }
        @{ GUID = '7ffbc6bf-bc32-4f92-8982-f9dd17fd3114'; PathPart = 'SAC'; GPO = 'Deferred'; ID = 'SemiAnnual'; Name = 'SAC'; AlternateNames = @('Deferred', 'Broad') }
        @{ GUID = 'b8f9b850-328d-4355-9145-c59439a0c4cf'; PathPart = 'SACT'; GPO = 'FirstReleaseDeferred'; ID = 'SemiAnnualPreview'; Name = 'SACT'; AlternateNames = @('FirstReleaseDeferred', 'Targeted') }
        @{ GUID = '5030841d-c919-4594-8d2d-84ae4f96e58e'; PathPart = 'LTSB2021'; ID = 'PerpetualVL2021'; Name = 'LTSB2021'; AlternateNames = @('Perpetual2021') }
        @{ GUID = 'f2e724c1-748f-4b47-8fb8-8e0d210e9208'; PathPart = 'LTSB'; ID = 'PerpetualVL2019'; Name = 'LTSB'; AlternateNames = @('Perpetual2019') }
        @{ GUID = '5440fd1f-7ecb-4221-8110-145efaa6372f'; PathPart = 'Beta'; GPO = 'InsiderFast'; ID = 'BetaChannel'; Name = 'Beta' }
    )

    ### Default channel if not determined
    $OfficeChannel = @{ Name = $null; PathPart = $null }

    ### Detect update channel for M365
    if ($IsM365 -and $IsC2R) {
        $OfficeUpdateChannelGPO = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\OfficeUpdate' -ErrorAction 'SilentlyContinue' | Select-Object -ExpandProperty UpdateBranch -ErrorAction 'SilentlyContinue')
        if ($OfficeUpdateChannelGPO) {
            Write-Output 'Office is configured to use a GPO update channel.'
            foreach ($Channel in $Channels) {
                if ($OfficeUpdateChannelGPO -eq $Channel.GPO) { $OfficeChannel = $Channel }
            }
        }
        else {
            $C2RConfigurationPath = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
            Write-Output 'Office is not configured to use a GPO update channel.'
            $OfficeUpdateURL = [System.Uri](Get-ItemProperty -Path $C2RConfigurationPath -ErrorAction 'SilentlyContinue' | Select-Object -ExpandProperty UpdateURL -ErrorAction 'SilentlyContinue')
            $OfficeUnmanagedUpdateURL = [System.Uri](Get-ItemProperty -Path $C2RConfigurationPath -ErrorAction 'SilentlyContinue' | Select-Object -ExpandProperty UnmanagedUpdateURL -ErrorAction 'SilentlyContinue')
            $OfficeUpdateChannelCDNURL = [System.Uri](Get-ItemProperty -Path $C2RConfigurationPath -ErrorAction 'SilentlyContinue' | Select-Object -ExpandProperty CDNBaseUrl -ErrorAction 'SilentlyContinue')
            if ($OfficeUpdateURL.IsAbsoluteUri) { $OfficeUpdateGUID = $OfficeUpdateURL.Segments[2] }
            elseif ($OfficeUnmanagedUpdateURL.IsAbsoluteUri) { $OfficeUpdateGUID = $OfficeUnmanagedUpdateURL.Segments[2] }
            elseif ($OfficeUpdateChannelCDNURL.IsAbsoluteUri) { $OfficeUpdateGUID = $OfficeUpdateChannelCDNURL.Segments[2] }
            else {
                Write-Output "Unable to determine Office update channel URL. Using default."
            }
            foreach ($Channel in $Channels) {
                if ($OfficeUpdateGUID -eq $Channel.GUID) { $OfficeChannel = $Channel }
            }
        }
        Write-Output ("{0} found using the {1} update channel. Channel ID: {2}. Detected Version: {3}" -f 'Microsoft 365 Apps', $OfficeChannel.Name, $OfficeChannel.ID, $OfficeVersion)
    }

    ### Get latest security update info
    if ($OfficeVersion.Major -eq 16 -and $IsC2R) {
        if ($IsM365) {
            $ChannelURLPathPart = $OfficeChannel.PathPart
            try {
                $UpdateAPIURL = "https://clients.config.office.net/releases/v1.0/LatestRelease/$ChannelURLPathPart`?ReleaseType=security"
                $ReleaseInfo = Invoke-RestMethod -Uri $UpdateAPIURL -Method 'GET' -ErrorAction 'Stop'
                if (-not $ReleaseInfo) {
                    $UpdateAPIURL = "https://clients.config.office.net/releases/v1.0/LatestRelease/$ChannelURLPathPart`?ReleaseType="
                    $ReleaseInfo = Invoke-RestMethod -Uri $UpdateAPIURL -Method 'GET' -ErrorAction 'Stop'
                }
            }
            catch {
                Write-Output "Unable to get latest update info: $_"
                $ReleaseInfo = [PSCustomObject]@{
                    releaseType      = $null
                    buildVersion     = $null
                    endOfSupportDate = $null
                    availabilityDate = $null
                    releaseVersion   = $null
                }
            }
        }
        elseif ($OfficeProductIds -like '*2019Volume*') {
            try {
                $UpdateAPIURL = 'https://clients.config.office.net/releases/v1.0/LatestRelease/LTSB?ReleaseType=security'
                $ReleaseInfo = Invoke-RestMethod -Uri $UpdateAPIURL -Method 'GET' -ErrorAction 'Stop'
                if (-not $ReleaseInfo) {
                    $UpdateAPIURL = 'https://clients.config.office.net/releases/v1.0/LatestRelease/LTSB?ReleaseType='
                    $ReleaseInfo = Invoke-RestMethod -Uri $UpdateAPIURL -Method 'GET' -ErrorAction 'Stop'
                }
            }
            catch {
                Write-Output "Unable to get latest update info: $_"
                $ReleaseInfo = [PSCustomObject]@{
                    releaseType      = $null
                    buildVersion     = $null
                    endOfSupportDate = $null
                    availabilityDate = $null
                    releaseVersion   = $null
                }
            }
        }
        elseif ($OfficeProductIds -like '*2021Volume*') {
            try {
                $UpdateAPIURL = 'https://clients.config.office.net/releases/v1.0/LatestRelease/LTSB2021?ReleaseType=security'
                $ReleaseInfo = Invoke-RestMethod -Uri $UpdateAPIURL -Method 'GET' -ErrorAction 'Stop'
                if (-not $ReleaseInfo) {
                    $UpdateAPIURL = 'https://clients.config.office.net/releases/v1.0/LatestRelease/LTSB2021?ReleaseType='
                    $ReleaseInfo = Invoke-RestMethod -Uri $UpdateAPIURL -Method 'GET' -ErrorAction 'Stop'
                }
            }
            catch {
                Write-Output "Unable to get latest update info: $_"
                $ReleaseInfo = [PSCustomObject]@{
                    releaseType      = $null
                    buildVersion     = $null
                    endOfSupportDate = $null
                    availabilityDate = $null
                    releaseVersion   = $null
                }
            }
        }
        else {
            Write-Output "Non-M365/Volume Office detected. Setting default release info."
            $ReleaseInfo = [PSCustomObject]@{
                releaseType      = $null
                buildVersion     = $null
                endOfSupportDate = $null
                availabilityDate = $null
                releaseVersion   = $null
            }
        }
    }
    else {
        $ReleaseInfo = [PSCustomObject]@{
            releaseType      = $null
            buildVersion     = $null
            endOfSupportDate = $null
            availabilityDate = $null
            releaseVersion   = $null
        }
    }

    ### Process release data
    $ReleaseTypes = @{
        1 = 'Feature Update'
        2 = 'Quality Update'
        3 = 'Security Update'
    }
    $Today = Get-Date
    $ReleaseType = if ($ReleaseTypes.ContainsKey([int32]$ReleaseInfo.releaseType)) { $ReleaseTypes[[int32]$ReleaseInfo.releaseType] } else { $null }
    $TargetVersion = [Version]$ReleaseInfo.buildVersion.buildVersionString

    ### Determine status
    if ($IsC2R) {  

        ### Preprocess data
        $InstalledVersion = if ($OfficeVersion) { 
            $OfficeVersion.ToString() 
        }
        else { 
            $null
        }

        $UpdateChannel = if ($OfficeChannel.Name) { 
            $OfficeChannel.Name
        }
        else { 
            $null
        }

        $LatestReleaseType = if ($ReleaseType) { 
            $ReleaseType.ToString() 
        }
        else { 
            $null
        }

        $LatestReleaseVersion = if ($TargetVersion) { 
            $TargetVersion.ToString() 
        }
        else { 
            $null
        }

        $EndOfSupportDate = if ($null -ne $ReleaseInfo.endOfSupportDate -and $ReleaseInfo.endOfSupportDate -ne '0001-01-01T00:00:00Z') { 
            $ReleaseInfo.endOfSupportDate.ToString() 
        }
        else { 
            $null 
        }

        $ReleaseDate = if ($ReleaseInfo -and $ReleaseInfo.availabilityDate -and $ReleaseInfo.availabilityDate -ne '0001-01-01T00:00:00Z') { 
            $ReleaseInfo.availabilityDate.ToString() 
        }
        else { 
            $null 
        }

        $ReleaseID = if ($ReleaseInfo -and $ReleaseInfo.releaseVersion) { 
            $ReleaseInfo.releaseVersion.ToString()
        }
        else { 
            $null 
        }

        ### Create data hashtable
        $OfficeVersionData = @{
            'InstalledVersion'     = $InstalledVersion
            'UpdateChannel'        = $UpdateChannel
            'LatestReleaseType'    = $LatestReleaseType
            'LatestReleaseVersion' = $LatestReleaseVersion
            'EndOfSupportDate'     = $EndOfSupportDate
            'ReleaseDate'          = $ReleaseDate
            'ReleaseID'            = $ReleaseID
        }
    }
    else {
        $OfficeVersionData = $null
    }

    Return $OfficeVersionData
}

# Function to get Installed Drivers
<#
Feel free to edit the query user to collect drivers.
#>
function Get-InstalledDrivers() {
    # Get PnP signed drivers
    $PNPSigned_Drivers = Get-CimInstance -ClassName Win32_PnPSignedDriver | Where-Object {
        ($_.Manufacturer -ne "Microsoft") -and 
        ($_.DriverProviderName -ne "Microsoft") -and 
        ($_.DeviceName -ne $null)
    } | Select-Object DeviceName, DriverVersion, DriverDate, DeviceClass, DeviceID, HardwareID, Manufacturer, InfName, Location, Description, DriverProviderName
    $PNPSigned_Drivers

    # Get installed MSU packages
    $InstalledDrivers = Get-Package -ProviderName msu | Where-Object {
        $_.Metadata.Item("SupportUrl") -match "target=hub"
    }
    $InstalledDrivers

    # Get optional updates
    $updateSession = New-Object -ComObject Microsoft.Update.Session
    $updateSearcher = $updateSession.CreateUpdateSearcher()
    $searchResult = $updateSearcher.Search("IsInstalled=0 AND Type='Driver'")
    $OptionalWUList = @()
    $searchResult.Updates.Count
    If ($searchResult.Updates.Count -gt 0) {
        For ($i = 0; $i -lt $searchResult.Updates.Count; $i++) {
            $update = $searchResult.Updates.Item($i)
            $OptionalWUList += [PSCustomObject]@{
                WUName             = $update.Title
                DriverName         = $update.DriverModel
                DriverVersion      = $null
                DriverReleaseDate  = $update.DriverVerDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffffffZ")
                DriverClass        = $update.DriverClass.ToUpper()
                DriverID           = $null
                DriverHardwareID   = $update.DriverHardwareID
                DriverManufacturer = $update.DriverManufacturer
                DriverInfName      = $null
                DriverLocation     = $null
                DriverDescription  = $update.Description
                DriverProvider     = $update.DriverProvider
                DriverPublishedOn  = $update.LastDeploymentChangeTime.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffffffZ")
                DriverStatus       = "Optional"
            }
        }
    }

    # Link installed drivers
    $LinkedDrivers = foreach ($installedDriver in $InstalledDrivers) {
        Write-Host "Attempting to link driver: $installedDriver"
        $versionFromName = $installedDriver.Name.Split()[-1]
        Write-Host "Driver version from name: $versionFromName"
        $matchingDriver = $PNPSigned_Drivers | Where-Object {
            $_.DriverVersion -eq $versionFromName
        } | Select-Object -First 1

        if ($matchingDriver) {
            [PSCustomObject]@{
                WUName             = $installedDriver.Name
                DriverName         = $matchingDriver.DeviceName
                DriverVersion      = $matchingDriver.DriverVersion
                DriverReleaseDate  = $matchingDriver.DriverDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffffffZ")
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
    $matchedVersions
    start-sleep 10
    $unmatchedDrivers = $PNPSigned_Drivers | Where-Object { $matchedVersions -notcontains $_.DriverVersion }
    $unmatchedDrivers

    # Combine both sets of drivers using the same foreach pattern
    $LinkedDrivers = @(
        $LinkedDrivers
        $LinkedDrivers  # Include existing linked drivers
        foreach ($driver in $unmatchedDrivers) {

            $driver.DeviceName
            $driver.DriverVersion
            # $driver.DriverDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffffffZ")
            $driver.DeviceClass 
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
            DriverReleaseDate  = $optionalDriver.DriverDate
            DriverClass        = $optionalDriver.DeviceClass
            DriverID           = $optionalDriver.DeviceID
            DriverHardwareID   = $optionalDriver.DriverHardwareID
            DriverManufacturer = $optionalDriver.Manufacturer
            DriverInfName      = $optionalDriver.InfName
            DriverLocation     = $optionalDriver.Location
            DriverDescription  = $optionalDriver.Description
            DriverProvider     = $optionalDriver.DriverProvider
            DriverPublishedOn  = $optionalDriver.DriverChangeTime
            DriverStatus       = $optionalDriver.DriverStatus
        }
    }

    Return $LinkedDrivers
}

# Function to get Dell Warranty
function Get-DellWarranty([Parameter(Mandatory = $true)]$SourceDevice) {
    $AuthURI = "https://apigtwb2c.us.dell.com/auth/oauth/v2/token"
    if ($Global:TokenAge -lt (get-date).AddMinutes(-55)) { $global:Token = $null }
    If ($null -eq $global:Token) {
        $OAuth = "$WarrantyDellClientID`:$WarrantyDellClientSecret"
        $Bytes = [System.Text.Encoding]::ASCII.GetBytes($OAuth)
        $EncodedOAuth = [Convert]::ToBase64String($Bytes)
        $headersAuth = @{ "authorization" = "Basic $EncodedOAuth" }
        $Authbody = 'grant_type=client_credentials'
        $AuthResult = Invoke-RESTMethod -Method Post -Uri $AuthURI -Body $AuthBody -Headers $HeadersAuth
        $global:token = $AuthResult.access_token
        $Global:TokenAge = (get-date)
    }
 
    $headersReq = @{ "Authorization" = "Bearer $global:Token" }
    $ReqBody = @{ servicetags = $SourceDevice }
    $WarReq = Invoke-RestMethod -Uri "https://apigtwb2c.us.dell.com/PROD/sbil/eapi/v5/asset-entitlements" -Headers $headersReq -Body $ReqBody -Method Get -ContentType "application/json"
    if ($warreq.entitlements.serviceleveldescription) {
        $WarObj = [PSCustomObject]@{
            'ServiceProvider'         = 'Dell'
            'ServiceModel'            = $warreq.productLineDescription
            'ServiceTag'              = $SourceDevice
            'ServiceLevelDescription' = $warreq.entitlements.serviceleveldescription -join "`n"
            'WarrantyStartDate'       = ($warreq.entitlements.startdate | sort-object -Descending | select-object -last 1)
            'WarrantyEndDate'         = ($warreq.entitlements.enddate | sort-object | select-object -last 1)
        }
    }
    else {
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
function Get-LenovoWarranty([Parameter(Mandatory = $true)]$SourceDevice) {
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
function Get-GetacWarranty([Parameter(Mandatory = $true)]$SourceDevice) {
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

# Function to create the authorization signature
Function New-Signature ($customerId, $sharedKey, $date, $contentLength, $method, $contentType, $resource) {
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
Function Send-LogAnalyticsData($customerId, $sharedKey, $body, $logType) {
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
        #Get Microsoft 365
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
        #Get Warranty Bios
        $WarrantyBios = Get-WmiObject Win32_Bios
        $WarrantyMake = $WarrantyBios.Manufacturer
        $WarrantySerialNumber = $WarrantyBios.SerialNumber

        if ($WarrantyDellClientID -ne $null -and $WarrantyDellClientSecret -ne $null -and $WarrantyMake -eq "Dell Inc.") {
            #write-host "Dell computer found" -ForegroundColor Green
            $WarrantyData = Get-DellWarranty -SourceDevice $WarrantySerialNumber
        } 
        elseif ($WarrantyLenovoClientID -ne $null -and $WarrantyMake -eq "LENOVO") {
            #write-host "LENOVO computer found" -ForegroundColor Green         
            $WarrantyData = Get-LenovoWarranty -SourceDevice $WarrantySerialNumber
        } 
        elseif ($GetacWarranty -and $WarrantyMake -eq "INSYDE Corp.") {
            #write-host "Getac computer found" -ForegroundColor Green
            $WarrantyData = Get-GetacWarranty -SourceDevice $WarrantySerialNumber
        }
        else {
            #write-host "$Make warranty not supported" -ForegroundColor Red
            $WarrantyData = $null
        }

        # Create custom PSObject
        $Warranty = New-Object -TypeName PSObject

        if ($WarrantyData) {
            $Warranty | Add-Member -MemberType NoteProperty -Name "ServiceProvider" -Value $WarrantyData.ServiceProvider
            $Warranty | Add-Member -MemberType NoteProperty -Name "ServiceModel" -Value $WarrantyData.ServiceModel
            $Warranty | Add-Member -MemberType NoteProperty -Name "ServiceTag" -Value $WarrantyData.ServiceTag
            $Warranty | Add-Member -MemberType NoteProperty -Name "ServiceLevelDescription" -Value $WarrantyData.ServiceLevelDescription
            $Warranty | Add-Member -MemberType NoteProperty -Name "WarrantyStartDate" -Value $WarrantyData.WarrantyStartDate
            $Warranty | Add-Member -MemberType NoteProperty -Name "WarrantyEndDate" -Value $WarrantyData.WarrantyEndDate
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
    if ($CollectWarranty) {
        $MainDevice | Add-Member -MemberType NoteProperty -Name "Warranty" -Value $true -Force
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
#endregion DEVICEINVENTORY

#region APPINVENTORY
if ($CollectAppInventory) {
    #Set Name of Log
    $AppLog = "PowerStacksAppInventory"

    #Get SID of current interactive users
    $CurrentLoggedOnUser = (Get-WmiObject -Class win32_computersystem).UserName
    if ($CurrentLoggedOnUser -eq $null) {
        $CurrentOwner = Get-CimInstance Win32_Process -Filter 'name = "explorer.exe"' | Invoke-CimMethod -MethodName getowner
        $CurrentLoggedOnUser = "$($CurrentOwner.Domain)\$($CurrentOwner.User)"
    }
    $AdObj = New-Object System.Security.Principal.NTAccount($CurrentLoggedOnUser)
    $strSID = $AdObj.Translate([System.Security.Principal.SecurityIdentifier])
    $UserSid = $strSID.Value

    #Get Apps for system and current user
    $MyApps = Get-InstalledApplications -UserSid $UserSid
    $UniqueApps = ($MyApps | Group-Object Displayname | Where-Object { $_.Count -eq 1 } ).Group
    $DuplicatedApps = ($MyApps | Group-Object Displayname | Where-Object { $_.Count -gt 1 } ).Group
    $NewestDuplicateApp = ($DuplicatedApps | Group-Object DisplayName) | ForEach-Object { $_.Group | Sort-Object [version]DisplayVersion -Descending | Select-Object -First 1 }
    $CleanAppList = $UniqueApps + $NewestDuplicateApp | Sort-Object DisplayName

    $AppArray = @()
    foreach ($App in $CleanAppList) {
        $tempapp = New-Object -TypeName PSObject
        $tempapp | Add-Member -MemberType NoteProperty -Name "AppName" -Value $App.DisplayName -Force
        $tempapp | Add-Member -MemberType NoteProperty -Name "AppVersion" -Value $App.DisplayVersion -Force
        $tempapp | Add-Member -MemberType NoteProperty -Name "AppInstallDate" -Value $App.InstallDate -Force -ErrorAction SilentlyContinue
        $tempapp | Add-Member -MemberType NoteProperty -Name "AppPublisher" -Value $App.Publisher -Force
        $tempapp | Add-Member -MemberType NoteProperty -Name "AppUninstallString" -Value $App.UninstallString -Force
        $tempapp | Add-Member -MemberType NoteProperty -Name "AppUninstallRegPath" -Value $app.PSPath.Split("::")[-1]
        $AppArray += $tempapp
    }

    $InstalledAppJson = $AppArray | ConvertTo-Json

    $ms = New-Object System.IO.MemoryStream
    $cs = New-Object System.IO.Compression.GZipStream($ms, [System.IO.Compression.CompressionMode]::Compress)
    $sw = New-Object System.IO.StreamWriter($cs)
    $sw.Write($InstalledAppJson)
    $sw.Close();
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

    # Submit the data to the API endpoint
    $ResponseAppInventory = Send-LogAnalyticsData -customerId $customerId -sharedKey $sharedKey -body ([System.Text.Encoding]::UTF8.GetBytes($AppJson)) -logType $AppLog
}
#endregion APPINVENTORY

#region DRIVERINVENTORY
if ($CollectDriverInventory) {
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

#Report back status
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
Write-Output $OutputMessage


#endregion script