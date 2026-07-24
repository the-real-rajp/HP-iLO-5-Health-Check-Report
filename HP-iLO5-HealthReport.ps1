<#+
.SYNOPSIS
Creates a Microsoft Word health report for an HPE iLO 5 server.

.DESCRIPTION
Uses Redfish session authentication, dynamically follows advertised resource
links, collects read-only health information, and creates a formatted .docx
report with its built-in Open XML generator. Microsoft Word COM automation is
available only as an explicit opt-in.
#>

[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$IloAddress,

    [PSCredential]$Credential,

    [string]$CustomerName,

    [string]$OutputPath,

    [ValidateRange(1, 3600)]
    [int]$TimeoutSec = 30,

    [ValidateRange(0, 10000)]
    [int]$MaxLogEntries = 100,

    [switch]$SkipCertificateCheck
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:WdPaperLetter = 2
$script:WdPageBreak = 7
$script:WdFieldNumPages = 26
$script:WdFieldPage = 33
$script:WdBorderTop = -1
$script:WdBorderBottom = -3
$script:WdPreferredWidthPoints = 3
$script:ReportBlue = '005F9E'
$script:ReportDark = '404040'
$script:ReportStripe = 'F2F7FB'
$script:ReportLogoPath = Join-Path $PSScriptRoot 'assets\winslowtg-logo.png'

function Get-ObjectProperty {
    param(
        [AllowNull()][object]$InputObject,
        [Parameter(Mandatory)][string]$Name,
        [AllowNull()][object]$Default = $null
    )

    if ($null -eq $InputObject) { return $Default }
    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -eq $property -or $null -eq $property.Value) { return $Default }
    return $property.Value
}

function Get-HeaderValue {
    param(
        [Parameter(Mandatory)][object]$Headers,
        [Parameter(Mandatory)][string]$Name
    )

    $values = @($Headers[$Name])
    if ($values.Count -eq 0 -or $null -eq $values[0]) { return $null }
    return [string]$values[0]
}

function Get-RedfishLink {
    param(
        [AllowNull()][object]$Resource,
        [Parameter(Mandatory)][string]$Name
    )

    $value = Get-ObjectProperty -InputObject $Resource -Name $Name
    if ($null -eq $value) { return $null }
    return Get-ObjectProperty -InputObject $value -Name '@odata.id'
}

function Get-RedfishLinkAny {
    param(
        [AllowNull()][object]$Resource,
        [Parameter(Mandatory)][string]$Name
    )

    $uri = Get-RedfishLink -Resource $Resource -Name $Name
    if ($uri) { return $uri }
    $links = Get-ObjectProperty -InputObject $Resource -Name 'Links'
    return Get-RedfishLink -Resource $links -Name $Name
}

function Get-HealthValue {
    param([AllowNull()][object]$Resource)

    $status = Get-ObjectProperty -InputObject $Resource -Name 'Status'
    $health = Get-ObjectProperty -InputObject $status -Name 'HealthRollup'
    if (-not $health) { $health = Get-ObjectProperty -InputObject $status -Name 'Health' }
    if (-not $health) { return 'Unknown' }
    return [string]$health
}

function Get-StateValue {
    param([AllowNull()][object]$Resource)

    $status = Get-ObjectProperty -InputObject $Resource -Name 'Status'
    return [string](Get-ObjectProperty -InputObject $status -Name 'State' -Default 'Unknown')
}

function Resolve-RedfishUri {
    param(
        [Parameter(Mandatory)][uri]$BaseUri,
        [Parameter(Mandatory)][string]$Uri
    )

    return [uri]::new($BaseUri, $Uri).AbsoluteUri
}

function New-IloSession {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][uri]$BaseUri,
        [Parameter(Mandatory)][PSCredential]$Credential,
        [Parameter(Mandatory)][int]$TimeoutSec,
        [bool]$IgnoreCertificateErrors = $false
    )

    $body = @{
        UserName = $Credential.UserName
        Password = $Credential.GetNetworkCredential().Password
    } | ConvertTo-Json

    $request = @{
        Uri = Resolve-RedfishUri -BaseUri $BaseUri -Uri '/redfish/v1/SessionService/Sessions'
        Method = 'Post'
        ContentType = 'application/json'
        Headers = @{ Accept = 'application/json'; 'OData-Version' = '4.0' }
        Body = $body
        TimeoutSec = $TimeoutSec
    }
    if ($PSVersionTable.PSVersion.Major -ge 7) {
        $request.SkipCertificateCheck = $IgnoreCertificateErrors
    }
    else {
        $request.UseBasicParsing = $true
    }
    $response = Invoke-WebRequest @request
    # PowerShell 7 returns header values as String[]. Passing that array back
    # through -Headers sends the literal text "System.String[]" instead of the
    # session token, which causes authenticated Redfish requests to return 401.
    $token = Get-HeaderValue -Headers $response.Headers -Name 'X-Auth-Token'
    if (-not $token) { throw 'iLO did not return a Redfish session token.' }
    $sessionUri = Get-HeaderValue -Headers $response.Headers -Name 'Location'

    [PSCustomObject]@{
        BaseUri = $BaseUri
        Headers = @{ Accept = 'application/json'; 'OData-Version' = '4.0'; 'X-Auth-Token' = $token }
        SessionUri = $sessionUri
        TimeoutSec = $TimeoutSec
        IgnoreCertificateErrors = $IgnoreCertificateErrors
    }
}

function Remove-IloSession {
    param([Parameter(Mandatory)][object]$Session)

    if (-not $Session.SessionUri) { return }
    try {
        $request = @{
            Uri = Resolve-RedfishUri -BaseUri $Session.BaseUri -Uri $Session.SessionUri
            Method = 'Delete'
            Headers = $Session.Headers
            TimeoutSec = $Session.TimeoutSec
        }
        if ($PSVersionTable.PSVersion.Major -ge 7) {
            $request.SkipCertificateCheck = [bool]$Session.IgnoreCertificateErrors
        }
        else {
            $request.UseBasicParsing = $true
        }
        Invoke-WebRequest @request | Out-Null
    }
    catch {
        if ($_.Exception.Message -match '(?i)underlying connection was closed|unexpected error occurred on a send') {
            Write-Verbose "iLO closed the connection while the Redfish session was being removed."
        }
        else {
            Write-Warning "Unable to close the Redfish session: $($_.Exception.Message)"
        }
    }
}

function Invoke-RedfishGet {
    param(
        [Parameter(Mandatory)][object]$Session,
        [Parameter(Mandatory)][string]$Uri
    )

    $request = @{
        Uri = Resolve-RedfishUri -BaseUri $Session.BaseUri -Uri $Uri
        Method = 'Get'
        Headers = $Session.Headers
        TimeoutSec = $Session.TimeoutSec
    }
    if ($PSVersionTable.PSVersion.Major -ge 7) {
        $request.SkipCertificateCheck = [bool]$Session.IgnoreCertificateErrors
    }

    for ($attempt = 1; $attempt -le 3; $attempt++) {
        try {
            return Invoke-RestMethod @request
        }
        catch {
            $transient = $_.Exception.Message -match '(?i)underlying connection was closed|unexpected error occurred on a send|connection reset|forcibly closed|temporarily unavailable|timed out'
            if (-not $transient -or $attempt -eq 3) { throw }
            Start-Sleep -Milliseconds (250 * $attempt)
        }
    }
}

function Get-RedfishCollection {
    param(
        [Parameter(Mandatory)][object]$Session,
        [AllowNull()][string]$Uri,
        [int]$Limit = [int]::MaxValue
    )

    $items = [System.Collections.Generic.List[object]]::new()
    if (-not $Uri -or $Limit -eq 0) { return $items.ToArray() }
    $nextUri = $Uri

    while ($nextUri -and $items.Count -lt $Limit) {
        $collection = Invoke-RedfishGet -Session $Session -Uri $nextUri
        $members = @(Get-ObjectProperty -InputObject $collection -Name 'Members' -Default @())
        foreach ($member in $members) {
            if ($items.Count -ge $Limit) { break }
            $memberUri = Get-ObjectProperty -InputObject $member -Name '@odata.id'
            if ($memberUri) {
                $items.Add((Invoke-RedfishGet -Session $Session -Uri $memberUri))
            }
            else {
                $items.Add($member)
            }
        }
        $nextUri = Get-ObjectProperty -InputObject $collection -Name 'Members@odata.nextLink'
    }
    return $items.ToArray()
}

function Add-CollectionNote {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][System.Collections.Generic.List[string]]$Notes,
        [Parameter(Mandatory)][string]$Message
    )
    if (-not $Notes.Contains($Message)) {
        $Notes.Add($Message)
    }
}

function Get-SafeCollectionFromUris {
    param(
        [Parameter(Mandatory)][object]$Session,
        [AllowEmptyCollection()][string[]]$Uris,
        [Parameter(Mandatory)][AllowEmptyCollection()][System.Collections.Generic.List[string]]$Notes,
        [string]$Label = 'resource',
        [int]$Limit = [int]::MaxValue
    )

    $candidates = @($Uris | Where-Object { $_ } | Select-Object -Unique)
    if ($candidates.Count -eq 0) {
        Add-CollectionNote -Notes $Notes -Message "$Label is not advertised by this system."
        return @()
    }

    $lastError = $null
    $successfulEmptyResponse = $false
    foreach ($uri in $candidates) {
        try {
            $result = @(Get-RedfishCollection -Session $Session -Uri $uri -Limit $Limit)
            if ($result.Count -gt 0) { return $result }
            $successfulEmptyResponse = $true
        }
        catch {
            $lastError = $_.Exception.Message
        }
    }

    if ($successfulEmptyResponse) { return @() }
    Add-CollectionNote -Notes $Notes -Message "Unable to collect $Label`: $lastError"
    return @()
}

function Get-SafeResourceFromUris {
    param(
        [Parameter(Mandatory)][object]$Session,
        [AllowEmptyCollection()][string[]]$Uris,
        [Parameter(Mandatory)][AllowEmptyCollection()][System.Collections.Generic.List[string]]$Notes,
        [string]$Label = 'resource'
    )

    $candidates = @($Uris | Where-Object { $_ } | Select-Object -Unique)
    if ($candidates.Count -eq 0) {
        Add-CollectionNote -Notes $Notes -Message "$Label is not advertised by this system."
        return $null
    }

    $lastError = $null
    foreach ($uri in $candidates) {
        try {
            return Invoke-RedfishGet -Session $Session -Uri $uri
        }
        catch {
            $lastError = $_.Exception.Message
        }
    }

    Add-CollectionNote -Notes $Notes -Message "Unable to collect $Label`: $lastError"
    return $null
}

function Get-SafeCollection {
    param(
        [Parameter(Mandatory)][object]$Session,
        [AllowNull()][string]$Uri,
        [Parameter(Mandatory)][AllowEmptyCollection()][System.Collections.Generic.List[string]]$Notes,
        [string]$Label = 'resource',
        [int]$Limit = [int]::MaxValue
    )

    if (-not $Uri) {
        Add-CollectionNote -Notes $Notes -Message "$Label is not advertised by this system."
        return @()
    }
    try {
        return @(Get-RedfishCollection -Session $Session -Uri $Uri -Limit $Limit)
    }
    catch {
        Add-CollectionNote -Notes $Notes -Message "Unable to collect $Label`: $($_.Exception.Message)"
        return @()
    }
}

function Convert-ServerStatus {
    param([Parameter(Mandatory)][object]$System)

    $oem = Get-ObjectProperty $System 'Oem'
    $hpe = Get-ObjectProperty $oem 'Hpe'
    $hostOs = Get-ObjectProperty $hpe 'HostOS'
    $osParts = @(@(
        Get-ObjectProperty $hostOs 'OsName'
        Get-ObjectProperty $hostOs 'OsVersion'
    ) | Where-Object { $_ })
    [ordered]@{
        'Product name' = Get-ObjectProperty $System 'Model' 'Unknown'
        'Server name' = (Get-ObjectProperty $System 'HostName' (Get-ObjectProperty $System 'Name' 'Unknown'))
        'Operating system' = if ($osParts.Count -gt 0) { $osParts -join ' ' } else { Get-ObjectProperty $hostOs 'OsDescription' 'Unknown' }
        'System ROM' = Get-ObjectProperty $System 'BiosVersion' 'Unknown'
        'Server serial number' = Get-ObjectProperty $System 'SerialNumber' 'Unknown'
        'Product ID' = Get-ObjectProperty $System 'SKU' 'Unknown'
        'UUID' = Get-ObjectProperty $System 'UUID' 'Unknown'
        'Manufacturer' = Get-ObjectProperty $System 'Manufacturer' 'Unknown'
        'Power state' = Get-ObjectProperty $System 'PowerState' 'Unknown'
        'Health' = Get-HealthValue $System
    }
}

function Convert-IloInformation {
    param([Parameter(Mandatory)][object]$Manager)

    [PSCustomObject][ordered]@{
        'Name' = Get-ObjectProperty $Manager 'Name' (Get-ObjectProperty $Manager 'Id' 'iLO Manager')
        'Model' = Get-ObjectProperty $Manager 'Model' 'Unknown'
        'Manager type' = Get-ObjectProperty $Manager 'ManagerType' 'Unknown'
        'Firmware version' = Get-ObjectProperty $Manager 'FirmwareVersion' 'Unknown'
        'Date and time' = Get-ObjectProperty $Manager 'DateTime' 'Unknown'
        'Health' = Get-HealthValue $Manager
    }
}

function Convert-ComputeOpsManagement {
    param([Parameter(Mandatory)][object]$Manager)

    $oem = Get-ObjectProperty $Manager 'Oem'
    $hpe = Get-ObjectProperty $oem 'Hpe'
    $cloudConnect = Get-ObjectProperty $hpe 'CloudConnect'
    [PSCustomObject][ordered]@{
        'Manager' = Get-ObjectProperty $Manager 'Name' (Get-ObjectProperty $Manager 'Id' 'iLO Manager')
        'Supported' = if ($null -eq $cloudConnect) { 'Not advertised' } else { 'Yes' }
        'Connection status' = Get-ObjectProperty $cloudConnect 'CloudConnectStatus' 'Unknown'
        'Workspace ID' = Get-ObjectProperty $cloudConnect 'WorkspaceId' 'Not reported'
    }
}

function Convert-SystemNetworkInterface {
    param(
        [Parameter(Mandatory)][object]$Item,
        [string]$Type = 'Ethernet interface'
    )

    $ipAddresses = @(
        @(Get-ObjectProperty $Item 'IPv4Addresses' @()) +
        @(Get-ObjectProperty $Item 'IPv6Addresses' @()) |
            ForEach-Object { Get-ObjectProperty $_ 'Address' } |
            Where-Object { $_ }
    ) -join '; '
    $ethernet = Get-ObjectProperty $Item 'Ethernet'
    $macAddress = Get-ObjectProperty $Item 'MACAddress' (Get-ObjectProperty $ethernet 'MACAddress' 'Unknown')
    $speed = Get-ObjectProperty $Item 'SpeedMbps' (Get-ObjectProperty $Item 'CurrentLinkSpeedMbps' 'N/A')

    [PSCustomObject][ordered]@{
        'Name' = Get-ObjectProperty $Item 'Name' (Get-ObjectProperty $Item 'Id' 'Unknown')
        'Type' = $Type
        'MAC address' = $macAddress
        'IP address' = if ($ipAddresses) { $ipAddresses } else { 'N/A' }
        'Link' = Get-ObjectProperty $Item 'LinkStatus' 'Unknown'
        'Speed (Mbps)' = $speed
        'Health' = Get-HealthValue $Item
    }
}

function Convert-DeviceInventory {
    param([Parameter(Mandatory)][object]$Item)

    $firmwareVersion = Get-ObjectProperty $Item 'FirmwareVersion'
    $currentFirmware = Get-ObjectProperty $firmwareVersion 'Current'
    $versionString = Get-ObjectProperty $currentFirmware 'VersionString'
    if (-not $versionString -and $firmwareVersion -is [string]) { $versionString = $firmwareVersion }
    $health = Get-HealthValue $Item
    $state = Get-StateValue $Item
    $displayStatus = if ($health -and $health -notmatch '(?i)^unknown$') { $health } else { $state }
    [PSCustomObject][ordered]@{
        'Location' = Get-ObjectProperty $Item 'Location' 'Unknown'
        'Product name' = Get-ObjectProperty $Item 'Name' (Get-ObjectProperty $Item 'Model' (Get-ObjectProperty $Item 'Id' 'Unknown'))
        'Product version' = Get-ObjectProperty $Item 'ProductVersion' 'Unknown'
        'Firmware version' = if ($versionString) { $versionString } else { 'Unknown' }
        'Status' = $displayStatus
    }
}

function Convert-RemoteSupportRegistration {
    param([Parameter(Mandatory)][object]$Item)

    $enabled = Get-ObjectProperty $Item 'RemoteSupportEnabled'
    $registration = if ($enabled -eq $true) {
        'Registered'
    }
    elseif ($enabled -eq $false) {
        'Not registered'
    }
    else {
        'Unknown'
    }
    [PSCustomObject][ordered]@{
        'Registration' = $registration
        'Connection model' = Get-ObjectProperty $Item 'ConnectModel' 'Unknown'
        'Destination' = Get-ObjectProperty $Item 'DestinationURL' 'Unknown'
        'Destination port' = Get-ObjectProperty $Item 'DestinationPort' 'Unknown'
        'External agent' = Get-ObjectProperty $Item 'ExternalAgentName' 'Unknown'
        'Last transmission' = Get-ObjectProperty $Item 'LastTransmissionDate' 'Unknown'
        'Last transmission type' = Get-ObjectProperty $Item 'LastTransmissionType' 'Unknown'
        'Last transmission error' = Get-ObjectProperty $Item 'LastTransmissionError' 'Unknown'
        'Maintenance mode' = Get-ObjectProperty $Item 'MaintenanceModeEnabled' 'Unknown'
    }
}

function Convert-Temperature {
    param([Parameter(Mandatory)][object]$Item)
    [PSCustomObject][ordered]@{
        'Name' = Get-ObjectProperty $Item 'Name' 'Unknown'
        'Reading (C)' = Get-ObjectProperty $Item 'ReadingCelsius' 'N/A'
        'Upper critical (C)' = Get-ObjectProperty $Item 'UpperThresholdCritical' 'N/A'
        'Health' = Get-HealthValue $Item
    }
}

function Convert-Fan {
    param([Parameter(Mandatory)][object]$Item)
    $units = Get-ObjectProperty $Item 'ReadingUnits' 'N/A'
    $reading = Get-ObjectProperty $Item 'Reading' 'N/A'
    if ($units -match '(?i)^percent$' -and $reading -notmatch '(?i)^n/?a$' -and [string]$reading -notmatch '%$') {
        $reading = "$reading%"
    }
    [PSCustomObject][ordered]@{
        'Name' = Get-ObjectProperty $Item 'Name' 'Unknown'
        'Reading' = $reading
        'Units' = $units
        'Health' = Get-HealthValue $Item
    }
}

function Convert-PowerSupply {
    param([Parameter(Mandatory)][object]$Item)
    [PSCustomObject][ordered]@{
        'Name' = Get-ObjectProperty $Item 'Name' 'Unknown'
        'Model' = Get-ObjectProperty $Item 'Model' 'Unknown'
        'Serial number' = Get-ObjectProperty $Item 'SerialNumber' 'Unknown'
        'Capacity (W)' = Get-ObjectProperty $Item 'PowerCapacityWatts' 'N/A'
        'Health' = Get-HealthValue $Item
    }
}

function Convert-Memory {
    param([Parameter(Mandatory)][object]$Item)
    [PSCustomObject][ordered]@{
        'Name' = Get-ObjectProperty $Item 'DeviceLocator' (Get-ObjectProperty $Item 'Name' 'Unknown')
        'Capacity (MiB)' = Get-ObjectProperty $Item 'CapacityMiB' 'N/A'
        'Type' = Get-ObjectProperty $Item 'MemoryDeviceType' 'Unknown'
        'Speed (MHz)' = Get-ObjectProperty $Item 'OperatingSpeedMhz' 'N/A'
        'Health' = Get-HealthValue $Item
    }
}

function Test-ReportRecordPresent {
    param([AllowNull()][object]$Record)

    if ($null -eq $Record) { return $false }
    $state = Get-ObjectProperty -InputObject $Record -Name 'State' -Default ''
    $status = Get-ObjectProperty -InputObject $Record -Name 'Status' -Default ''
    return ([string]$state -notmatch '(?i)^absent$' -and [string]$status -notmatch '(?i)^absent$')
}

function Test-UnknownReportValue {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) { return $true }
    return ([string]$Value).Trim() -match '(?i)^(unknown|absent)?$'
}

function Test-NotApplicableReportValue {
    param([AllowNull()][object]$Value)

    return $null -ne $Value -and ([string]$Value).Trim() -match '(?i)^n/?a$'
}

function Get-ReportRecords {
    param(
    [Parameter(Mandatory)][object]$Data,
    [Parameter(Mandatory)][string]$PropertyName,
        [switch]$PropertyMap,
        [string[]]$ExcludeColumns = @(),
        [switch]$OmitNAValues
    )

    $property = $Data.PSObject.Properties[$PropertyName]
    if ($null -eq $property -or $null -eq $property.Value) { return @() }
    if ($PropertyMap -and $property.Value -is [Collections.IDictionary]) {
        return @($property.Value.GetEnumerator() | ForEach-Object {
            if (-not (Test-UnknownReportValue $_.Value)) {
                [PSCustomObject][ordered]@{ Item = $_.Key; Value = $_.Value }
            }
        })
    }
    $records = @($property.Value | Where-Object { Test-ReportRecordPresent $_ })
    if ($records.Count -eq 0) { return @() }

    $columns = [Collections.Generic.List[string]]::new()
    foreach ($record in $records) {
        foreach ($recordProperty in $record.PSObject.Properties) {
            if (-not $columns.Contains($recordProperty.Name)) { $columns.Add($recordProperty.Name) }
        }
    }
    $usefulColumns = @($columns | Where-Object {
        $columnName = $_
        if ($columnName -in $ExcludeColumns) { return $false }
        @($records | Where-Object {
            $value = Get-ObjectProperty $_ $columnName
            -not (Test-UnknownReportValue $value) -and (-not $OmitNAValues -or -not (Test-NotApplicableReportValue $value))
        }).Count -gt 0
    })
    if ($usefulColumns.Count -eq 0) { return @() }

    return @($records | ForEach-Object {
        $source = $_
        $row = [ordered]@{}
        foreach ($columnName in $usefulColumns) {
            $value = Get-ObjectProperty $source $columnName
            $row[$columnName] = if ((Test-UnknownReportValue $value) -or ($OmitNAValues -and (Test-NotApplicableReportValue $value))) { '' } else { $value }
        }
        [PSCustomObject]$row
    })
}

function Get-ReportPropertyRows {
    param([AllowNull()][object]$Record)

    if ($null -eq $Record) { return @() }
    return @($Record.PSObject.Properties | ForEach-Object {
        if (-not (Test-UnknownReportValue $_.Value)) {
            [PSCustomObject][ordered]@{ Item = $_.Name; Value = $_.Value }
        }
    })
}

function Convert-Processor {
    param([Parameter(Mandatory)][object]$Item)
    [PSCustomObject][ordered]@{
        'Name' = Get-ObjectProperty $Item 'Socket' (Get-ObjectProperty $Item 'Name' 'Unknown')
        'Model' = Get-ObjectProperty $Item 'Model' 'Unknown'
        'Cores' = Get-ObjectProperty $Item 'TotalCores' 'N/A'
        'Threads' = Get-ObjectProperty $Item 'TotalThreads' 'N/A'
        'Speed (MHz)' = Get-ObjectProperty $Item 'OperatingSpeedMHz' (Get-ObjectProperty $Item 'MaxSpeedMHz' 'Unknown')
        'Instruction set' = Get-ObjectProperty $Item 'InstructionSet' 'Unknown'
        'Health' = Get-HealthValue $Item
    }
}

function Convert-Firmware {
    param([Parameter(Mandatory)][object]$Item)
    $health = Get-HealthValue $Item
    if ($health -match '(?i)^unknown$') { $health = 'OK' }
    [PSCustomObject][ordered]@{
        'Name' = Get-ObjectProperty $Item 'Name' (Get-ObjectProperty $Item 'Id' 'Unknown')
        'Version' = Get-ObjectProperty $Item 'Version' 'Unknown'
        'Health' = $health
    }
}

function Convert-LogEntry {
    param(
        [Parameter(Mandatory)][object]$Item,
        [Parameter(Mandatory)][string]$LogName
    )
    $oem = Get-ObjectProperty $Item 'Oem'
    $hpe = Get-ObjectProperty $oem 'Hpe'
    $severity = Get-ObjectProperty $Item 'Severity' (Get-ObjectProperty $hpe 'Severity' 'Unknown')
    [PSCustomObject][ordered]@{
        'Log' = $LogName
        'Created' = Get-ObjectProperty $Item 'Created' 'Unknown'
        'Severity' = $severity
        'Message' = Get-ObjectProperty $Item 'Message' ''
        'Repaired' = Get-ObjectProperty $hpe 'Repaired' 'N/A'
    }
}

function Convert-SecurityDashboardOverview {
    param([Parameter(Mandatory)][object]$Dashboard)

    [PSCustomObject][ordered]@{
        'Name' = 'Overall Security Status'
        'Security status' = Get-ObjectProperty $Dashboard 'OverallSecurityStatus' 'Unknown'
        'Current value' = "Server configuration lock: $(Get-ObjectProperty $Dashboard 'ServerConfigurationLockStatus' 'Unknown')"
    }
}

function Convert-SecurityParameter {
    param([Parameter(Mandatory)][object]$Item)

    $ignored = [bool](Get-ObjectProperty $Item 'Ignore' $false)
    [PSCustomObject][ordered]@{
        'Name' = Get-ObjectProperty $Item 'Name' (Get-ObjectProperty $Item 'Id' 'Security parameter')
        'Security status' = if ($ignored) { 'Ignored' } else { Get-ObjectProperty $Item 'SecurityStatus' 'Unknown' }
        'Current value' = Get-ObjectProperty $Item 'State' 'Unknown'
    }
}

function Convert-IloNetworkInterface {
    param([Parameter(Mandatory)][object]$Interface)

    $oem = Get-ObjectProperty $Interface 'Oem'
    $hpe = Get-ObjectProperty $oem 'Hpe'
    $interfaceType = [string](Get-ObjectProperty $hpe 'InterfaceType' '')
    $interfaceName = [string](Get-ObjectProperty $Interface 'Name' (Get-ObjectProperty $Interface 'Id' 'iLO network interface'))
    if (-not $interfaceType) {
        if ($interfaceName -match '(?i)dedicated') { $interfaceType = 'Dedicated' }
        elseif ($interfaceName -match '(?i)shared') { $interfaceType = 'Shared' }
        else { $interfaceType = 'Unknown' }
    }

    $enabled = Get-ObjectProperty $Interface 'InterfaceEnabled'
    if ($null -eq $enabled) { $enabled = Get-ObjectProperty $hpe 'NICEnabled' }
    $configuredText = if ($null -eq $enabled) { 'Unknown' } elseif ([bool]$enabled) { 'True' } else { 'False' }
    $ipv4Addresses = @(Get-ObjectProperty $Interface 'IPv4Addresses' @())
    $ipv4Address = @($ipv4Addresses | ForEach-Object { Get-ObjectProperty $_ 'Address' } | Where-Object { $_ }) -join '; '
    $subnetMask = @($ipv4Addresses | ForEach-Object { Get-ObjectProperty $_ 'SubnetMask' } | Where-Object { $_ }) -join '; '
    $gateway = @($ipv4Addresses | ForEach-Object { Get-ObjectProperty $_ 'Gateway' } | Where-Object { $_ }) -join '; '
    $addressOrigin = @($ipv4Addresses | ForEach-Object { Get-ObjectProperty $_ 'AddressOrigin' } | Where-Object { $_ }) -join '; '
    $nameServers = @(Get-ObjectProperty $Interface 'NameServers' @()) -join '; '
    $vlan = Get-ObjectProperty $Interface 'VLAN'
    $sharedOptions = Get-ObjectProperty $hpe 'SharedNetworkPortOptions'

    $rows = [Collections.Generic.List[object]]::new()
    foreach ($pair in @(
        @('Interface name', $interfaceName),
        @('Interface type', $interfaceType),
        @('Configured for iLO', $configuredText),
        @('Link status', (Get-ObjectProperty $Interface 'LinkStatus' 'Unknown')),
        @('Health', (Get-HealthValue $Interface)),
        @('Host name', (Get-ObjectProperty $Interface 'HostName' 'Unknown')),
        @('FQDN', (Get-ObjectProperty $Interface 'FQDN' 'Unknown')),
        @('MAC address', (Get-ObjectProperty $Interface 'MACAddress' 'Unknown')),
        @('Permanent MAC address', (Get-ObjectProperty $Interface 'PermanentMACAddress' 'Unknown')),
        @('Speed (Mbps)', (Get-ObjectProperty $Interface 'SpeedMbps' 'Unknown')),
        @('Full duplex', (Get-ObjectProperty $Interface 'FullDuplex' 'Unknown')),
        @('IPv4 address', $(if ($ipv4Address) { $ipv4Address } else { 'Not configured' })),
        @('IPv4 address origin', $(if ($addressOrigin) { $addressOrigin } else { 'Unknown' })),
        @('Subnet mask', $(if ($subnetMask) { $subnetMask } else { 'Unknown' })),
        @('Gateway', $(if ($gateway) { $gateway } else { 'Unknown' })),
        @('Name servers', $(if ($nameServers) { $nameServers } else { 'Not configured' })),
        @('VLAN enabled', (Get-ObjectProperty $vlan 'VLANEnable' 'Unknown')),
        @('VLAN ID', (Get-ObjectProperty $vlan 'VLANId' 'Not configured'))
    )) {
        $rows.Add([PSCustomObject][ordered]@{ Setting = $pair[0]; Value = $pair[1] })
    }
    if ($interfaceType -eq 'Shared') {
        $rows.Add([PSCustomObject][ordered]@{ Setting = 'Shared NIC'; Value = Get-ObjectProperty $sharedOptions 'NIC' 'Unknown' })
        $rows.Add([PSCustomObject][ordered]@{ Setting = 'Shared port'; Value = Get-ObjectProperty $sharedOptions 'Port' 'Unknown' })
    }
    if ($configuredText -eq 'False') {
        $rows.Add([PSCustomObject][ordered]@{ Setting = 'Assessment note'; Value = 'iLO is not configured to use this NIC.' })
    }
    return [PSCustomObject]@{
        InterfaceType = $interfaceType
        Rows = $rows.ToArray()
    }
}

function Get-IloHealthData {
    param(
        [Parameter(Mandatory)][object]$Session,
        [Parameter(Mandatory)][int]$MaxLogEntries
    )

    $notes = [System.Collections.Generic.List[string]]::new()
    $root = Invoke-RedfishGet -Session $Session -Uri '/redfish/v1/'
    $systems = @(Get-SafeCollectionFromUris `
        -Session $Session `
        -Uris @((Get-RedfishLink $root 'Systems'), '/redfish/v1/Systems/', '/redfish/v1/Systems') `
        -Notes $notes `
        -Label 'systems')
    if ($systems.Count -eq 0) {
        $discoveryDetail = @($notes | Where-Object { $_ -match '(?i)systems' }) -join ' '
        if (-not $discoveryDetail) { $discoveryDetail = 'The Systems collection returned no members.' }
        throw "No ComputerSystem resource was found. $discoveryDetail"
    }
    $system = $systems[0]
    $chassis = @(Get-SafeCollectionFromUris `
        -Session $Session `
        -Uris @((Get-RedfishLink $root 'Chassis'), '/redfish/v1/Chassis/', '/redfish/v1/Chassis') `
        -Notes $notes `
        -Label 'chassis')
    $managers = @(Get-SafeCollectionFromUris `
        -Session $Session `
        -Uris @((Get-RedfishLink $root 'Managers'), '/redfish/v1/Managers/', '/redfish/v1/Managers') `
        -Notes $notes `
        -Label 'managers')

    $iloInformation = @($managers | ForEach-Object { Convert-IloInformation $_ })
    $computeOpsManagement = @($managers | ForEach-Object { Convert-ComputeOpsManagement $_ })
    $remoteSupportRegistration = [Collections.Generic.List[object]]::new()
    foreach ($manager in $managers) {
        $managerUri = Get-ObjectProperty $manager '@odata.id'
        $managerOem = Get-ObjectProperty $manager 'Oem'
        $managerHpe = Get-ObjectProperty $managerOem 'Hpe'
        $managerHpeLinks = Get-ObjectProperty $managerHpe 'Links'
        $remoteSupportUri = Get-RedfishLink $managerHpeLinks 'RemoteSupport'
        $remoteSupportFallback = if ($managerUri) { "$($managerUri.TrimEnd('/'))/RemoteSupportService" } else { $null }
        $remoteSupport = Get-SafeResourceFromUris `
            -Session $Session `
            -Uris @($remoteSupportUri, $remoteSupportFallback) `
            -Notes $notes `
            -Label 'remote support registration'
        if ($null -ne $remoteSupport) {
            $remoteSupportRegistration.Add((Convert-RemoteSupportRegistration $remoteSupport))
        }
    }
    $statusInformation = [System.Collections.Generic.List[object]]::new()
    $statusInformation.Add([PSCustomObject][ordered]@{
        'Component' = 'Server'
        'Health' = Get-HealthValue $system
        'Detail' = "Power: $(Get-ObjectProperty $system 'PowerState' 'Unknown')"
    })
    foreach ($manager in $managers) {
        $statusInformation.Add([PSCustomObject][ordered]@{
            'Component' = Get-ObjectProperty $manager 'Name' (Get-ObjectProperty $manager 'Id' 'iLO Manager')
            'Health' = Get-HealthValue $manager
            'Detail' = "Firmware: $(Get-ObjectProperty $manager 'FirmwareVersion' 'Unknown')"
        })
    }

    $iloDedicatedNetworkPort = @()
    $iloSharedNetworkPort = @()
    foreach ($manager in $managers) {
        $managerUri = Get-ObjectProperty $manager '@odata.id'
        $ethernetInterfacesUri = Get-RedfishLink $manager 'EthernetInterfaces'
        if (-not $ethernetInterfacesUri -and $managerUri) {
            $ethernetInterfacesUri = "$($managerUri.TrimEnd('/'))/EthernetInterfaces"
        }
        foreach ($ethernetInterface in @(Get-SafeCollection $Session $ethernetInterfacesUri $notes 'iLO ethernet interfaces')) {
            $convertedInterface = Convert-IloNetworkInterface $ethernetInterface
            if ($convertedInterface.InterfaceType -eq 'Dedicated') {
                $iloDedicatedNetworkPort += $convertedInterface.Rows
            }
            elseif ($convertedInterface.InterfaceType -eq 'Shared') {
                $iloSharedNetworkPort += $convertedInterface.Rows
            }
        }
    }

    $temperatures = @()
    $fans = @()
    $powerSupplies = @()
    if ($chassis.Count -gt 0) {
        $chassisItem = $chassis[0]
        $chassisUri = Get-ObjectProperty $chassisItem '@odata.id'
        $thermalUri = Get-RedfishLink $chassisItem 'Thermal'
        $powerUri = Get-RedfishLink $chassisItem 'Power'
        if (-not $thermalUri -and $chassisUri) { $thermalUri = "$($chassisUri.TrimEnd('/'))/Thermal" }
        if (-not $powerUri -and $chassisUri) { $powerUri = "$($chassisUri.TrimEnd('/'))/Power" }
        try {
            $thermal = Invoke-RedfishGet $Session $thermalUri
            $temperatures = @(@(Get-ObjectProperty $thermal 'Temperatures' @()) |
                ForEach-Object { Convert-Temperature $_ } |
                Where-Object { Test-ReportRecordPresent $_ })
            $fans = @(@(Get-ObjectProperty $thermal 'Fans' @()) | ForEach-Object { Convert-Fan $_ })
        }
        catch { Add-CollectionNote $notes "Unable to collect thermal data: $($_.Exception.Message)" }
        try {
            $power = Invoke-RedfishGet $Session $powerUri
            $powerSupplies = @(@(Get-ObjectProperty $power 'PowerSupplies' @()) | ForEach-Object { Convert-PowerSupply $_ })
        }
        catch { Add-CollectionNote $notes "Unable to collect power data: $($_.Exception.Message)" }
    }

    $memory = @((Get-SafeCollection $Session (Get-RedfishLink $system 'Memory') $notes 'memory') |
        ForEach-Object { Convert-Memory $_ } |
        Where-Object { Test-ReportRecordPresent $_ })
    $processors = @((Get-SafeCollection $Session (Get-RedfishLink $system 'Processors') $notes 'processors') | ForEach-Object { Convert-Processor $_ })

    $systemNetwork = @()
    $systemUri = Get-ObjectProperty $system '@odata.id'
    $ethernetInterfacesUri = Get-RedfishLink $system 'EthernetInterfaces'
    if (-not $ethernetInterfacesUri -and $systemUri) {
        $ethernetInterfacesUri = "$($systemUri.TrimEnd('/'))/EthernetInterfaces"
    }
    $networkInterfacesUri = Get-RedfishLink $system 'NetworkInterfaces'
    if (-not $networkInterfacesUri -and $systemUri) {
        $networkInterfacesUri = "$($systemUri.TrimEnd('/'))/NetworkInterfaces"
    }
    $systemNetwork = @(Get-SafeCollectionFromUris `
        -Session $Session `
        -Uris @($ethernetInterfacesUri, $networkInterfacesUri) `
        -Notes $notes `
        -Label 'system network interfaces' | ForEach-Object {
            $odataType = [string](Get-ObjectProperty $_ '@odata.type' '')
            $interfaceType = if ($odataType -match '(?i)NetworkInterface') { 'Network interface' } else { 'Ethernet interface' }
            Convert-SystemNetworkInterface -Item $_ -Type $interfaceType
        })

    $deviceInventory = [System.Collections.Generic.List[object]]::new()
    $deviceUris = [System.Collections.Generic.List[string]]::new()
    foreach ($owner in $chassis) {
        $ownerUri = Get-ObjectProperty $owner '@odata.id'
        $ownerOem = Get-ObjectProperty $owner 'Oem'
        $ownerHpe = Get-ObjectProperty $ownerOem 'Hpe'
        $ownerHpeLinks = Get-ObjectProperty $ownerHpe 'Links'
        foreach ($uri in @(
            (Get-RedfishLink $ownerHpeLinks 'Devices'),
            $(if ($ownerUri) { "$($ownerUri.TrimEnd('/'))/Devices" }),
            (Get-RedfishLinkAny $owner 'PCIeDevices'),
            $(if ($ownerUri) { "$($ownerUri.TrimEnd('/'))/PCIeDevices" })
        )) {
            if ($uri -and -not $deviceUris.Contains($uri)) { $deviceUris.Add($uri) }
        }
    }
    if ($deviceUris.Count -eq 0) {
        Add-CollectionNote $notes 'Device inventory is not advertised by this system.'
    }
    else {
        foreach ($device in @(Get-SafeCollectionFromUris $Session $deviceUris.ToArray() $notes 'device inventory')) {
            $deviceInventory.Add((Convert-DeviceInventory $device))
        }
    }

    $storage = [System.Collections.Generic.List[object]]::new()
    $standardStorageUri = if ($systemUri) { "$($systemUri.TrimEnd('/'))/Storage" } else { $null }
    $storageUris = @($standardStorageUri, (Get-RedfishLink $system 'Storage'))
    foreach ($item in @(Get-SafeCollectionFromUris $Session $storageUris $notes 'storage')) {
        $storage.Add([PSCustomObject][ordered]@{
            'Name' = Get-ObjectProperty $item 'Name' (Get-ObjectProperty $item 'Id' 'Unknown')
            'Description' = Get-ObjectProperty $item 'Description' ''
            'Health' = Get-HealthValue $item
        })
        foreach ($relation in @('Drives', 'Volumes')) {
            $label = if ($relation -eq 'Drives') { 'Drive' } else { 'Volume' }
            $children = @()
            $inline = Get-ObjectProperty $item $relation
            if ($inline -is [System.Array]) {
                foreach ($reference in $inline) {
                    $childUri = Get-ObjectProperty $reference '@odata.id'
                    if ($childUri) {
                        try { $children += Invoke-RedfishGet $Session $childUri }
                        catch { Add-CollectionNote $notes "Unable to collect $label`: $($_.Exception.Message)" }
                    }
                }
            }
            else {
                $children = @(Get-SafeCollection $Session (Get-RedfishLink $item $relation) $notes $relation)
            }
            foreach ($child in $children) {
                $storage.Add([PSCustomObject][ordered]@{
                    'Name' = "$label`: $(Get-ObjectProperty $child 'Name' (Get-ObjectProperty $child 'Id' 'Unknown'))"
                    'Description' = Get-ObjectProperty $child 'Model' (Get-ObjectProperty $child 'VolumeType' '')
                    'Health' = Get-HealthValue $child
                })
            }
        }
    }

    $firmware = @()
    try {
        $updateUri = Get-RedfishLink $root 'UpdateService'
        if (-not $updateUri) { $updateUri = '/redfish/v1/UpdateService' }
        $updateService = Invoke-RedfishGet $Session $updateUri
        $firmware = @((Get-SafeCollection $Session (Get-RedfishLink $updateService 'FirmwareInventory') $notes 'firmware inventory') | ForEach-Object { Convert-Firmware $_ })
    }
    catch { Add-CollectionNote $notes "Unable to collect firmware inventory: $($_.Exception.Message)" }

    $eventLogs = [System.Collections.Generic.List[object]]::new()
    $logServiceUris = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($owner in @($system) + $managers) {
        $ownerUri = Get-ObjectProperty $owner '@odata.id'
        $advertisedLogUri = Get-RedfishLink $owner 'LogServices'
        if ($advertisedLogUri) {
            [void]$logServiceUris.Add($advertisedLogUri)
        }
        elseif ($ownerUri) {
            [void]$logServiceUris.Add("$($ownerUri.TrimEnd('/'))/LogServices")
        }
    }
    if ($logServiceUris.Count -eq 0) {
        Add-CollectionNote $notes 'Log services are not advertised by this system.'
    }
    else {
        foreach ($logServiceUri in $logServiceUris) {
            foreach ($service in @(Get-SafeCollection $Session $logServiceUri $notes 'log services')) {
                $logName = Get-ObjectProperty $service 'Name' (Get-ObjectProperty $service 'Id' 'Event log')
                foreach ($entry in @(Get-SafeCollection $Session (Get-RedfishLink $service 'Entries') $notes "$logName entries" $MaxLogEntries)) {
                    $eventLogs.Add((Convert-LogEntry $entry $logName))
                }
            }
        }
    }

    $securityDashboard = [System.Collections.Generic.List[object]]::new()
    foreach ($manager in $managers) {
        $managerUri = Get-ObjectProperty $manager '@odata.id'
        $securityServiceUri = Get-RedfishLinkAny $manager 'SecurityService'
        $securityServiceFallback = if ($managerUri) { "$($managerUri.TrimEnd('/'))/SecurityService" } else { $null }
        $securityService = Get-SafeResourceFromUris `
            -Session $Session `
            -Uris @($securityServiceUri, $securityServiceFallback) `
            -Notes $notes `
            -Label 'security service'
        if ($null -eq $securityService) { continue }

        $resolvedSecurityServiceUri = Get-ObjectProperty $securityService '@odata.id' $securityServiceUri
        $dashboardUri = Get-RedfishLinkAny $securityService 'SecurityDashboard'
        $dashboardFallback = if ($resolvedSecurityServiceUri) { "$($resolvedSecurityServiceUri.TrimEnd('/'))/SecurityDashboard" } else { $null }
        $dashboard = Get-SafeResourceFromUris `
            -Session $Session `
            -Uris @($dashboardUri, $dashboardFallback) `
            -Notes $notes `
            -Label 'security dashboard'
        if ($null -eq $dashboard) { continue }

        $securityDashboard.Add((Convert-SecurityDashboardOverview $dashboard))
        $resolvedDashboardUri = Get-ObjectProperty $dashboard '@odata.id' $dashboardUri
        $parametersUri = Get-RedfishLinkAny $dashboard 'SecurityParameters'
        $parametersFallback = if ($resolvedDashboardUri) { "$($resolvedDashboardUri.TrimEnd('/'))/SecurityParams" } else { $null }
        foreach ($parameter in @(Get-SafeCollectionFromUris `
            -Session $Session `
            -Uris @($parametersUri, $parametersFallback) `
            -Notes $notes `
            -Label 'security dashboard parameters')) {
            $securityDashboard.Add((Convert-SecurityParameter $parameter))
        }
    }

    [PSCustomObject][ordered]@{
        GeneratedAt = [DateTimeOffset]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ')
        Target = $Session.BaseUri.AbsoluteUri.TrimEnd('/')
        ServerStatus = Convert-ServerStatus $system
        IloInformation = $iloInformation
        StatusInformation = $statusInformation.ToArray()
        ComputeOpsManagement = $computeOpsManagement
        RemoteSupportRegistration = $remoteSupportRegistration.ToArray()
        Temperatures = $temperatures
        Fans = $fans
        PowerSupplies = $powerSupplies
        Storage = $storage.ToArray()
        Memory = $memory
        Processors = $processors
        SystemNetwork = $systemNetwork
        DeviceInventory = $deviceInventory.ToArray()
        Firmware = $firmware
        IloDedicatedNetworkPort = $iloDedicatedNetworkPort
        IloSharedNetworkPort = $iloSharedNetworkPort
        EventLogs = $eventLogs.ToArray()
        SecurityDashboard = $securityDashboard.ToArray()
        CollectionNotes = $notes.ToArray()
    }
}

function ConvertTo-WordColor {
    param([Parameter(Mandatory)][ValidatePattern('^[0-9A-Fa-f]{6}$')][string]$Hex)
    $red = [Convert]::ToInt32($Hex.Substring(0, 2), 16)
    $green = [Convert]::ToInt32($Hex.Substring(2, 2), 16)
    $blue = [Convert]::ToInt32($Hex.Substring(4, 2), 16)
    return $red + ($green * 256) + ($blue * 65536)
}

function Get-StatusColor {
    param([AllowNull()][object]$Value)
    $text = [string]$Value
    if ($text -match '(?i)critical|failed|fatal|risk') { return ConvertTo-WordColor 'D00000' }
    if ($text -match '(?i)warning|degraded|caution|recommended|ignored') { return ConvertTo-WordColor 'E59B00' }
    if ($text -match '(?i)^ok$|^enabled$|^healthy$|^connected$|^registered$') { return ConvertTo-WordColor '008A3B' }
    return ConvertTo-WordColor '555555'
}

function Get-AssessmentStatus {
    param([AllowEmptyCollection()][object[]]$Items)

    if (@($Items).Count -eq 0) { return $null }
    $signals = [System.Collections.Generic.List[string]]::new()
    foreach ($item in $Items) {
        $ignored = [bool](Get-ObjectProperty -InputObject $item -Name 'Ignored' -Default $false)
        foreach ($name in @('Health', 'HealthRollup', 'State', 'Status', 'Severity', 'SecurityStatus', 'OverallSecurityStatus', 'Security status', 'Connection status')) {
            $value = if ($item -is [Collections.IDictionary] -and $item.Contains($name)) {
                $item[$name]
            }
            else {
                Get-ObjectProperty -InputObject $item -Name $name
            }
            if ($null -ne $value -and [string]$value -notmatch '(?i)^\s*(unknown|absent)?\s*$') {
                if ($ignored -and $name -match '(?i)security' -and [string]$value -match '(?i)^risk$') {
                    $signals.Add('Ignored')
                }
                else {
                    $signals.Add([string]$value)
                }
            }
        }
    }
    if ($signals.Count -eq 0) { return $null }
    $evidence = $signals -join ' '
    if ($evidence -match '(?i)critical|failed|fatal|risk') { return 'CRITICAL' }
    if ($evidence -match '(?i)warning|degraded|caution|disabled|ignored|notconnected|notenabled|retryinprogress|connectioninprogress') { return 'RECOMMENDED' }
    return 'HEALTHY'
}

function Get-CriticalRecentEventLogs {
    param([Parameter(Mandatory)][object]$Data)

    $referenceTime = [DateTimeOffset]::Now
    $generatedAt = Get-ObjectProperty -InputObject $Data -Name 'GeneratedAt'
    if ($null -ne $generatedAt) {
        $parsedReference = [DateTimeOffset]::MinValue
        if ([DateTimeOffset]::TryParse([string]$generatedAt, [ref]$parsedReference)) {
            $referenceTime = $parsedReference
        }
    }
    $cutoff = $referenceTime.AddMonths(-1)

    return @(
        foreach ($event in @(Get-ObjectProperty -InputObject $Data -Name 'EventLogs' -Default @())) {
            $severity = [string](Get-ObjectProperty -InputObject $event -Name 'Severity' -Default '')
            if ($severity -notmatch '(?i)^critical$') { continue }

            $created = [DateTimeOffset]::MinValue
            $createdText = [string](Get-ObjectProperty -InputObject $event -Name 'Created' -Default '')
            if (-not [DateTimeOffset]::TryParse($createdText, [ref]$created)) { continue }
            if ($created -lt $cutoff -or $created -gt $referenceTime) { continue }
            $event
        }
    ) | Sort-Object {
        [DateTimeOffset]::Parse([string](Get-ObjectProperty -InputObject $_ -Name 'Created'))
    } -Descending
}

function Get-SecurityAssessmentStatus {
    param([AllowEmptyCollection()][object[]]$Items)

    $overallStatus = @($Items | Where-Object {
        [string](Get-ObjectProperty -InputObject $_ -Name 'Name' -Default '') -eq 'Overall Security Status'
    } | Select-Object -First 1)
    if ($overallStatus.Count -gt 0 -and
        [string](Get-ObjectProperty -InputObject $overallStatus[0] -Name 'Security status' -Default '') -match '(?i)^ignored$') {
        return 'HEALTHY'
    }
    $activeItems = @($Items | Where-Object {
        $ignored = [bool](Get-ObjectProperty -InputObject $_ -Name 'Ignored' -Default $false)
        $status = [string](Get-ObjectProperty -InputObject $_ -Name 'Security status' -Default '')
        -not $ignored -and $status -notmatch '(?i)^ignored$'
    })
    if ($activeItems.Count -eq 0 -and @($Items).Count -gt 0) { return 'HEALTHY' }
    return Get-AssessmentStatus $activeItems
}

function Get-IloNetworkPortAssessmentStatus {
    param(
        [AllowEmptyCollection()][object[]]$Rows,
        [switch]$OmitWhenUnconfigured
    )

    if (@($Rows).Count -eq 0) { return $null }
    $configuredRow = @($Rows | Where-Object Setting -eq 'Configured for iLO' | Select-Object -First 1)
    if ($configuredRow.Count -eq 0) { return $null }
    $configured = [string]$configuredRow[0].Value
    if ($configured -match '(?i)^false$') {
        if ($OmitWhenUnconfigured) { return $null }
        return 'IGNORED'
    }
    if ($configured -notmatch '(?i)^true$') { return $null }

    $healthRow = @($Rows | Where-Object Setting -eq 'Health' | Select-Object -First 1)
    $linkStatusRow = @($Rows | Where-Object Setting -eq 'Link status' | Select-Object -First 1)
    $health = if ($healthRow.Count -gt 0) { [string](Get-ObjectProperty $healthRow[0] 'Value' '') } else { '' }
    $linkStatus = if ($linkStatusRow.Count -gt 0) { [string](Get-ObjectProperty $linkStatusRow[0] 'Value' '') } else { '' }
    if ("$health$linkStatus" -match '(?i)^\s*(unknown)?\s*$') { return $null }
    if ("$health $linkStatus" -match '(?i)critical|failed|fatal') { return 'CRITICAL' }
    if ("$health $linkStatus" -match '(?i)warning|degraded|unknown|nolink|linkdown') { return 'RECOMMENDED' }
    return 'HEALTHY'
}

function Get-RemoteSupportAssessmentStatus {
    param([AllowEmptyCollection()][object[]]$Rows)

    if (@($Rows).Count -eq 0) { return $null }
    $registration = [string](Get-ObjectProperty $Rows[0] 'Registration' '')
    $lastError = [string](Get-ObjectProperty $Rows[0] 'Last transmission error' '')
    if ($lastError -and $lastError -notmatch '(?i)^(none|no error|unknown)$') { return 'RECOMMENDED' }
    if ($registration -match '(?i)^registered$') { return 'HEALTHY' }
    if ($registration -match '(?i)^not registered$') { return 'RECOMMENDED' }
    return $null
}

function New-AssessmentSummary {
    param([Parameter(Mandatory)][object]$Data)

    $criticalRecentEvents = @(Get-CriticalRecentEventLogs $Data)
    $securityEvents = @($criticalRecentEvents | Where-Object {
        $_.Log -match '(?i)security' -or $_.Message -match '(?i)security|unauthorized|authentication'
    })
    $securityDashboard = @((Get-ObjectProperty $Data 'SecurityDashboard' @()))
    $securityEvidence = $securityDashboard + $securityEvents
    $powerThermal = @((Get-ObjectProperty $Data 'Temperatures' @())) +
        @((Get-ObjectProperty $Data 'Fans' @())) +
        @((Get-ObjectProperty $Data 'PowerSupplies' @()))
    $performance = @((Get-ObjectProperty $Data 'Memory' @())) +
        @((Get-ObjectProperty $Data 'Processors' @()))
    $informationEvidence = @((Get-ObjectProperty $Data 'ServerStatus' @())) +
        @((Get-ObjectProperty $Data 'IloInformation' @())) +
        @((Get-ObjectProperty $Data 'StatusInformation' @())) +
        @((Get-ObjectProperty $Data 'ComputeOpsManagement' @()))
    $systemEvidence = @((Get-ObjectProperty $Data 'ServerStatus' @())) +
        @((Get-ObjectProperty $Data 'Processors' @())) +
        @((Get-ObjectProperty $Data 'Memory' @())) +
        @((Get-ObjectProperty $Data 'SystemNetwork' @())) +
        @((Get-ObjectProperty $Data 'DeviceInventory' @())) +
        @((Get-ObjectProperty $Data 'Storage' @()))
    $candidates = @(
        [PSCustomObject]@{ Section = 'Information'; Status = Get-AssessmentStatus $informationEvidence }
        [PSCustomObject]@{ Section = 'System Information'; Status = Get-AssessmentStatus $systemEvidence }
        [PSCustomObject]@{ Section = 'Firmware & OS Software'; Status = Get-AssessmentStatus @((Get-ObjectProperty $Data 'Firmware' @())) }
        [PSCustomObject]@{ Section = 'Power & Thermal'; Status = Get-AssessmentStatus $powerThermal }
        [PSCustomObject]@{ Section = 'Performance'; Status = Get-AssessmentStatus $performance }
        [PSCustomObject]@{ Section = 'iLO Dedicated Network Port'; Status = Get-IloNetworkPortAssessmentStatus @((Get-ObjectProperty $Data 'IloDedicatedNetworkPort' @())) }
        [PSCustomObject]@{ Section = 'iLO Shared Network Port'; Status = Get-IloNetworkPortAssessmentStatus @((Get-ObjectProperty $Data 'IloSharedNetworkPort' @())) -OmitWhenUnconfigured }
        [PSCustomObject]@{ Section = 'Remote Support'; Status = Get-RemoteSupportAssessmentStatus @((Get-ObjectProperty $Data 'RemoteSupportRegistration' @())) }
        [PSCustomObject]@{ Section = 'Security Dashboard'; Status = if ($securityEvents.Count -gt 0) { Get-AssessmentStatus $securityEvidence } else { Get-SecurityAssessmentStatus $securityDashboard } }
    )
    return @($candidates | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.Status) })
}

function Get-RecommendedActionText {
    param(
        [Parameter(Mandatory)][object]$Data,
        [Parameter(Mandatory)][object[]]$Assessment
    )

    $actions = [Collections.Generic.List[string]]::new()
    if (@($Assessment | Where-Object { $_.Section -eq 'Security Dashboard' -and $_.Status -eq 'CRITICAL' }).Count -gt 0) {
        $actions.Add('Review the iLO Security Dashboard and remediate all active security risks.')
    }
    if (@(Get-CriticalRecentEventLogs $Data).Count -gt 0) {
        $actions.Add('Investigate and resolve the critical iLO event-log entries recorded during the previous month.')
    }
    $otherCritical = @($Assessment | Where-Object {
        $_.Status -eq 'CRITICAL' -and $_.Section -ne 'Security Dashboard'
    } | Select-Object -ExpandProperty Section)
    if ($otherCritical.Count -gt 0) {
        $actions.Add("Review and remediate the critical findings in: $($otherCritical -join ', ').")
    }
    if (@($Assessment | Where-Object Status -eq 'RECOMMENDED').Count -gt 0) {
        $actions.Add('Review the sections marked RECOMMENDED in the Assessment Summary and validate unavailable or uncollected configuration areas.')
    }
    if ($actions.Count -eq 0) {
        $actions.Add('No immediate corrective action is required. Continue routine monitoring and periodic health checks.')
    }
    return $actions -join ' '
}

function Get-EndRange {
    param([Parameter(Mandatory)][object]$Document)
    return $Document.Range($Document.Content.End - 1, $Document.Content.End - 1)
}

function Add-WordParagraph {
    param(
        [Parameter(Mandatory)][object]$Document,
        [Parameter(Mandatory)][string]$Text,
        [double]$Size = 11,
        [bool]$Bold = $false,
        [string]$Color = '222222',
        [double]$SpaceAfter = 6,
        [bool]$Italic = $false,
        [ValidateSet('Left', 'Center', 'Right')][string]$Alignment = 'Left',
        [double]$SpaceBefore = 0
    )
    $range = Get-EndRange $Document
    $range.Text = "$Text`r"
    $paragraph = $range.Paragraphs.Item(1)
    $paragraph.Range.Font.Name = 'Aptos'
    $paragraph.Range.Font.Size = $Size
    $paragraph.Range.Font.Bold = [int]$Bold
    $paragraph.Range.Font.Italic = [int]$Italic
    $paragraph.Range.Font.Color = ConvertTo-WordColor $Color
    $paragraph.Format.Alignment = switch ($Alignment) { 'Center' { 1 } 'Right' { 2 } default { 0 } }
    $paragraph.Format.SpaceBefore = $SpaceBefore
    $paragraph.Format.SpaceAfter = $SpaceAfter
    $paragraph.Format.LineSpacingRule = 0
}

function Add-WordHeading {
    param(
        [Parameter(Mandatory)][object]$Document,
        [Parameter(Mandatory)][string]$Text,
        [ValidateSet(1, 2)][int]$Level = 2
    )
    $range = Get-EndRange $Document
    $range.Text = "$Text`r"
    $paragraph = $range.Paragraphs.Item(1)
    $paragraph.Range.Style = "Heading $Level"
    $paragraph.Range.Font.Name = 'Aptos Display'
    $paragraph.Range.Font.Size = if ($Level -eq 1) { 16 } else { 13 }
    $paragraph.Range.Font.Bold = -1
    $paragraph.Range.Font.Color = ConvertTo-WordColor $(if ($Level -eq 1) { $script:ReportBlue } else { $script:ReportDark })
    $paragraph.Format.SpaceBefore = if ($Level -eq 1) { 18 } else { 8 }
    $paragraph.Format.SpaceAfter = if ($Level -eq 1) { 6 } else { 3 }
    $paragraph.Format.KeepWithNext = -1
    if ($Level -eq 1) {
        $border = $paragraph.Borders.Item($script:WdBorderBottom)
        $border.LineStyle = 1
        $border.LineWidth = 4
        $border.Color = ConvertTo-WordColor $script:ReportBlue
    }
}

function Add-WordTable {
    param(
        [Parameter(Mandatory)][object]$Document,
        [AllowEmptyCollection()][object[]]$Records,
        [Parameter(Mandatory)][string]$EmptyMessage
    )

    if ($Records.Count -eq 0) {
        Add-WordParagraph $Document $EmptyMessage 9.5 $false '666666' 6
        return
    }
    $properties = @($Records[0].PSObject.Properties.Name)
    $range = Get-EndRange $Document
    $table = $Document.Tables.Add($range, $Records.Count + 1, $properties.Count)
    $table.Style = 'Table Grid'
    $table.AllowAutoFit = $false
    $table.PreferredWidthType = $script:WdPreferredWidthPoints
    $table.PreferredWidth = 540
    $table.Rows.Item(1).HeadingFormat = -1

    $scores = foreach ($name in $properties) {
        $lengths = @($name.Length) + @($Records | ForEach-Object { ([string]$_.PSObject.Properties[$name].Value).Length })
        [Math]::Max(10, [Math]::Min(36, ($lengths | Measure-Object -Maximum).Maximum))
    }
    $totalScore = ($scores | Measure-Object -Sum).Sum
    for ($column = 1; $column -le $properties.Count; $column++) {
        $table.Columns.Item($column).Width = [Math]::Max(45, 540 * $scores[$column - 1] / $totalScore)
    }

    for ($column = 1; $column -le $properties.Count; $column++) {
        $cell = $table.Cell(1, $column)
        $cell.Range.Text = $properties[$column - 1]
        $cell.Range.Font.Name = 'Aptos'
        $cell.Range.Font.Size = 9
        $cell.Range.Font.Bold = -1
        $cell.Range.Font.Color = ConvertTo-WordColor 'FFFFFF'
        $cell.Shading.BackgroundPatternColor = ConvertTo-WordColor $script:ReportBlue
        $cell.VerticalAlignment = 1
    }

    for ($row = 1; $row -le $Records.Count; $row++) {
        $table.Rows.Item($row + 1).AllowBreakAcrossPages = 0
        for ($column = 1; $column -le $properties.Count; $column++) {
            $name = $properties[$column - 1]
            $value = $Records[$row - 1].PSObject.Properties[$name].Value
            $cell = $table.Cell($row + 1, $column)
            $cell.Range.Text = [string]$value
            $cell.Range.Font.Name = 'Aptos'
            $cell.Range.Font.Size = 9
            $statusColumns = @('Health', 'Status', 'Severity', 'State', 'Security status', 'SecurityStatus', 'Registration')
            $cell.Range.Font.Bold = if ($name -in $statusColumns) { -1 } else { 0 }
            $cell.Range.Font.Color = if ($name -in $statusColumns) { Get-StatusColor $value } else { ConvertTo-WordColor '222222' }
            if ($row % 2 -eq 0) { $cell.Shading.BackgroundPatternColor = ConvertTo-WordColor $script:ReportStripe }
            $cell.VerticalAlignment = 1
        }
    }
    $table.TopPadding = 4
    $table.BottomPadding = 4
    $table.LeftPadding = 6
    $table.RightPadding = 6
    $afterTable = $table.Range
    $afterTable.Collapse(0)
    $afterTable.InsertParagraphAfter()
}

function Add-AssessmentSummaryTable {
    param(
        [Parameter(Mandatory)][object]$Document,
        [Parameter(Mandatory)][object[]]$Assessment
    )

    $range = Get-EndRange $Document
    $dataRowCount = [Math]::Max(1, [Math]::Ceiling($Assessment.Count / 2.0))
    $table = $Document.Tables.Add($range, $dataRowCount + 1, 4)
    $table.Style = 'Table Grid'
    $table.AllowAutoFit = $false
    $table.PreferredWidthType = $script:WdPreferredWidthPoints
    $table.PreferredWidth = 540
    $widths = @(175, 95, 175, 95)
    for ($column = 1; $column -le 4; $column++) {
        $table.Columns.Item($column).Width = $widths[$column - 1]
        $header = $table.Cell(1, $column)
        $header.Range.Text = if ($column % 2 -eq 1) { 'Section' } else { 'Severity' }
        $header.Range.Font.Name = 'Aptos'
        $header.Range.Font.Size = 9
        $header.Range.Font.Bold = -1
        $header.Range.Font.Color = ConvertTo-WordColor 'FFFFFF'
        $header.Shading.BackgroundPatternColor = ConvertTo-WordColor $script:ReportBlue
        $header.VerticalAlignment = 1
    }
    $table.Rows.Item(1).HeadingFormat = -1

    for ($row = 0; $row -lt $dataRowCount; $row++) {
        $leftIndex = $row * 2
        $rightIndex = $leftIndex + 1
        $left = if ($leftIndex -lt $Assessment.Count) { $Assessment[$leftIndex] } else { $null }
        $right = if ($rightIndex -lt $Assessment.Count) { $Assessment[$rightIndex] } else { $null }
        $values = @(
            $(if ($null -ne $left) { $left.Section } else { '' }),
            $(if ($null -ne $left) { $left.Status } else { '' }),
            $(if ($null -ne $right) { $right.Section } else { '' }),
            $(if ($null -ne $right) { $right.Status } else { '' })
        )
        for ($column = 1; $column -le 4; $column++) {
            $cell = $table.Cell($row + 2, $column)
            $cell.Range.Text = [string]$values[$column - 1]
            $cell.Range.Font.Name = 'Aptos'
            $cell.Range.Font.Size = 9
            if ($column % 2 -eq 1) {
                $cell.Range.Font.Color = ConvertTo-WordColor '0070C0'
                $cell.Range.Font.Underline = 1
            }
            else {
                $cell.Range.Font.Bold = -1
                $cell.Range.Font.Color = Get-StatusColor $values[$column - 1]
            }
            if ($row % 2 -eq 1) {
                $cell.Shading.BackgroundPatternColor = ConvertTo-WordColor $script:ReportStripe
            }
            $cell.VerticalAlignment = 1
        }
        $table.Rows.Item($row + 2).AllowBreakAcrossPages = 0
    }
    $table.TopPadding = 4
    $table.BottomPadding = 4
    $table.LeftPadding = 6
    $table.RightPadding = 6
    $afterTable = $table.Range
    $afterTable.Collapse(0)
    $afterTable.InsertParagraphAfter()
}

function Set-WordReportFurniture {
    param(
        [Parameter(Mandatory)][object]$Section,
        [Parameter(Mandatory)][string]$Target,
        [Parameter(Mandatory)][string]$CustomerName,
        [Parameter(Mandatory)][string]$ReportDate
    )

    if (-not (Test-Path $script:ReportLogoPath)) {
        throw "The WinslowTG logo is missing: $script:ReportLogoPath"
    }
    $Section.PageSetup.DifferentFirstPageHeaderFooter = 0
    $Section.PageSetup.OddAndEvenPagesHeaderFooter = 0

    # Populate every header/footer variant for template compatibility.
    foreach ($headerIndex in 1..3) {
        $Section.Headers.Item($headerIndex).LinkToPrevious = $false
        $header = $Section.Headers.Item($headerIndex).Range
        $header.Text = "`tHP iLO Check - $Target - $CustomerName"
        $header.Font.Name = 'Aptos'
        $header.Font.Size = 9
        $header.Font.Color = ConvertTo-WordColor $script:ReportDark
        $header.ParagraphFormat.TabStops.Add(540, 2, 0) | Out-Null
        $logoRange = $header.Duplicate
        $logoRange.SetRange($header.Start, $header.Start)
        [void]$header.InlineShapes.AddPicture($script:ReportLogoPath, $false, $true, $logoRange)
        $headerBorder = $header.Paragraphs.Item(1).Borders.Item($script:WdBorderBottom)
        $headerBorder.LineStyle = 1
        $headerBorder.LineWidth = 4
        $headerBorder.Color = ConvertTo-WordColor $script:ReportBlue
    }

    foreach ($footerIndex in 1..3) {
        $Section.Footers.Item($footerIndex).LinkToPrevious = $false
        $footer = $Section.Footers.Item($footerIndex).Range
        $footer.Text = "$([char]0x00A9)2026 Winslow Tech Group. All Right Reserved`tPage "
        $footer.Font.Name = 'Aptos'
        $footer.Font.Size = 9
        $footer.Font.Color = ConvertTo-WordColor $script:ReportDark
        $footer.ParagraphFormat.TabStops.Add(540, 2, 0) | Out-Null
        $pageRange = $footer.Duplicate
        $pageRange.SetRange($pageRange.End - 1, $pageRange.End - 1)
        $pageRange.Fields.Add($pageRange, $script:WdFieldPage) | Out-Null
        $suffixRange = $footer.Duplicate
        $suffixRange.SetRange($suffixRange.End - 1, $suffixRange.End - 1)
        $suffixRange.InsertAfter(' of ')
        $totalRange = $footer.Duplicate
        $totalRange.SetRange($totalRange.End - 1, $totalRange.End - 1)
        $totalRange.Fields.Add($totalRange, $script:WdFieldNumPages) | Out-Null
        $footerBorder = $footer.Paragraphs.Item(1).Borders.Item($script:WdBorderTop)
        $footerBorder.LineStyle = 1
        $footerBorder.LineWidth = 4
        $footerBorder.Color = ConvertTo-WordColor $script:ReportBlue
    }
}

function Add-WordCover {
    param(
        [Parameter(Mandatory)][object]$Document,
        [Parameter(Mandatory)][string]$Target,
        [Parameter(Mandatory)][string]$CustomerName,
        [Parameter(Mandatory)][string]$ReportDate
    )

    Add-WordParagraph $Document 'HP iLO Health Check' 28 $true $script:ReportBlue 10 $false 'Center' 144
    Add-WordParagraph $Document $Target 20 $false $script:ReportDark 6 $false 'Center'
    Add-WordParagraph $Document $CustomerName 16 $false $script:ReportDark 6 $false 'Center'
    Add-WordParagraph $Document $ReportDate 13 $false '888888' 0 $false 'Center'
    Add-WordParagraph $Document '[Confidential]' 11 $false 'AAAAAA' 0 $true 'Center' 12
    $breakRange = Get-EndRange $Document
    $breakRange.InsertBreak($script:WdPageBreak)
}

function ConvertTo-OpenXmlText {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) { return '' }
    return [Security.SecurityElement]::Escape([string]$Value)
}

function New-OpenXmlRun {
    param(
        [AllowNull()][object]$Text,
        [string]$Font = 'Aptos',
        [double]$Size = 10,
        [string]$Color = '222222',
        [switch]$Bold,
        [switch]$Italic
    )

    $properties = @(
        '<w:rFonts w:ascii="' + $Font + '" w:hAnsi="' + $Font + '"/>'
        '<w:sz w:val="' + [int]($Size * 2) + '"/>'
        '<w:szCs w:val="' + [int]($Size * 2) + '"/>'
        '<w:color w:val="' + $Color + '"/>'
    )
    if ($Bold) { $properties += '<w:b/>' }
    if ($Italic) { $properties += '<w:i/>' }
    return '<w:r><w:rPr>' + ($properties -join '') + '</w:rPr><w:t xml:space="preserve">' +
        (ConvertTo-OpenXmlText $Text) + '</w:t></w:r>'
}

function New-OpenXmlParagraph {
    param(
        [AllowNull()][object]$Text,
        [string]$Style = 'Normal',
        [string]$Alignment = 'left',
        [int]$Before = 0,
        [int]$After = 120,
        [switch]$Bold,
        [switch]$Italic,
        [string]$Color = '222222',
        [double]$Size = 10,
        [switch]$KeepNext,
        [switch]$PageBreakAfter
    )

    $paragraphProperties = @(
        '<w:pStyle w:val="' + $Style + '"/>'
        '<w:jc w:val="' + $Alignment + '"/>'
        '<w:spacing w:before="' + $Before + '" w:after="' + $After + '" w:line="276" w:lineRule="auto"/>'
    )
    if ($KeepNext) { $paragraphProperties += '<w:keepNext/>' }
    $xml = '<w:p><w:pPr>' + ($paragraphProperties -join '') + '</w:pPr>' +
        (New-OpenXmlRun -Text $Text -Size $Size -Color $Color -Bold:$Bold -Italic:$Italic)
    if ($PageBreakAfter) { $xml += '<w:r><w:br w:type="page"/></w:r>' }
    return $xml + '</w:p>'
}

function Get-OpenXmlStatusColor {
    param([AllowNull()][object]$Status)

    switch -Regex ([string]$Status) {
        '^(CRITICAL|Critical|Failed|Fatal|Risk)$' { return 'C00000' }
        '^(RECOMMENDED|Warning|Degraded|Caution|Ignored)$' { return 'BF7200' }
        '^(HEALTHY|OK|Ok|Enabled|Connected)$' { return '00843D' }
        default { return '666666' }
    }
}

function New-OpenXmlCell {
    param(
        [AllowNull()][object]$Value,
        [int]$Width,
        [string]$Fill = 'FFFFFF',
        [string]$Color = '222222',
        [switch]$Bold,
        [string]$Alignment = 'left',
        [double]$Size = 8.5
    )

    $verticalAlignment = if ($Alignment -eq 'center') { 'center' } else { 'left' }
    return @"
<w:tc><w:tcPr><w:tcW w:w="$Width" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="$Fill"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="$verticalAlignment"/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>$(New-OpenXmlRun -Text $Value -Size $Size -Color $Color -Bold:$Bold)</w:p></w:tc>
"@
}

function Get-OpenXmlColumnWidths {
    param(
        [Parameter(Mandatory)][object[]]$Rows,
        [Parameter(Mandatory)][string[]]$Columns,
        [int]$TotalWidth = 10800
    )

    if ($Columns.Count -eq 1) { return ,$TotalWidth }
    $weights = foreach ($column in $Columns) {
        $maximum = [Math]::Max(8, [Math]::Min(32, $column.Length))
        foreach ($row in $Rows | Select-Object -First 30) {
            $valueLength = ([string]$row.$column).Length
            $maximum = [Math]::Max($maximum, [Math]::Min(32, $valueLength))
        }
        $maximum
    }
    $minimumWidth = if ($Columns.Count -le 8) { 900 } else {
        [int][Math]::Floor(($TotalWidth / $Columns.Count) * 0.80)
    }
    $remainingWidth = $TotalWidth - ($minimumWidth * $Columns.Count)
    $weightTotal = ($weights | Measure-Object -Sum).Sum
    $widths = @($weights | ForEach-Object {
        $minimumWidth + [int][Math]::Floor($remainingWidth * ($_ / $weightTotal))
    })
    $currentTotal = ($widths | Measure-Object -Sum).Sum
    if ($currentTotal -ne $TotalWidth) {
        $widths[$widths.Count - 1] += ($TotalWidth - $currentTotal)
    }
    return $widths
}

function New-OpenXmlTable {
    param(
        [Parameter(Mandatory)][object[]]$Rows,
        [string[]]$Columns,
        [int[]]$Widths,
        [switch]$StatusColumns
    )

    if (-not $Columns -or $Columns.Count -eq 0) {
        $Columns = @($Rows[0].PSObject.Properties.Name)
    }
    if (-not $Widths -or $Widths.Count -ne $Columns.Count) {
        $Widths = @(Get-OpenXmlColumnWidths -Rows $Rows -Columns $Columns)
    }

    $grid = ($Widths | ForEach-Object { '<w:gridCol w:w="' + $_ + '"/>' }) -join ''
    $headerCells = for ($index = 0; $index -lt $Columns.Count; $index++) {
        New-OpenXmlCell -Value $Columns[$index] -Width $Widths[$index] -Fill '005F9E' -Color 'FFFFFF' -Bold -Size 8.5
    }
    $tableRows = @(
        '<w:tr><w:trPr><w:tblHeader/></w:trPr>' + ($headerCells -join '') + '</w:tr>'
    )
    for ($rowIndex = 0; $rowIndex -lt $Rows.Count; $rowIndex++) {
        $fill = if ($rowIndex % 2 -eq 1) { 'EEF4F8' } else { 'FFFFFF' }
        $cells = for ($columnIndex = 0; $columnIndex -lt $Columns.Count; $columnIndex++) {
            $column = $Columns[$columnIndex]
            $value = $Rows[$rowIndex].$column
            $isStatus = $StatusColumns -and ($column -match '(?i)status|health|state|severity')
            $alignment = if ($isStatus) { 'center' } else { 'left' }
            $color = if ($isStatus) { Get-OpenXmlStatusColor $value } else { '222222' }
            New-OpenXmlCell -Value $value -Width $Widths[$columnIndex] -Fill $fill -Color $color -Bold:$isStatus -Alignment $alignment
        }
        $tableRows += '<w:tr>' + ($cells -join '') + '</w:tr>'
    }

    return @"
<w:tbl><w:tblPr><w:tblW w:w="10800" w:type="dxa"/><w:tblInd w:w="0" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="80" w:type="dxa"/><w:start w:w="120" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:end w:w="120" w:type="dxa"/></w:tblCellMar><w:tblBorders><w:top w:val="single" w:sz="4" w:color="C8D5DF"/><w:left w:val="single" w:sz="4" w:color="C8D5DF"/><w:bottom w:val="single" w:sz="4" w:color="C8D5DF"/><w:right w:val="single" w:sz="4" w:color="C8D5DF"/><w:insideH w:val="single" w:sz="4" w:color="DCE5EB"/><w:insideV w:val="single" w:sz="4" w:color="DCE5EB"/></w:tblBorders></w:tblPr><w:tblGrid>$grid</w:tblGrid>$($tableRows -join '')</w:tbl>
$(New-OpenXmlParagraph -Text '' -After 80 -Size 2)
"@
}

function New-OpenXmlBulletParagraph {
    param([Parameter(Mandatory)][string]$Text)

    return '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:spacing w:after="80" w:line="276" w:lineRule="auto"/></w:pPr>' +
        (New-OpenXmlRun -Text $Text -Size 9 -Color '666666') + '</w:p>'
}

function Add-OpenXmlPackageEntry {
    param(
        [Parameter(Mandatory)][object]$Archive,
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Content
    )

    $entry = $Archive.CreateEntry($Path)
    $stream = $entry.Open()
    $writer = $null
    try {
        $writer = New-Object IO.StreamWriter($stream, (New-Object Text.UTF8Encoding($false)))
        $writer.Write($Content)
    }
    finally {
        if ($null -ne $writer) { $writer.Dispose() }
        else { $stream.Dispose() }
    }
}

function Add-OpenXmlBinaryPackageEntry {
    param(
        [Parameter(Mandatory)][object]$Archive,
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$SourcePath
    )

    $entry = $Archive.CreateEntry($Path)
    $destination = $entry.Open()
    $source = $null
    try {
        $source = [IO.File]::OpenRead($SourcePath)
        $source.CopyTo($destination)
    }
    finally {
        if ($null -ne $source) { $source.Dispose() }
        $destination.Dispose()
    }
}

function New-OpenXmlLogoRun {
    return @'
<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="751840" cy="228600"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="1" name="Winslow Technology Group logo" descr="Winslow Technology Group logo"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="winslowtg-logo.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="751840" cy="228600"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>
'@
}

function New-OpenXmlHealthReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Data,
        [Parameter(Mandatory)][string]$OutputPath,
        [Parameter(Mandatory)][string]$CustomerName
    )

    $resolved = [IO.Path]::GetFullPath($OutputPath)
    $directory = Split-Path -Parent $resolved
    if (-not (Test-Path $directory)) { New-Item -ItemType Directory -Path $directory -Force | Out-Null }
    if (-not (Test-Path $script:ReportLogoPath)) {
        throw "The WinslowTG logo is missing: $script:ReportLogoPath"
    }

    $reportDate = Get-Date -Format 'MMMM d, yyyy'
    $displayTarget = ([uri]$Data.Target).Host
    $assessment = @(New-AssessmentSummary $Data)
    $recommendedAction = Get-RecommendedActionText -Data $Data -Assessment $assessment
    $body = [Collections.Generic.List[string]]::new()

    [void]$body.Add((New-OpenXmlParagraph -Text 'HP iLO Health Check' -Style 'Title' -Alignment center -Before 2300 -After 180 -Bold -Color '005F9E' -Size 28))
    [void]$body.Add((New-OpenXmlParagraph -Text $displayTarget -Style 'Subtitle' -Alignment center -After 100 -Color '203647' -Size 20))
    [void]$body.Add((New-OpenXmlParagraph -Text $CustomerName -Style 'Subtitle' -Alignment center -After 100 -Color '203647' -Size 16))
    [void]$body.Add((New-OpenXmlParagraph -Text $reportDate -Alignment center -After 80 -Color '777777' -Size 12))
    [void]$body.Add((New-OpenXmlParagraph -Text '[Confidential]' -Alignment center -After 0 -Italic -Color '999999' -Size 10 -PageBreakAfter))

    [void]$body.Add((New-OpenXmlParagraph -Text 'Executive Overview' -Style 'Heading1' -Before 0 -After 160 -Bold -Color '005F9E' -Size 16 -KeepNext))
    [void]$body.Add((New-OpenXmlParagraph -Text "$CustomerName engaged Professional Services to conduct an HP iLO Health Check of $displayTarget. This report documents the discovery, analysis, and recommendations from the assessment." -After 160 -Size 10.5))
    [void]$body.Add((New-OpenXmlParagraph -Text 'Recommended Action' -Style 'Heading2' -Before 120 -After 100 -Bold -Color '005F9E' -Size 13 -KeepNext))
    [void]$body.Add((New-OpenXmlParagraph -Text $recommendedAction -After 160 -Size 10.5))
    [void]$body.Add((New-OpenXmlParagraph -Text 'Assessment Summary' -Style 'Heading2' -Before 160 -After 100 -Bold -Color '005F9E' -Size 13 -KeepNext))
    $assessmentRowCount = [Math]::Max(1, [Math]::Ceiling($assessment.Count / 2.0))
    $assessmentRows = for ($index = 0; $index -lt $assessmentRowCount; $index++) {
        $leftIndex = $index * 2
        $rightIndex = $leftIndex + 1
        $left = if ($leftIndex -lt $assessment.Count) { $assessment[$leftIndex] } else { $null }
        $right = if ($rightIndex -lt $assessment.Count) { $assessment[$rightIndex] } else { $null }
        [PSCustomObject][ordered]@{
            Assessment = if ($null -ne $left) { $left.Section } else { '' }
            Severity = if ($null -ne $left) { $left.Status } else { '' }
            'Assessment ' = if ($null -ne $right) { $right.Section } else { '' }
            'Severity ' = if ($null -ne $right) { $right.Status } else { '' }
        }
    }
    [void]$body.Add((New-OpenXmlTable -Rows $assessmentRows -Columns @('Assessment', 'Severity', 'Assessment ', 'Severity ') -Widths @(3460, 1940, 3460, 1940) -StatusColumns))

    [void]$body.Add((New-OpenXmlParagraph -Text 'Information' -Style 'Heading1' -Before 220 -After 120 -Bold -Color '005F9E' -Size 16 -KeepNext))
    foreach ($item in @(
        @('Server', 'ServerStatus', $true),
        @('iLO', 'IloInformation', $false),
        @('Status', 'StatusInformation', $false),
        @('HPE Compute Ops Management', 'ComputeOpsManagement', $false)
    )) {
        $records = @(Get-ReportRecords -Data $Data -PropertyName $item[1] -PropertyMap:([bool]$item[2]) -ExcludeColumns $(if ($item[0] -eq 'Status') { @('State') } elseif ($item[0] -eq 'HPE Compute Ops Management') { @('Next retry time') } else { @() }))
        if ($records.Count -eq 0) { continue }
        [void]$body.Add((New-OpenXmlParagraph -Text $item[0] -Style 'Heading2' -Before 160 -After 100 -Bold -Color '005F9E' -Size 13 -KeepNext))
        $tableArguments = @{ Rows = $records; StatusColumns = $true }
        if ([bool]$item[2]) { $tableArguments.Columns = @('Item', 'Value'); $tableArguments.Widths = @(3000, 7800) }
        [void]$body.Add((New-OpenXmlTable @tableArguments))
    }

    $remoteSupportRows = @(Get-ReportRecords -Data $Data -PropertyName 'RemoteSupportRegistration')
    if ($remoteSupportRows.Count -gt 0) {
        $registrationRows = @(Get-ReportPropertyRows $remoteSupportRows[0])
        [void]$body.Add((New-OpenXmlParagraph -Text 'Remote Support' -Style 'Heading1' -Before 220 -After 120 -Bold -Color '005F9E' -Size 16 -KeepNext))
        [void]$body.Add((New-OpenXmlParagraph -Text 'Registration' -Style 'Heading2' -Before 160 -After 100 -Bold -Color '005F9E' -Size 13 -KeepNext))
        [void]$body.Add((New-OpenXmlTable -Rows $registrationRows -Columns @('Item', 'Value') -Widths @(3000, 7800) -StatusColumns))
    }

    $securityRows = @(Get-ReportRecords -Data $Data -PropertyName 'SecurityDashboard')
    if ($securityRows.Count -gt 0) {
        [void]$body.Add((New-OpenXmlParagraph -Text 'Security Dashboard' -Style 'Heading1' -Before 220 -After 120 -Bold -Color '005F9E' -Size 16 -KeepNext))
        [void]$body.Add((New-OpenXmlTable -Rows $securityRows -StatusColumns))
    }

    $systemSections = @(
        @('Summary', 'ServerStatus', $true),
        @('Processors', 'Processors', $false),
        @('Memory', 'Memory', $false),
        @('Network', 'SystemNetwork', $false),
        @('Device Inventory', 'DeviceInventory', $false),
        @('Storage', 'Storage', $false)
    )
    $systemSectionRecords = @($systemSections | ForEach-Object {
        ,@($_[0], @(Get-ReportRecords -Data $Data -PropertyName $_[1] -PropertyMap:([bool]$_[2]) -OmitNAValues:($_[0] -in @('Memory', 'Network'))), [bool]$_[2])
    })
    if (@($systemSectionRecords | Where-Object { $_[1].Count -gt 0 }).Count -gt 0) {
        [void]$body.Add((New-OpenXmlParagraph -Text 'System Information' -Style 'Heading1' -Before 220 -After 120 -Bold -Color '005F9E' -Size 16 -KeepNext))
        foreach ($item in $systemSectionRecords) {
            $records = @($item[1])
            if ($records.Count -eq 0) { continue }
            [void]$body.Add((New-OpenXmlParagraph -Text $item[0] -Style 'Heading2' -Before 160 -After 100 -Bold -Color '005F9E' -Size 13 -KeepNext))
            $tableArguments = @{ Rows = $records; StatusColumns = $true }
            if ([bool]$item[2]) { $tableArguments.Columns = @('Item', 'Value'); $tableArguments.Widths = @(3000, 7800) }
            [void]$body.Add((New-OpenXmlTable @tableArguments))
        }
    }

    $firmwareRows = @(Get-ReportRecords -Data $Data -PropertyName 'Firmware')
    if ($firmwareRows.Count -gt 0) {
        [void]$body.Add((New-OpenXmlParagraph -Text 'Firmware & OS Software' -Style 'Heading1' -Before 220 -After 120 -Bold -Color '005F9E' -Size 16 -KeepNext))
        [void]$body.Add((New-OpenXmlTable -Rows $firmwareRows -StatusColumns))
    }

    $powerSections = @(
        @('Power Supply', 'PowerSupplies'),
        @('Fans', 'Fans'),
        @('Temperatures', 'Temperatures')
    )
    $powerSectionRecords = @($powerSections | ForEach-Object {
        ,@($_[0], @(Get-ReportRecords -Data $Data -PropertyName $_[1] -OmitNAValues:($_[0] -eq 'Temperatures')))
    })
    if (@($powerSectionRecords | Where-Object { $_[1].Count -gt 0 }).Count -gt 0) {
        [void]$body.Add((New-OpenXmlParagraph -Text 'Power & Thermal' -Style 'Heading1' -Before 220 -After 120 -Bold -Color '005F9E' -Size 16 -KeepNext))
        foreach ($item in $powerSectionRecords) {
            $records = @($item[1])
            if ($records.Count -eq 0) { continue }
            [void]$body.Add((New-OpenXmlParagraph -Text $item[0] -Style 'Heading2' -Before 160 -After 100 -Bold -Color '005F9E' -Size 13 -KeepNext))
            [void]$body.Add((New-OpenXmlTable -Rows $records -StatusColumns))
        }
    }

    $documentXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>$($body -join '')<w:sectPr><w:headerReference w:type="default" r:id="rId1"/><w:headerReference w:type="first" r:id="rId1"/><w:headerReference w:type="even" r:id="rId1"/><w:footerReference w:type="default" r:id="rId2"/><w:footerReference w:type="first" r:id="rId2"/><w:footerReference w:type="even" r:id="rId2"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="432" w:footer="432" w:gutter="0"/></w:sectPr></w:body></w:document>
"@
    $stylesXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/><w:sz w:val="20"/><w:szCs w:val="20"/><w:color w:val="222222"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr><w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:rPr><w:rFonts w:ascii="Aptos Display" w:hAnsi="Aptos Display"/><w:b/><w:color w:val="005F9E"/><w:sz w:val="56"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/><w:color w:val="203647"/><w:sz w:val="32"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="320" w:after="160"/></w:pPr><w:rPr><w:rFonts w:ascii="Aptos Display" w:hAnsi="Aptos Display"/><w:b/><w:color w:val="005F9E"/><w:sz w:val="32"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="280" w:after="140"/></w:pPr><w:rPr><w:rFonts w:ascii="Aptos Display" w:hAnsi="Aptos Display"/><w:b/><w:color w:val="005F9E"/><w:sz w:val="26"/></w:rPr></w:style></w:styles>
"@
    $numberingXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="540"/></w:tabs><w:ind w:left="540" w:hanging="260"/><w:spacing w:after="80" w:line="300" w:lineRule="auto"/></w:pPr><w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/><w:sz w:val="18"/></w:rPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>
"@
    $headerXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:p><w:pPr><w:tabs><w:tab w:val="right" w:pos="10800"/></w:tabs><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="4" w:color="005F9E"/></w:pBdr><w:spacing w:after="80"/></w:pPr>$(New-OpenXmlLogoRun)<w:r><w:tab/></w:r>$(New-OpenXmlRun -Text "HP iLO Check - $displayTarget - $CustomerName" -Size 8.5 -Color '506675')</w:p></w:hdr>
"@
    $footerXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:pPr><w:tabs><w:tab w:val="right" w:pos="10800"/></w:tabs><w:pBdr><w:top w:val="single" w:sz="6" w:space="4" w:color="005F9E"/></w:pBdr><w:spacing w:before="80"/></w:pPr>$(New-OpenXmlRun -Text "$([char]0x00A9)2026 Winslow Tech Group. All Right Reserved" -Size 8.5 -Color '506675')<w:r><w:tab/></w:r>$(New-OpenXmlRun -Text 'Page ' -Size 8.5 -Color '506675')<w:fldSimple w:instr="PAGE"><w:r><w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/><w:sz w:val="17"/><w:color w:val="506675"/></w:rPr><w:t>1</w:t></w:r></w:fldSimple>$(New-OpenXmlRun -Text ' of ' -Size 8.5 -Color '506675')<w:fldSimple w:instr="NUMPAGES"><w:r><w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/><w:sz w:val="17"/><w:color w:val="506675"/></w:rPr><w:t>1</w:t></w:r></w:fldSimple></w:p></w:ftr>
"@
    $contentTypes = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>
"@
    $rootRelationships = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>
"@
    $documentRelationships = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>
"@
    $headerRelationships = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/winslowtg-logo.png"/></Relationships>
"@
    $created = (Get-Date).ToUniversalTime().ToString('s') + 'Z'
    $coreXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>HP iLO Health Check - $(ConvertTo-OpenXmlText $displayTarget)</dc:title><dc:creator>HP iLO 5 Health Check Report</dc:creator><cp:lastModifiedBy>HP iLO 5 Health Check Report</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">$created</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">$created</dcterms:modified></cp:coreProperties>
"@
    $appXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>HP iLO 5 Health Check Report</Application><AppVersion>1.0</AppVersion></Properties>
"@

    Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
    $fileStream = [IO.File]::Open($resolved, [IO.FileMode]::Create, [IO.FileAccess]::ReadWrite)
    $archive = $null
    try {
        $archive = New-Object IO.Compression.ZipArchive($fileStream, [IO.Compression.ZipArchiveMode]::Create, $false)
        Add-OpenXmlPackageEntry $archive '[Content_Types].xml' $contentTypes
        Add-OpenXmlPackageEntry $archive '_rels/.rels' $rootRelationships
        Add-OpenXmlPackageEntry $archive 'word/document.xml' $documentXml
        Add-OpenXmlPackageEntry $archive 'word/styles.xml' $stylesXml
        Add-OpenXmlPackageEntry $archive 'word/numbering.xml' $numberingXml
        Add-OpenXmlPackageEntry $archive 'word/_rels/document.xml.rels' $documentRelationships
        Add-OpenXmlPackageEntry $archive 'word/header1.xml' $headerXml
        Add-OpenXmlPackageEntry $archive 'word/_rels/header1.xml.rels' $headerRelationships
        Add-OpenXmlBinaryPackageEntry $archive 'word/media/winslowtg-logo.png' $script:ReportLogoPath
        Add-OpenXmlPackageEntry $archive 'word/footer1.xml' $footerXml
        Add-OpenXmlPackageEntry $archive 'docProps/core.xml' $coreXml
        Add-OpenXmlPackageEntry $archive 'docProps/app.xml' $appXml
    }
    finally {
        if ($null -ne $archive) { $archive.Dispose() }
        else { $fileStream.Dispose() }
    }
    return $resolved
}

function New-WordHealthReport {
    param(
        [Parameter(Mandatory)][object]$Data,
        [Parameter(Mandatory)][string]$OutputPath,
        [Parameter(Mandatory)][string]$CustomerName
    )

    # The self-contained generator produces a stable, fully branded report
    # across Word versions. Set HP_ILO_USE_WORD_COM=1 only to opt into the
    # legacy Word automation path.
    if ($env:OS -ne 'Windows_NT' -or $env:HP_ILO_USE_WORD_COM -ne '1') {
        return New-OpenXmlHealthReport -Data $Data -OutputPath $OutputPath -CustomerName $CustomerName
    }
    $word = $null
    $document = $null
    try {
        try {
            $word = New-Object -ComObject Word.Application
        }
        catch {
            Write-Warning 'Microsoft Word is not installed. Creating the report with the built-in DOCX generator.'
            return New-OpenXmlHealthReport -Data $Data -OutputPath $OutputPath -CustomerName $CustomerName
        }
        $word.Visible = $false
        $word.DisplayAlerts = 0
        $document = $word.Documents.Add()
        $section = $document.Sections.Item(1)
        try {
            # WdPaperSize.wdPaperLetter = 2. Some printer drivers do not
            # advertise Letter; in that case, retain Word's current page size.
            $section.PageSetup.PaperSize = $script:WdPaperLetter
        }
        catch [Runtime.InteropServices.COMException] {
            Write-Verbose "Word cannot select Letter with the active printer; using Word's current page size."
        }
        $section.PageSetup.TopMargin = 36
        $section.PageSetup.BottomMargin = 36
        $section.PageSetup.LeftMargin = 36
        $section.PageSetup.RightMargin = 36
        $section.PageSetup.HeaderDistance = 21.6
        $section.PageSetup.FooterDistance = 21.6

        $normal = $document.Styles.Item('Normal')
        $normal.Font.Name = 'Aptos'
        $normal.Font.Size = 11
        $normal.Font.Color = ConvertTo-WordColor $script:ReportDark
        $normal.ParagraphFormat.SpaceAfter = 2

        $reportDate = Get-Date -Format 'MMMM d, yyyy'
        $displayTarget = ([uri]$Data.Target).Host
        Set-WordReportFurniture $section $displayTarget $CustomerName $reportDate
        Add-WordCover $document $displayTarget $CustomerName $reportDate

        Add-WordHeading $document 'Executive Overview' 1
        Add-WordParagraph $document "$CustomerName engaged Professional Services to conduct an HP iLO Health Check of $displayTarget. This report documents the discovery, analysis, and recommendations from the assessment." 11 $false '222222' 2

        $assessment = @(New-AssessmentSummary $Data)
        $recommendedAction = Get-RecommendedActionText -Data $Data -Assessment $assessment
        Add-WordHeading $document 'Recommended Action' 2
        Add-WordParagraph $document $recommendedAction 11 $false '222222' 6
        Add-WordHeading $document 'Assessment Summary' 2
        Add-AssessmentSummaryTable $document $assessment

        Add-WordHeading $document 'Information' 1
        foreach ($item in @(
            @('Server', 'ServerStatus', $true),
            @('iLO', 'IloInformation', $false),
            @('Status', 'StatusInformation', $false),
            @('HPE Compute Ops Management', 'ComputeOpsManagement', $false)
        )) {
        $records = @(Get-ReportRecords -Data $Data -PropertyName $item[1] -PropertyMap:([bool]$item[2]) -ExcludeColumns $(if ($item[0] -eq 'Status') { @('State') } elseif ($item[0] -eq 'HPE Compute Ops Management') { @('Next retry time') } else { @() }))
            if ($records.Count -eq 0) { continue }
            Add-WordHeading $document $item[0] 2
            Add-WordTable $document $records 'No records were returned.'
        }

        $remoteSupportRows = @(Get-ReportRecords -Data $Data -PropertyName 'RemoteSupportRegistration')
        if ($remoteSupportRows.Count -gt 0) {
            $registrationRows = @(Get-ReportPropertyRows $remoteSupportRows[0])
            Add-WordHeading $document 'Remote Support' 1
            Add-WordHeading $document 'Registration' 2
            Add-WordTable $document $registrationRows 'No Remote Support registration data was returned.'
        }

        $securityRows = @(Get-ReportRecords -Data $Data -PropertyName 'SecurityDashboard')
        if ($securityRows.Count -gt 0) {
            Add-WordHeading $document 'Security Dashboard' 1
            Add-WordTable $document $securityRows 'No Security Dashboard data was returned.'
        }

        $systemSections = @(
            @('Summary', 'ServerStatus', $true),
            @('Processors', 'Processors', $false),
            @('Memory', 'Memory', $false),
            @('Network', 'SystemNetwork', $false),
            @('Device Inventory', 'DeviceInventory', $false),
            @('Storage', 'Storage', $false)
        )
        $hasSystemInformation = $false
        foreach ($item in $systemSections) {
            if (@(Get-ReportRecords -Data $Data -PropertyName $item[1] -PropertyMap:([bool]$item[2])).Count -gt 0) {
                $hasSystemInformation = $true
                break
            }
        }
        if ($hasSystemInformation) {
            Add-WordHeading $document 'System Information' 1
            foreach ($item in $systemSections) {
                $records = @(Get-ReportRecords -Data $Data -PropertyName $item[1] -PropertyMap:([bool]$item[2]) -OmitNAValues:($item[0] -in @('Memory', 'Network')))
                if ($records.Count -eq 0) { continue }
                Add-WordHeading $document $item[0] 2
                Add-WordTable $document $records 'No records were returned.'
            }
        }

        $firmwareRows = @(Get-ReportRecords -Data $Data -PropertyName 'Firmware')
        if ($firmwareRows.Count -gt 0) {
            Add-WordHeading $document 'Firmware & OS Software' 1
            Add-WordTable $document $firmwareRows 'No firmware data was returned.'
        }

        $powerSections = @(
            @('Power Supply', 'PowerSupplies'),
            @('Fans', 'Fans'),
            @('Temperatures', 'Temperatures')
        )
        $hasPowerThermal = $false
        foreach ($item in $powerSections) {
            if (@(Get-ReportRecords -Data $Data -PropertyName $item[1]).Count -gt 0) {
                $hasPowerThermal = $true
                break
            }
        }
        if ($hasPowerThermal) {
            Add-WordHeading $document 'Power & Thermal' 1
            foreach ($item in $powerSections) {
                $records = @(Get-ReportRecords -Data $Data -PropertyName $item[1] -OmitNAValues:($item[0] -eq 'Temperatures'))
                if ($records.Count -eq 0) { continue }
                Add-WordHeading $document $item[0] 2
                Add-WordTable $document $records 'No records were returned.'
            }
        }

        $resolved = [IO.Path]::GetFullPath($OutputPath)
        $directory = Split-Path -Parent $resolved
        if (-not (Test-Path $directory)) { New-Item -ItemType Directory -Path $directory -Force | Out-Null }
        $document.Fields.Update() | Out-Null
        $section.Footers.Item(1).Range.Fields.Update() | Out-Null
        $document.SaveAs2($resolved, 16)
        return $resolved
    }
    finally {
        if ($null -ne $document) { $document.Close(0); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($document) }
        if ($null -ne $word) { $word.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Invoke-IloHealthReport {
    [CmdletBinding()]
    param(
        [string]$IloAddress,
        [PSCredential]$Credential,
        [string]$CustomerName,
        [string]$OutputPath,
        [int]$TimeoutSec = 30,
        [int]$MaxLogEntries = 100,
        [switch]$SkipCertificateCheck
    )

    if (-not $IloAddress) { $IloAddress = Read-Host 'Enter the iLO 5 IP address or FQDN' }
    if (-not $IloAddress) { throw 'An iLO IP address or FQDN is required.' }
    if (-not $CustomerName) { $CustomerName = Read-Host 'Enter customer name' }
    if (-not $CustomerName) { throw 'A customer name is required.' }
    if (-not $Credential) { $Credential = Get-Credential -Message "Credentials for $IloAddress" }
    if (-not $Credential) { throw 'Credentials are required.' }

    $normalized = $IloAddress.Trim().TrimEnd('/')
    if ($normalized -notmatch '^[a-zA-Z][a-zA-Z0-9+.-]*://') { $normalized = "https://$normalized" }
    $baseUri = [uri]$normalized
    if ($baseUri.Scheme -ne 'https') { throw 'Only HTTPS iLO endpoints are supported.' }

    if (-not $OutputPath) {
        $safeTarget = ($baseUri.Host -replace '[^A-Za-z0-9.-]', '-')
        $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
        $OutputPath = Join-Path $PSScriptRoot "ilo-health-$safeTarget-$stamp.docx"
    }

    $session = $null
    $previousCertificateCallback = $null
    $certificateCallbackChanged = $false
    $previousSecurityProtocol = $null
    $securityProtocolChanged = $false
    try {
        if ($PSVersionTable.PSVersion.Major -lt 7) {
            $previousSecurityProtocol = [Net.ServicePointManager]::SecurityProtocol
            [Net.ServicePointManager]::SecurityProtocol = $previousSecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
            $securityProtocolChanged = $true
            if ($SkipCertificateCheck) {
                $previousCertificateCallback = [Net.ServicePointManager]::ServerCertificateValidationCallback
                [Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
                $certificateCallbackChanged = $true
            }
        }
        Write-Host "Connecting to $($baseUri.AbsoluteUri.TrimEnd('/')) ..."
        $session = New-IloSession `
            -BaseUri $baseUri `
            -Credential $Credential `
            -TimeoutSec $TimeoutSec `
            -IgnoreCertificateErrors ([bool]$SkipCertificateCheck)
        $data = Get-IloHealthData $session $MaxLogEntries
        $reportPath = New-WordHealthReport -Data $data -OutputPath $OutputPath -CustomerName $CustomerName
        Write-Host "Word report created: $reportPath" -ForegroundColor Green
    }
    finally {
        if ($null -ne $session) { Remove-IloSession $session }
        if ($certificateCallbackChanged) {
            [Net.ServicePointManager]::ServerCertificateValidationCallback = $previousCertificateCallback
        }
        if ($securityProtocolChanged) {
            [Net.ServicePointManager]::SecurityProtocol = $previousSecurityProtocol
        }
    }
}

if ($MyInvocation.InvocationName -ne '.') {
    Invoke-IloHealthReport @PSBoundParameters
}
