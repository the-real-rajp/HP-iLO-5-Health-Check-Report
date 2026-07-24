<#+
.SYNOPSIS
Creates a Microsoft Word health report for an HPE iLO 5 server.

.DESCRIPTION
Uses Redfish session authentication, dynamically follows advertised resource
links, collects read-only health information, and creates a formatted .docx
report. Microsoft Word COM automation is used when available; otherwise the
script creates the DOCX directly with its built-in Open XML generator.
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

    [ordered]@{
        'Name' = (Get-ObjectProperty $System 'HostName' (Get-ObjectProperty $System 'Name' 'Unknown'))
        'Model' = Get-ObjectProperty $System 'Model' 'Unknown'
        'Manufacturer' = Get-ObjectProperty $System 'Manufacturer' 'Unknown'
        'Serial number' = Get-ObjectProperty $System 'SerialNumber' 'Unknown'
        'Power state' = Get-ObjectProperty $System 'PowerState' 'Unknown'
        'BIOS version' = Get-ObjectProperty $System 'BiosVersion' 'Unknown'
        'Health' = Get-HealthValue $System
    }
}

function Convert-Temperature {
    param([Parameter(Mandatory)][object]$Item)
    [PSCustomObject][ordered]@{
        'Name' = Get-ObjectProperty $Item 'Name' 'Unknown'
        'Reading (C)' = Get-ObjectProperty $Item 'ReadingCelsius' 'N/A'
        'Upper critical (C)' = Get-ObjectProperty $Item 'UpperThresholdCritical' 'N/A'
        'Health' = Get-HealthValue $Item
        'State' = Get-StateValue $Item
    }
}

function Convert-Fan {
    param([Parameter(Mandatory)][object]$Item)
    [PSCustomObject][ordered]@{
        'Name' = Get-ObjectProperty $Item 'Name' 'Unknown'
        'Reading' = Get-ObjectProperty $Item 'Reading' 'N/A'
        'Units' = Get-ObjectProperty $Item 'ReadingUnits' 'N/A'
        'Health' = Get-HealthValue $Item
        'State' = Get-StateValue $Item
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
        'State' = Get-StateValue $Item
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
        'State' = Get-StateValue $Item
    }
}

function Test-ReportRecordPresent {
    param([AllowNull()][object]$Record)

    if ($null -eq $Record) { return $false }
    $state = Get-ObjectProperty -InputObject $Record -Name 'State' -Default ''
    return ([string]$state -notmatch '(?i)^absent$')
}

function Convert-Processor {
    param([Parameter(Mandatory)][object]$Item)
    [PSCustomObject][ordered]@{
        'Name' = Get-ObjectProperty $Item 'Socket' (Get-ObjectProperty $Item 'Name' 'Unknown')
        'Model' = Get-ObjectProperty $Item 'Model' 'Unknown'
        'Cores' = Get-ObjectProperty $Item 'TotalCores' 'N/A'
        'Threads' = Get-ObjectProperty $Item 'TotalThreads' 'N/A'
        'Health' = Get-HealthValue $Item
        'State' = Get-StateValue $Item
    }
}

function Convert-Firmware {
    param([Parameter(Mandatory)][object]$Item)
    [PSCustomObject][ordered]@{
        'Name' = Get-ObjectProperty $Item 'Name' (Get-ObjectProperty $Item 'Id' 'Unknown')
        'Version' = Get-ObjectProperty $Item 'Version' 'Unknown'
        'Updateable' = Get-ObjectProperty $Item 'Updateable' 'Unknown'
        'Health' = Get-HealthValue $Item
        'State' = Get-StateValue $Item
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
        'Recommended action' = ''
        'Ignored' = $false
    }
}

function Convert-SecurityParameter {
    param([Parameter(Mandatory)][object]$Item)

    [PSCustomObject][ordered]@{
        'Name' = Get-ObjectProperty $Item 'Name' (Get-ObjectProperty $Item 'Id' 'Security parameter')
        'Security status' = Get-ObjectProperty $Item 'SecurityStatus' 'Unknown'
        'Current value' = Get-ObjectProperty $Item 'State' 'Unknown'
        'Recommended action' = Get-ObjectProperty $Item 'RecommendedAction' ''
        'Ignored' = [bool](Get-ObjectProperty $Item 'Ignore' $false)
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

    $storage = [System.Collections.Generic.List[object]]::new()
    $systemUri = Get-ObjectProperty $system '@odata.id'
    $standardStorageUri = if ($systemUri) { "$($systemUri.TrimEnd('/'))/Storage" } else { $null }
    $storageUris = @($standardStorageUri, (Get-RedfishLink $system 'Storage'))
    foreach ($item in @(Get-SafeCollectionFromUris $Session $storageUris $notes 'storage')) {
        $storage.Add([PSCustomObject][ordered]@{
            'Name' = Get-ObjectProperty $item 'Name' (Get-ObjectProperty $item 'Id' 'Unknown')
            'Description' = Get-ObjectProperty $item 'Description' ''
            'Health' = Get-HealthValue $item
            'State' = Get-StateValue $item
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
                    'State' = Get-StateValue $child
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

    $management = @($managers | ForEach-Object {
        [PSCustomObject][ordered]@{
            'Name' = Get-ObjectProperty $_ 'Name' (Get-ObjectProperty $_ 'Id' 'iLO Manager')
            'Firmware version' = Get-ObjectProperty $_ 'FirmwareVersion' 'Unknown'
            'Health' = Get-HealthValue $_
            'State' = Get-StateValue $_
        }
    })

    [PSCustomObject][ordered]@{
        GeneratedAt = [DateTimeOffset]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ')
        Target = $Session.BaseUri.AbsoluteUri.TrimEnd('/')
        ServerStatus = Convert-ServerStatus $system
        Temperatures = $temperatures
        Fans = $fans
        PowerSupplies = $powerSupplies
        Storage = $storage.ToArray()
        Memory = $memory
        Processors = $processors
        Firmware = $firmware
        Management = $management
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
    if ($text -match '(?i)^ok$|^enabled$|^healthy$') { return ConvertTo-WordColor '008A3B' }
    return ConvertTo-WordColor '555555'
}

function Get-AssessmentStatus {
    param([AllowEmptyCollection()][object[]]$Items)

    if (@($Items).Count -eq 0) { return 'RECOMMENDED' }
    $signals = [System.Collections.Generic.List[string]]::new()
    foreach ($item in $Items) {
        $ignored = [bool](Get-ObjectProperty -InputObject $item -Name 'Ignored' -Default $false)
        foreach ($name in @('Health', 'HealthRollup', 'State', 'Severity', 'SecurityStatus', 'OverallSecurityStatus', 'Security status')) {
            $value = if ($item -is [Collections.IDictionary] -and $item.Contains($name)) {
                $item[$name]
            }
            else {
                Get-ObjectProperty -InputObject $item -Name $name
            }
            if ($null -ne $value) {
                if ($ignored -and $name -match '(?i)security' -and [string]$value -match '(?i)^risk$') {
                    $signals.Add('Ignored')
                }
                else {
                    $signals.Add([string]$value)
                }
            }
        }
    }
    if ($signals.Count -eq 0) { return 'RECOMMENDED' }
    $evidence = $signals -join ' '
    if ($evidence -match '(?i)critical|failed|fatal|risk') { return 'CRITICAL' }
    if ($evidence -match '(?i)warning|degraded|caution|unknown|disabled|ignored') { return 'RECOMMENDED' }
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

function New-AssessmentSummary {
    param([Parameter(Mandatory)][object]$Data)

    $criticalRecentEvents = @(Get-CriticalRecentEventLogs $Data)
    $securityEvents = @($criticalRecentEvents | Where-Object {
        $_.Log -match '(?i)security' -or $_.Message -match '(?i)security|unauthorized|authentication'
    })
    $securityEvidence = @((Get-ObjectProperty $Data 'SecurityDashboard' @())) + $securityEvents
    $powerThermal = @($Data.Temperatures) + @($Data.Fans) + @($Data.PowerSupplies)
    $performance = @($Data.Memory) + @($Data.Processors)

    return @(
        [PSCustomObject]@{ Section = 'Information'; Status = Get-AssessmentStatus @($Data.ServerStatus) }
        [PSCustomObject]@{ Section = 'System Information'; Status = Get-AssessmentStatus @($Data.ServerStatus) }
        [PSCustomObject]@{ Section = 'Firmware & OS Software'; Status = Get-AssessmentStatus @($Data.Firmware) }
        [PSCustomObject]@{ Section = 'iLO Federation'; Status = 'RECOMMENDED' }
        [PSCustomObject]@{ Section = 'Remote Console & Media'; Status = 'RECOMMENDED' }
        [PSCustomObject]@{ Section = 'Power & Thermal'; Status = Get-AssessmentStatus $powerThermal }
        [PSCustomObject]@{ Section = 'Performance'; Status = Get-AssessmentStatus $performance }
        [PSCustomObject]@{ Section = 'iLO Dedicated Network Port'; Status = 'RECOMMENDED' }
        [PSCustomObject]@{ Section = 'iLO Shared Network Port'; Status = 'RECOMMENDED' }
        [PSCustomObject]@{ Section = 'Remote Support'; Status = 'RECOMMENDED' }
        [PSCustomObject]@{ Section = 'Administration'; Status = Get-AssessmentStatus $criticalRecentEvents }
        [PSCustomObject]@{ Section = 'Security'; Status = Get-AssessmentStatus $securityEvidence }
        [PSCustomObject]@{ Section = 'Management'; Status = Get-AssessmentStatus @($Data.Management) }
        [PSCustomObject]@{ Section = 'Lifecycle Management'; Status = Get-AssessmentStatus @($Data.Firmware) }
    )
}

function Get-RecommendedActionText {
    param(
        [Parameter(Mandatory)][object]$Data,
        [Parameter(Mandatory)][object[]]$Assessment
    )

    $actions = [Collections.Generic.List[string]]::new()
    $securityDashboard = @(Get-ObjectProperty -InputObject $Data -Name 'SecurityDashboard' -Default @())
    foreach ($finding in $securityDashboard) {
        $ignored = [bool](Get-ObjectProperty -InputObject $finding -Name 'Ignored' -Default $false)
        $status = [string](Get-ObjectProperty -InputObject $finding -Name 'Security status' -Default '')
        if ($ignored -or $status -notmatch '(?i)risk|critical|warning') { continue }
        $recommendation = [string](Get-ObjectProperty -InputObject $finding -Name 'Recommended action' -Default '')
        if ($recommendation -and -not $actions.Contains($recommendation)) {
            $actions.Add($recommendation)
        }
    }
    if (@(Get-CriticalRecentEventLogs $Data).Count -gt 0) {
        $actions.Add('Investigate and resolve the critical iLO event-log entries recorded during the previous month.')
    }
    $otherCritical = @($Assessment | Where-Object {
        $_.Status -eq 'CRITICAL' -and $_.Section -notin @('Administration', 'Security')
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
            $statusColumns = @('Health', 'Severity', 'State', 'Security status', 'SecurityStatus')
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
    $table = $Document.Tables.Add($range, 8, 4)
    $table.Style = 'Table Grid'
    $table.AllowAutoFit = $false
    $table.PreferredWidthType = $script:WdPreferredWidthPoints
    $table.PreferredWidth = 540
    $widths = @(175, 95, 175, 95)
    for ($column = 1; $column -le 4; $column++) {
        $table.Columns.Item($column).Width = $widths[$column - 1]
        $header = $table.Cell(1, $column)
        $header.Range.Text = if ($column % 2 -eq 1) { 'Section' } else { 'Status' }
        $header.Range.Font.Name = 'Aptos'
        $header.Range.Font.Size = 9
        $header.Range.Font.Bold = -1
        $header.Range.Font.Color = ConvertTo-WordColor 'FFFFFF'
        $header.Shading.BackgroundPatternColor = ConvertTo-WordColor $script:ReportBlue
        $header.VerticalAlignment = 1
    }
    $table.Rows.Item(1).HeadingFormat = -1

    for ($row = 0; $row -lt 7; $row++) {
        $left = $Assessment[$row * 2]
        $right = $Assessment[($row * 2) + 1]
        $values = @($left.Section, $left.Status, $right.Section, $right.Status)
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

    $header = $Section.Headers.Item(1).Range
    $header.Text = "HP iLO Check - $Target - $CustomerName`t$ReportDate"
    $header.Font.Name = 'Aptos'
    $header.Font.Size = 9
    $header.Font.Color = ConvertTo-WordColor $script:ReportDark
    $header.ParagraphFormat.TabStops.Add(540, 2, 0) | Out-Null
    $headerBorder = $header.Paragraphs.Item(1).Borders.Item($script:WdBorderBottom)
    $headerBorder.LineStyle = 1
    $headerBorder.LineWidth = 4
    $headerBorder.Color = ConvertTo-WordColor $script:ReportBlue

    $footer = $Section.Footers.Item(1).Range
    $footer.Text = "Confidential`tPage "
    $footer.Font.Name = 'Aptos'
    $footer.Font.Size = 9
    $footer.Font.Color = ConvertTo-WordColor $script:ReportDark
    $footer.ParagraphFormat.TabStops.Add(540, 2, 0) | Out-Null
    $footer.Paragraphs.Item(1).Range.Words.Item(1).Font.Italic = -1
    $footer.Paragraphs.Item(1).Range.Words.Item(1).Font.Color = ConvertTo-WordColor '888888'
    $pageRange = $Section.Footers.Item(1).Range.Duplicate
    $pageRange.SetRange($pageRange.End - 1, $pageRange.End - 1)
    $pageRange.Fields.Add($pageRange, $script:WdFieldPage) | Out-Null
    $suffixRange = $Section.Footers.Item(1).Range.Duplicate
    $suffixRange.SetRange($suffixRange.End - 1, $suffixRange.End - 1)
    $suffixRange.InsertAfter(' of ')
    $totalRange = $Section.Footers.Item(1).Range.Duplicate
    $totalRange.SetRange($totalRange.End - 1, $totalRange.End - 1)
    $totalRange.Fields.Add($totalRange, $script:WdFieldNumPages) | Out-Null
    $footerBorder = $Section.Footers.Item(1).Range.Paragraphs.Item(1).Borders.Item($script:WdBorderTop)
    $footerBorder.LineStyle = 1
    $footerBorder.LineWidth = 4
    $footerBorder.Color = ConvertTo-WordColor $script:ReportBlue
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
        '^(HEALTHY|OK|Ok|Enabled)$' { return '00843D' }
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
        [int]$TotalWidth = 9360
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
<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/><w:tblInd w:w="120" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="80" w:type="dxa"/><w:start w:w="120" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:end w:w="120" w:type="dxa"/></w:tblCellMar><w:tblBorders><w:top w:val="single" w:sz="4" w:color="C8D5DF"/><w:left w:val="single" w:sz="4" w:color="C8D5DF"/><w:bottom w:val="single" w:sz="4" w:color="C8D5DF"/><w:right w:val="single" w:sz="4" w:color="C8D5DF"/><w:insideH w:val="single" w:sz="4" w:color="DCE5EB"/><w:insideV w:val="single" w:sz="4" w:color="DCE5EB"/></w:tblBorders></w:tblPr><w:tblGrid>$grid</w:tblGrid>$($tableRows -join '')</w:tbl>
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
    $assessmentRows = for ($index = 0; $index -lt 7; $index++) {
        [PSCustomObject][ordered]@{
            Assessment = $assessment[$index * 2].Section
            Status = $assessment[$index * 2].Status
            'Assessment ' = $assessment[($index * 2) + 1].Section
            'Status ' = $assessment[($index * 2) + 1].Status
        }
    }
    [void]$body.Add((New-OpenXmlTable -Rows $assessmentRows -Columns @('Assessment', 'Status', 'Assessment ', 'Status ') -Widths @(3000, 1680, 3000, 1680) -StatusColumns))

    [void]$body.Add((New-OpenXmlParagraph -Text 'Information' -Style 'Heading2' -Before 140 -After 100 -Bold -Color '005F9E' -Size 13 -KeepNext))
    $summaryRows = @($Data.ServerStatus.GetEnumerator() | ForEach-Object {
        [PSCustomObject][ordered]@{ Item = $_.Key; Value = $_.Value }
    })
    if ($summaryRows.Count -gt 0) {
        [void]$body.Add((New-OpenXmlTable -Rows $summaryRows -Columns @('Item', 'Value') -Widths @(2700, 6660) -StatusColumns))
    }

    $sections = @(
        @('Power & Thermal - Temperatures', 'Temperatures'),
        @('Power & Thermal - Fans', 'Fans'),
        @('Power & Thermal - Power Supplies', 'PowerSupplies'),
        @('Storage Controllers, Drives & Volumes', 'Storage'),
        @('Memory', 'Memory'),
        @('Processors', 'Processors'),
        @('Firmware & OS Software', 'Firmware'),
        @('Management', 'Management'),
        @('Security Dashboard', 'SecurityDashboard'),
        @('Administration - Event Logs', 'EventLogs')
    )
    foreach ($item in $sections) {
        $property = $Data.PSObject.Properties[$item[1]]
        if ($null -eq $property) { continue }
        if ($item[1] -eq 'EventLogs') {
            $records = @(Get-CriticalRecentEventLogs $Data)
        }
        else {
            $records = @($property.Value | Where-Object { Test-ReportRecordPresent $_ })
        }
        if ($records.Count -eq 0) { continue }
        [void]$body.Add((New-OpenXmlParagraph -Text $item[0] -Style 'Heading2' -Before 160 -After 100 -Bold -Color '005F9E' -Size 13 -KeepNext))
        [void]$body.Add((New-OpenXmlTable -Rows $records -StatusColumns))
    }

    if (@($Data.CollectionNotes).Count -gt 0) {
        [void]$body.Add((New-OpenXmlParagraph -Text 'Collection Notes' -Style 'Heading2' -Before 160 -After 100 -Bold -Color '005F9E' -Size 13 -KeepNext))
        foreach ($note in $Data.CollectionNotes) {
            [void]$body.Add((New-OpenXmlBulletParagraph -Text $note))
        }
    }

    $documentXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>$($body -join '')<w:sectPr><w:headerReference w:type="default" r:id="rId1"/><w:footerReference w:type="default" r:id="rId2"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>
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
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:pPr><w:tabs><w:tab w:val="right" w:pos="9360"/></w:tabs><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="4" w:color="005F9E"/></w:pBdr><w:spacing w:after="80"/></w:pPr>$(New-OpenXmlRun -Text "HP iLO Check - $displayTarget - $CustomerName" -Size 8.5 -Color '506675')$(New-OpenXmlRun -Text "`t$reportDate" -Size 8.5 -Color '506675')</w:p></w:hdr>
"@
    $footerXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:pPr><w:tabs><w:tab w:val="right" w:pos="9360"/></w:tabs><w:pBdr><w:top w:val="single" w:sz="6" w:space="4" w:color="005F9E"/></w:pBdr><w:spacing w:before="80"/></w:pPr>$(New-OpenXmlRun -Text 'Confidential' -Size 8.5 -Color '888888' -Italic)$(New-OpenXmlRun -Text "`tPage " -Size 8.5 -Color '506675')<w:fldSimple w:instr="PAGE"><w:r><w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/><w:sz w:val="17"/><w:color w:val="506675"/></w:rPr><w:t>1</w:t></w:r></w:fldSimple>$(New-OpenXmlRun -Text ' of ' -Size 8.5 -Color '506675')<w:fldSimple w:instr="NUMPAGES"><w:r><w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/><w:sz w:val="17"/><w:color w:val="506675"/></w:rPr><w:t>1</w:t></w:r></w:fldSimple></w:p></w:ftr>
"@
    $contentTypes = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>
"@
    $rootRelationships = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>
"@
    $documentRelationships = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>
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

    if ($env:OS -ne 'Windows_NT') {
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

        Add-WordHeading $document 'Information' 2
        $summaryRows = @($Data.ServerStatus.GetEnumerator() | ForEach-Object {
            [PSCustomObject][ordered]@{ Item = $_.Key; Value = $_.Value }
        })
        Add-WordTable $document $summaryRows 'Server status was not available.'

        $sections = @(
            @('Power & Thermal - Temperatures', 'Temperatures'),
            @('Power & Thermal - Fans', 'Fans'),
            @('Power & Thermal - Power Supplies', 'PowerSupplies'),
            @('Storage Controllers, Drives & Volumes', 'Storage'),
            @('Memory', 'Memory'),
            @('Processors', 'Processors'),
            @('Firmware & OS Software', 'Firmware'),
            @('Management', 'Management'),
            @('Security Dashboard', 'SecurityDashboard'),
            @('Administration - Event Logs', 'EventLogs')
        )
        foreach ($item in $sections) {
            if ($item[1] -eq 'EventLogs') {
                $records = @(Get-CriticalRecentEventLogs $Data)
            }
            else {
                $records = @($Data.PSObject.Properties[$item[1]].Value | Where-Object { Test-ReportRecordPresent $_ })
            }
            if ($records.Count -eq 0) { continue }
            Add-WordHeading $document $item[0] 2
            Add-WordTable $document $records 'No records were returned.'
        }

        if (@($Data.CollectionNotes).Count -gt 0) {
            Add-WordHeading $document 'Collection Notes' 2
            foreach ($note in $Data.CollectionNotes) {
                Add-WordParagraph $document "- $note" 9.5 $false '666666' 4
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
