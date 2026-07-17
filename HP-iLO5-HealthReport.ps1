<#+
.SYNOPSIS
Creates a Microsoft Word health report for an HPE iLO 5 server.

.DESCRIPTION
Uses Redfish session authentication, dynamically follows advertised resource
links, collects read-only health information, and creates a formatted .docx
report through Microsoft Word COM automation.
#>

[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$IloAddress,

    [PSCredential]$Credential,

    [string]$CustomerName = 'Customer',

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

function Get-RedfishLink {
    param(
        [AllowNull()][object]$Resource,
        [Parameter(Mandatory)][string]$Name
    )

    $value = Get-ObjectProperty -InputObject $Resource -Name $Name
    if ($null -eq $value) { return $null }
    return Get-ObjectProperty -InputObject $value -Name '@odata.id'
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
    $token = $response.Headers['X-Auth-Token']
    if (-not $token) { throw 'iLO did not return a Redfish session token.' }

    [PSCustomObject]@{
        BaseUri = $BaseUri
        Headers = @{ Accept = 'application/json'; 'OData-Version' = '4.0'; 'X-Auth-Token' = $token }
        SessionUri = [string]$response.Headers['Location']
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
        DisableKeepAlive = $true
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
    foreach ($uri in $candidates) {
        try {
            return @(Get-RedfishCollection -Session $Session -Uri $uri -Limit $Limit)
        }
        catch {
            $lastError = $_.Exception.Message
        }
    }

    Add-CollectionNote -Notes $Notes -Message "Unable to collect $Label`: $lastError"
    return @()
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

function Get-IloHealthData {
    param(
        [Parameter(Mandatory)][object]$Session,
        [Parameter(Mandatory)][int]$MaxLogEntries
    )

    $notes = [System.Collections.Generic.List[string]]::new()
    $root = Invoke-RedfishGet -Session $Session -Uri '/redfish/v1/'
    $systems = @(Get-SafeCollection $Session (Get-RedfishLink $root 'Systems') $notes 'systems')
    if ($systems.Count -eq 0) { throw 'No ComputerSystem resource was found.' }
    $system = $systems[0]
    $chassis = @(Get-SafeCollection $Session (Get-RedfishLink $root 'Chassis') $notes 'chassis')
    $managers = @(Get-SafeCollection $Session (Get-RedfishLink $root 'Managers') $notes 'managers')

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
    if ($text -match '(?i)critical|failed|fatal') { return ConvertTo-WordColor 'D00000' }
    if ($text -match '(?i)warning|degraded|caution|recommended') { return ConvertTo-WordColor 'E59B00' }
    if ($text -match '(?i)^ok$|^enabled$|^healthy$') { return ConvertTo-WordColor '008A3B' }
    return ConvertTo-WordColor '555555'
}

function Get-AssessmentStatus {
    param([AllowEmptyCollection()][object[]]$Items)

    if (@($Items).Count -eq 0) { return 'RECOMMENDED' }
    $signals = [System.Collections.Generic.List[string]]::new()
    foreach ($item in $Items) {
        foreach ($name in @('Health', 'HealthRollup', 'State', 'Severity')) {
            $value = if ($item -is [Collections.IDictionary] -and $item.Contains($name)) {
                $item[$name]
            }
            else {
                Get-ObjectProperty -InputObject $item -Name $name
            }
            if ($null -ne $value) { $signals.Add([string]$value) }
        }
    }
    if ($signals.Count -eq 0) { return 'RECOMMENDED' }
    $evidence = $signals -join ' '
    if ($evidence -match '(?i)critical|failed|fatal') { return 'CRITICAL' }
    if ($evidence -match '(?i)warning|degraded|caution|unknown|disabled') { return 'RECOMMENDED' }
    return 'HEALTHY'
}

function New-AssessmentSummary {
    param([Parameter(Mandatory)][object]$Data)

    $activeEvents = @($Data.EventLogs | Where-Object {
        $_.Repaired -notin @($true, 'True', 'true')
    })
    $securityEvents = @($activeEvents | Where-Object {
        $_.Log -match '(?i)security' -or $_.Message -match '(?i)security|unauthorized|authentication'
    })
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
        [PSCustomObject]@{ Section = 'Administration'; Status = Get-AssessmentStatus $activeEvents }
        [PSCustomObject]@{ Section = 'Security'; Status = Get-AssessmentStatus $securityEvents }
        [PSCustomObject]@{ Section = 'Management'; Status = Get-AssessmentStatus @($Data.Management) }
        [PSCustomObject]@{ Section = 'Lifecycle Management'; Status = Get-AssessmentStatus @($Data.Firmware) }
    )
}

function Get-OverallHealthScore {
    param([Parameter(Mandatory)][object[]]$Assessment)
    $critical = @($Assessment | Where-Object Status -eq 'CRITICAL').Count
    $recommended = @($Assessment | Where-Object Status -eq 'RECOMMENDED').Count
    return [Math]::Max(0, 100 - ($critical * 15) - ($recommended * 5))
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
            $cell.Range.Font.Bold = if ($name -in @('Health', 'Severity')) { -1 } else { 0 }
            $cell.Range.Font.Color = if ($name -in @('Health', 'Severity', 'State')) { Get-StatusColor $value } else { ConvertTo-WordColor '222222' }
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

function Add-OverallHealthScore {
    param(
        [Parameter(Mandatory)][object]$Document,
        [Parameter(Mandatory)][object[]]$Assessment
    )

    $score = Get-OverallHealthScore $Assessment
    $critical = @($Assessment | Where-Object Status -eq 'CRITICAL').Count
    $recommended = @($Assessment | Where-Object Status -eq 'RECOMMENDED').Count
    $action = if ($critical -gt 0) { 'Action Required' } elseif ($recommended -gt 0) { 'Review Recommended' } else { 'Healthy' }

    $range = Get-EndRange $Document
    $table = $Document.Tables.Add($range, 1, 2)
    $table.Style = 'Table Grid'
    $table.AllowAutoFit = $false
    $table.PreferredWidthType = $script:WdPreferredWidthPoints
    $table.PreferredWidth = 540
    $table.Columns.Item(1).Width = 180
    $table.Columns.Item(2).Width = 360
    $scoreCell = $table.Cell(1, 1)
    $scoreCell.Range.Text = "$score`r/ 100"
    $scoreCell.Range.Font.Name = 'Aptos Display'
    $scoreCell.Range.Font.Size = 18
    $scoreCell.Range.Font.Bold = -1
    $scoreCell.Range.Font.Color = ConvertTo-WordColor 'FFFFFF'
    $scoreCell.Range.ParagraphFormat.Alignment = 1
    $scoreCell.Shading.BackgroundPatternColor = ConvertTo-WordColor $script:ReportBlue
    $scoreCell.VerticalAlignment = 1

    $detailCell = $table.Cell(1, 2)
    $detailCell.Range.Text = "$action`rCritical sections: $critical    Recommended reviews: $recommended`rScore is calculated from section status evidence returned by Redfish."
    $detailCell.Range.Font.Name = 'Aptos'
    $detailCell.Range.Font.Size = 10
    $detailCell.Range.Paragraphs.Item(1).Range.Font.Size = 13
    $detailCell.Range.Paragraphs.Item(1).Range.Font.Bold = -1
    $detailCell.Range.Paragraphs.Item(1).Range.Font.Color = Get-StatusColor $(if ($critical -gt 0) { 'CRITICAL' } elseif ($recommended -gt 0) { 'RECOMMENDED' } else { 'HEALTHY' })
    $detailCell.Range.Paragraphs.Item(3).Range.Font.Italic = -1
    $detailCell.Range.Paragraphs.Item(3).Range.Font.Color = ConvertTo-WordColor $script:ReportDark
    $detailCell.VerticalAlignment = 1
    $table.TopPadding = 6
    $table.BottomPadding = 6
    $table.LeftPadding = 8
    $table.RightPadding = 8
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

function New-WordHealthReport {
    param(
        [Parameter(Mandatory)][object]$Data,
        [Parameter(Mandatory)][string]$OutputPath,
        [Parameter(Mandatory)][string]$CustomerName
    )

    if ($env:OS -ne 'Windows_NT') { throw 'Microsoft Word report generation requires Windows.' }
    $word = $null
    $document = $null
    try {
        $word = New-Object -ComObject Word.Application
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
        Add-WordHeading $document 'Assessment Summary' 2
        Add-AssessmentSummaryTable $document $assessment
        Add-WordHeading $document 'Overall Health Score' 2
        Add-OverallHealthScore $document $assessment

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
            @('Administration - Event Logs', 'EventLogs')
        )
        foreach ($item in $sections) {
            $records = @($Data.PSObject.Properties[$item[1]].Value | Where-Object { Test-ReportRecordPresent $_ })
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
        [string]$CustomerName = 'Customer',
        [string]$OutputPath,
        [int]$TimeoutSec = 30,
        [int]$MaxLogEntries = 100,
        [switch]$SkipCertificateCheck
    )

    if (-not $IloAddress) { $IloAddress = Read-Host 'Enter the iLO 5 IP address or FQDN' }
    if (-not $IloAddress) { throw 'An iLO IP address or FQDN is required.' }
    if (-not $Credential) { $Credential = Get-Credential -Message "Credentials for $IloAddress" }
    if (-not $Credential) { throw 'Credentials are required.' }

    $normalized = $IloAddress.Trim().TrimEnd('/')
    if ($normalized -notmatch '^[a-zA-Z][a-zA-Z0-9+.-]*://') { $normalized = "https://$normalized" }
    $baseUri = [uri]$normalized
    if ($baseUri.Scheme -ne 'https') { throw 'Only HTTPS iLO endpoints are supported.' }

    if (-not $OutputPath) {
        $safeTarget = ($baseUri.Host -replace '[^A-Za-z0-9.-]', '-')
        $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
        $OutputPath = Join-Path 'reports' "ilo-health-$safeTarget-$stamp.docx"
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
