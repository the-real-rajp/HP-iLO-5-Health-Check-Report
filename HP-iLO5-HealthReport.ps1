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

    [string]$OutputPath,

    [ValidateRange(1, 3600)]
    [int]$TimeoutSec = 30,

    [ValidateRange(0, 10000)]
    [int]$MaxLogEntries = 100,

    [switch]$SkipCertificateCheck
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

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
        Write-Warning "Unable to close the Redfish session: $($_.Exception.Message)"
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
    Invoke-RestMethod @request
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
    $Notes.Add($Message)
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
            $temperatures = @(@(Get-ObjectProperty $thermal 'Temperatures' @()) | ForEach-Object { Convert-Temperature $_ })
            $fans = @(@(Get-ObjectProperty $thermal 'Fans' @()) | ForEach-Object { Convert-Fan $_ })
        }
        catch { Add-CollectionNote $notes "Unable to collect thermal data: $($_.Exception.Message)" }
        try {
            $power = Invoke-RedfishGet $Session $powerUri
            $powerSupplies = @(@(Get-ObjectProperty $power 'PowerSupplies' @()) | ForEach-Object { Convert-PowerSupply $_ })
        }
        catch { Add-CollectionNote $notes "Unable to collect power data: $($_.Exception.Message)" }
    }

    $memory = @((Get-SafeCollection $Session (Get-RedfishLink $system 'Memory') $notes 'memory') | ForEach-Object { Convert-Memory $_ })
    $processors = @((Get-SafeCollection $Session (Get-RedfishLink $system 'Processors') $notes 'processors') | ForEach-Object { Convert-Processor $_ })

    $storage = [System.Collections.Generic.List[object]]::new()
    foreach ($item in @(Get-SafeCollection $Session (Get-RedfishLink $system 'Storage') $notes 'storage')) {
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
    foreach ($owner in @($system) + $managers) {
        foreach ($service in @(Get-SafeCollection $Session (Get-RedfishLink $owner 'LogServices') $notes 'log services')) {
            $logName = Get-ObjectProperty $service 'Name' (Get-ObjectProperty $service 'Id' 'Event log')
            foreach ($entry in @(Get-SafeCollection $Session (Get-RedfishLink $service 'Entries') $notes "$logName entries" $MaxLogEntries)) {
                $eventLogs.Add((Convert-LogEntry $entry $logName))
            }
        }
    }

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
    if ($text -match '(?i)critical|failed|fatal') { return ConvertTo-WordColor '9B1C1C' }
    if ($text -match '(?i)warning|degraded|caution') { return ConvertTo-WordColor '7A5A00' }
    if ($text -match '(?i)^ok$|^enabled$') { return ConvertTo-WordColor '2E6B3A' }
    return ConvertTo-WordColor '555555'
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
        [double]$SpaceAfter = 6
    )
    $range = Get-EndRange $Document
    $range.Text = "$Text`r"
    $paragraph = $range.Paragraphs.Item(1)
    $paragraph.Range.Font.Name = 'Calibri'
    $paragraph.Range.Font.Size = $Size
    $paragraph.Range.Font.Bold = [int]$Bold
    $paragraph.Range.Font.Color = ConvertTo-WordColor $Color
    $paragraph.Format.SpaceAfter = $SpaceAfter
    $paragraph.Format.LineSpacingRule = 0
}

function Add-WordHeading {
    param(
        [Parameter(Mandatory)][object]$Document,
        [Parameter(Mandatory)][string]$Text
    )
    $range = Get-EndRange $Document
    $range.Text = "$Text`r"
    $paragraph = $range.Paragraphs.Item(1)
    $paragraph.Range.Style = 'Heading 1'
    $paragraph.Range.Font.Name = 'Calibri'
    $paragraph.Range.Font.Size = 16
    $paragraph.Range.Font.Bold = -1
    $paragraph.Range.Font.Color = ConvertTo-WordColor '2E74B5'
    $paragraph.Format.SpaceBefore = 18
    $paragraph.Format.SpaceAfter = 10
    $paragraph.Format.KeepWithNext = -1
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
    $table.PreferredWidthType = 2
    $table.PreferredWidth = 468
    $table.Rows.Item(1).HeadingFormat = -1

    $scores = foreach ($name in $properties) {
        $lengths = @($name.Length) + @($Records | ForEach-Object { ([string]$_.PSObject.Properties[$name].Value).Length })
        [Math]::Max(10, [Math]::Min(36, ($lengths | Measure-Object -Maximum).Maximum))
    }
    $totalScore = ($scores | Measure-Object -Sum).Sum
    for ($column = 1; $column -le $properties.Count; $column++) {
        $table.Columns.Item($column).Width = [Math]::Max(45, 468 * $scores[$column - 1] / $totalScore)
    }

    for ($column = 1; $column -le $properties.Count; $column++) {
        $cell = $table.Cell(1, $column)
        $cell.Range.Text = $properties[$column - 1]
        $cell.Range.Font.Name = 'Calibri'
        $cell.Range.Font.Size = 9
        $cell.Range.Font.Bold = -1
        $cell.Range.Font.Color = ConvertTo-WordColor 'FFFFFF'
        $cell.Shading.BackgroundPatternColor = ConvertTo-WordColor '1F4E78'
        $cell.VerticalAlignment = 1
    }

    for ($row = 1; $row -le $Records.Count; $row++) {
        $table.Rows.Item($row + 1).AllowBreakAcrossPages = 0
        for ($column = 1; $column -le $properties.Count; $column++) {
            $name = $properties[$column - 1]
            $value = $Records[$row - 1].PSObject.Properties[$name].Value
            $cell = $table.Cell($row + 1, $column)
            $cell.Range.Text = [string]$value
            $cell.Range.Font.Name = 'Calibri'
            $cell.Range.Font.Size = 8.5
            $cell.Range.Font.Bold = if ($name -in @('Health', 'Severity')) { -1 } else { 0 }
            $cell.Range.Font.Color = if ($name -in @('Health', 'Severity', 'State')) { Get-StatusColor $value } else { ConvertTo-WordColor '222222' }
            if ($row % 2 -eq 0) { $cell.Shading.BackgroundPatternColor = ConvertTo-WordColor 'F2F4F7' }
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

function New-WordHealthReport {
    param(
        [Parameter(Mandatory)][object]$Data,
        [Parameter(Mandatory)][string]$OutputPath
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
        $section.PageSetup.PaperSize = 0
        $section.PageSetup.TopMargin = 72
        $section.PageSetup.BottomMargin = 72
        $section.PageSetup.LeftMargin = 72
        $section.PageSetup.RightMargin = 72
        $section.PageSetup.HeaderDistance = 35.4
        $section.PageSetup.FooterDistance = 35.4

        $normal = $document.Styles.Item('Normal')
        $normal.Font.Name = 'Calibri'
        $normal.Font.Size = 11
        $normal.ParagraphFormat.SpaceAfter = 6

        $footer = $section.Footers.Item(1).Range
        $footer.Text = 'HP iLO 5 Health Check Report  |  Page '
        $footer.Font.Name = 'Calibri'
        $footer.Font.Size = 8.5
        $footer.Font.Color = ConvertTo-WordColor '666666'
        $footer.ParagraphFormat.Alignment = 2
        $footer.Collapse(0)
        $footer.Fields.Add($footer, 33) | Out-Null

        Add-WordParagraph $document 'HP iLO 5 Health Check Report' 25 $false '1F4E78' 6
        Add-WordParagraph $document "Target: $($Data.Target)  |  Generated: $($Data.GeneratedAt)" 10 $false '555555' 15
        Add-WordHeading $document 'Executive health summary'
        $summaryRows = @($Data.ServerStatus.GetEnumerator() | ForEach-Object {
            [PSCustomObject][ordered]@{ Item = $_.Key; Value = $_.Value }
        })
        Add-WordTable $document $summaryRows 'Server status was not available.'

        $sections = @(
            @('Temperatures', 'Temperatures'),
            @('Fans', 'Fans'),
            @('Power supplies', 'PowerSupplies'),
            @('Storage', 'Storage'),
            @('Memory', 'Memory'),
            @('Processors', 'Processors'),
            @('Firmware', 'Firmware'),
            @('Event logs', 'EventLogs')
        )
        foreach ($item in $sections) {
            Add-WordHeading $document $item[0]
            $records = @($Data.PSObject.Properties[$item[1]].Value)
            Add-WordTable $document $records "No $($item[0].ToLowerInvariant()) data was returned."
        }

        if (@($Data.CollectionNotes).Count -gt 0) {
            Add-WordHeading $document 'Collection notes'
            foreach ($note in $Data.CollectionNotes) {
                Add-WordParagraph $document "- $note" 9.5 $false '666666' 4
            }
        }

        $resolved = [IO.Path]::GetFullPath($OutputPath)
        $directory = Split-Path -Parent $resolved
        if (-not (Test-Path $directory)) { New-Item -ItemType Directory -Path $directory -Force | Out-Null }
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
        $reportPath = New-WordHealthReport $data $OutputPath
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
