$ErrorActionPreference = 'Stop'
$scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'HP-iLO5-HealthReport.ps1'

$tokens = $null
$errors = $null
[void][Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$errors)
if ($errors.Count -gt 0) {
    throw "PowerShell parser errors: $($errors.Message -join '; ')"
}

. $scriptPath

if ($script:WdPaperLetter -ne 2) {
    throw 'The Word Letter paper-size constant must be 2.'
}
if ($script:WdPageBreak -ne 7 -or $script:WdFieldPage -ne 33 -or $script:WdFieldNumPages -ne 26) {
    throw 'One or more Word interop constants are incorrect.'
}
if ($script:WdPreferredWidthPoints -ne 3) {
    throw 'The Word preferred-width type must use points.'
}

function Assert-Equal {
    param($Actual, $Expected, [string]$Message)
    if ($Actual -ne $Expected) { throw "$Message. Expected '$Expected'; got '$Actual'." }
}

$resource = [PSCustomObject]@{
    Status = [PSCustomObject]@{ Health = 'OK'; State = 'Enabled' }
    Memory = [PSCustomObject]@{ '@odata.id' = '/redfish/v1/Systems/1/Memory' }
    Links = [PSCustomObject]@{
        SecurityService = [PSCustomObject]@{ '@odata.id' = '/redfish/v1/Managers/1/SecurityService' }
    }
}
Assert-Equal (Get-HealthValue $resource) 'OK' 'Health extraction failed'
Assert-Equal (Get-StateValue $resource) 'Enabled' 'State extraction failed'
Assert-Equal (Get-HeaderValue @{ 'X-Auth-Token' = [string[]]@('test-token') } 'X-Auth-Token') 'test-token' 'Array response header was not reduced to a scalar value'
Assert-Equal (Get-HeaderValue @{ Location = '/redfish/v1/SessionService/Sessions/1' } 'Location') '/redfish/v1/SessionService/Sessions/1' 'Scalar response header extraction failed'
Assert-Equal (Get-RedfishLink $resource 'Memory') '/redfish/v1/Systems/1/Memory' 'Link extraction failed'
Assert-Equal (Get-RedfishLinkAny $resource 'SecurityService') '/redfish/v1/Managers/1/SecurityService' 'Nested link extraction failed'
Assert-Equal (Resolve-RedfishUri ([uri]'https://ilo.example.com/') '/redfish/v1/') 'https://ilo.example.com/redfish/v1/' 'URI resolution failed'
Assert-Equal (ConvertTo-WordColor '1F4E78') 7884319 'Word color conversion failed'

$temperature = Convert-Temperature ([PSCustomObject]@{
    Name = 'Ambient'
    ReadingCelsius = 22
    UpperThresholdCritical = 42
    Status = [PSCustomObject]@{ Health = 'OK'; State = 'Enabled' }
})
Assert-Equal $temperature.Name 'Ambient' 'Temperature name conversion failed'
Assert-Equal $temperature.Health 'OK' 'Temperature health conversion failed'
Assert-Equal (Test-ReportRecordPresent $temperature) $true 'Enabled temperature should be included'
$absentTemperature = [PSCustomObject]@{ Name = 'Unused sensor'; State = 'Absent' }
Assert-Equal (Test-ReportRecordPresent $absentTemperature) $false 'Absent temperature should be excluded'

$securityParameter = Convert-SecurityParameter ([PSCustomObject]@{
    Name = 'Minimum password length'
    SecurityStatus = 'Risk'
    State = '8 characters'
    RecommendedAction = 'Increase the minimum password length.'
    Ignore = $false
})
Assert-Equal $securityParameter.'Security status' 'Risk' 'Security Dashboard status conversion failed'
if ($securityParameter.PSObject.Properties.Name -contains 'Recommended action') {
    throw 'Security Dashboard output should not contain Recommended action.'
}
if ($securityParameter.PSObject.Properties.Name -contains 'Ignored') {
    throw 'Security Dashboard output should not contain an Ignored column.'
}
Assert-Equal (Get-AssessmentStatus @($securityParameter)) 'CRITICAL' 'An unignored Security Dashboard risk must be critical'
$ignoredSecurityParameter = Convert-SecurityParameter ([PSCustomObject]@{
    Name = 'Ignored security setting'
    SecurityStatus = 'Risk'
    State = 'Enabled'
    Ignore = $true
})
Assert-Equal $ignoredSecurityParameter.'Security status' 'Ignored' 'An ignored Security Dashboard item should display Ignored'
Assert-Equal (Get-SecurityAssessmentStatus @($ignoredSecurityParameter)) 'HEALTHY' 'An ignored Security Dashboard item should not reduce section health'
$ignoredSecurityOverview = Convert-SecurityDashboardOverview ([PSCustomObject]@{
    OverallSecurityStatus = 'Ignored'
    ServerConfigurationLockStatus = 'Disabled'
})
Assert-Equal (Get-SecurityAssessmentStatus @($ignoredSecurityOverview, $securityParameter)) 'HEALTHY' 'Ignored Overall Security Status should mark Security healthy'

$firmwareRecord = Convert-Firmware ([PSCustomObject]@{
    Name = 'iLO 5'
    Version = '3.10'
    Updateable = $true
})
if ($firmwareRecord.PSObject.Properties.Name -contains 'State') {
    throw 'Firmware report output should not contain State.'
}
if ($firmwareRecord.PSObject.Properties.Name -contains 'Updateable') {
    throw 'Firmware report output should not contain Updateable.'
}
Assert-Equal $firmwareRecord.Health 'OK' 'Firmware without an advertised health value should display OK'

$sharedInterface = Convert-IloNetworkInterface ([PSCustomObject]@{
    Name = 'Manager Shared Network Interface'
    InterfaceEnabled = $true
    LinkStatus = 'LinkUp'
    MACAddress = '00:11:22:33:44:55'
    PermanentMACAddress = '00:11:22:33:44:55'
    IPv4Addresses = @([PSCustomObject]@{
        Address = '192.0.2.20'
        AddressOrigin = 'Static'
        SubnetMask = '255.255.255.0'
        Gateway = '192.0.2.1'
    })
    Status = [PSCustomObject]@{ Health = 'OK' }
    Oem = [PSCustomObject]@{
        Hpe = [PSCustomObject]@{
            InterfaceType = 'Shared'
            SharedNetworkPortOptions = [PSCustomObject]@{ NIC = 'LOM'; Port = 1 }
        }
    }
})
Assert-Equal $sharedInterface.InterfaceType 'Shared' 'Shared iLO interface type conversion failed'
Assert-Equal (Get-IloNetworkPortAssessmentStatus $sharedInterface.Rows) 'HEALTHY' 'Enabled shared iLO network interface should be healthy'
Assert-Equal (Get-IloNetworkPortAssessmentStatus @([PSCustomObject]@{ Setting = 'Configured for iLO'; Value = 'True' })) $null 'A configured NIC with no health evidence should not be assessed'
Assert-Equal (@($sharedInterface.Rows | Where-Object Setting -eq 'Shared NIC')[0].Value) 'LOM' 'Shared NIC configuration was not collected'
$dedicatedInterface = Convert-IloNetworkInterface ([PSCustomObject]@{
    Name = 'Manager Dedicated Network Interface'
    InterfaceEnabled = $false
    LinkStatus = 'NoLink'
    Status = [PSCustomObject]@{ Health = 'OK' }
    Oem = [PSCustomObject]@{ Hpe = [PSCustomObject]@{ InterfaceType = 'Dedicated' } }
})
Assert-Equal (Get-IloNetworkPortAssessmentStatus $dedicatedInterface.Rows) 'IGNORED' 'Unconfigured dedicated iLO NIC should be ignored'
Assert-Equal (Get-IloNetworkPortAssessmentStatus $dedicatedInterface.Rows -OmitWhenUnconfigured) $null 'An unconfigured shared iLO NIC should be omitted'
Assert-Equal (@($dedicatedInterface.Rows | Where-Object Setting -eq 'Assessment note')[0].Value) 'iLO is not configured to use this NIC.' 'Dedicated NIC ignore note is missing'

$server = Convert-ServerStatus ([PSCustomObject]@{
    Name = 'Server'
    Model = 'ProLiant DL380 Gen10'
    Status = [PSCustomObject]@{ HealthRollup = 'Warning' }
})
Assert-Equal $server.Health 'Warning' 'Server health conversion failed'
Assert-Equal $server.'Product name' 'ProLiant DL380 Gen10' 'Server model conversion failed'

$iloInformation = Convert-IloInformation ([PSCustomObject]@{
    Name = 'iLO 5'
    Model = 'iLO 5'
    ManagerType = 'BMC'
    FirmwareVersion = '3.10'
    DateTime = '2026-07-24T12:00:00Z'
    Status = [PSCustomObject]@{ Health = 'OK'; State = 'Enabled' }
})
Assert-Equal $iloInformation.'Firmware version' '3.10' 'iLO firmware conversion failed'
Assert-Equal $iloInformation.Health 'OK' 'iLO health conversion failed'
if ($iloInformation.PSObject.Properties.Name -contains 'State') {
    throw 'iLO report output should not contain State.'
}

$computeOps = Convert-ComputeOpsManagement ([PSCustomObject]@{
    Name = 'iLO 5'
    Oem = [PSCustomObject]@{
        Hpe = [PSCustomObject]@{
            CloudConnect = [PSCustomObject]@{
                ActivationKey = 'must-not-be-reported'
                CloudConnectStatus = 'Connected'
                WorkspaceId = 'workspace-123'
                FailReason = 'None'
            }
        }
    }
})
Assert-Equal $computeOps.'Connection status' 'Connected' 'Compute Ops Management status conversion failed'
Assert-Equal $computeOps.'Workspace ID' 'workspace-123' 'Compute Ops Management workspace conversion failed'
if ($computeOps.PSObject.Properties.Name -contains 'Failure reason') {
    throw 'Compute Ops Management output should not contain Failure reason.'
}
if ($computeOps.PSObject.Properties.Name -contains 'Next retry time') {
    throw 'Compute Ops Management output should not contain Next retry time.'
}
Assert-Equal (Get-AssessmentStatus @([PSCustomObject]@{ 'Connection status' = 'ConnectionFailed' })) 'CRITICAL' 'Failed Compute Ops connection should be critical'
if ($computeOps.PSObject.Properties.Name -contains 'Activation key' -or
    (($computeOps.PSObject.Properties.Value | ForEach-Object { [string]$_ }) -join ' ') -match 'must-not-be-reported') {
    throw 'Compute Ops Management output exposed the activation key.'
}

$systemNic = Convert-SystemNetworkInterface ([PSCustomObject]@{
    Name = 'Embedded LOM 1'
    MACAddress = '00:11:22:33:44:66'
    LinkStatus = 'LinkUp'
    SpeedMbps = 1000
    IPv4Addresses = @([PSCustomObject]@{ Address = '192.0.2.30' })
    Status = [PSCustomObject]@{ Health = 'OK' }
})
Assert-Equal $systemNic.'MAC address' '00:11:22:33:44:66' 'System network MAC conversion failed'
Assert-Equal $systemNic.'IP address' '192.0.2.30' 'System network IP conversion failed'

$deviceRecord = Convert-DeviceInventory ([PSCustomObject]@{
    Name = 'Smart Array'
    DeviceType = 'StorageController'
    Location = 'Embedded RAID'
    ProductVersion = 'B'
    FirmwareVersion = '6.52'
    Status = [PSCustomObject]@{ Health = 'OK'; State = 'Enabled' }
})
Assert-Equal $deviceRecord.Location 'Embedded RAID' 'Device inventory location conversion failed'
Assert-Equal $deviceRecord.'Firmware version' '6.52' 'Device inventory firmware conversion failed'
Assert-Equal $deviceRecord.Status 'OK' 'Device inventory status conversion failed'
$percentFan = Convert-Fan ([PSCustomObject]@{ Name = 'Fan 2'; Reading = 23; ReadingUnits = 'Percent'; Status = [PSCustomObject]@{ Health = 'OK' } })
Assert-Equal $percentFan.Reading '23%' 'Percent fan reading should include the percent symbol'
$absentDeviceData = [PSCustomObject]@{
    DeviceInventory = @([PSCustomObject]@{ Location = 'Empty slot'; Status = 'Absent' })
}
Assert-Equal @(Get-ReportRecords -Data $absentDeviceData -PropertyName 'DeviceInventory').Count 0 'Absent device inventory records should be omitted'
Assert-Equal (Get-AssessmentStatus @([PSCustomObject]@{ Status = 'Absent' })) $null 'Absent device inventory records should not be assessed'

foreach ($convertedRecord in @(
    $temperature,
    (Convert-Fan ([PSCustomObject]@{ Name = 'Fan 1'; Status = [PSCustomObject]@{ Health = 'OK'; State = 'Enabled' } })),
    (Convert-PowerSupply ([PSCustomObject]@{ Name = 'Power Supply 1'; Status = [PSCustomObject]@{ Health = 'OK'; State = 'Enabled' } })),
    (Convert-Memory ([PSCustomObject]@{ Name = 'DIMM 1'; Status = [PSCustomObject]@{ Health = 'OK'; State = 'Enabled' } })),
    (Convert-Processor ([PSCustomObject]@{ Name = 'CPU 1'; Status = [PSCustomObject]@{ Health = 'OK'; State = 'Enabled' } }))
)) {
    if ($convertedRecord.PSObject.Properties.Name -contains 'State') {
        throw 'Hardware detail output should not contain State.'
    }
}

$remoteSupport = Convert-RemoteSupportRegistration ([PSCustomObject]@{
    RemoteSupportEnabled = $true
    ConnectModel = 'CentralConnect'
    DestinationURL = 'https://remote-support.example'
    LastTransmissionError = 'None'
})
Assert-Equal $remoteSupport.Registration 'Registered' 'Remote Support registration conversion failed'
Assert-Equal (Get-RemoteSupportAssessmentStatus @($remoteSupport)) 'HEALTHY' 'Registered Remote Support should be healthy'
Assert-Equal (Get-RemoteSupportAssessmentStatus @([PSCustomObject]@{ Registration = 'Not registered' })) 'RECOMMENDED' 'Unregistered Remote Support should be recommended'
Assert-Equal (Get-RemoteSupportAssessmentStatus @()) $null 'Uncollected Remote Support should not be assessed'
Assert-Equal (Get-AssessmentStatus @([PSCustomObject]@{ Health = 'Unknown'; State = 'Unknown' })) $null 'Unknown-only evidence should not be assessed'

$unknownReportData = [PSCustomObject]@{
    Example = @([PSCustomObject][ordered]@{ Name = 'Adapter 1'; Health = 'Unknown'; State = 'Enabled' })
}
$unknownReportRows = @(Get-ReportRecords -Data $unknownReportData -PropertyName 'Example')
if ($unknownReportRows[0].PSObject.Properties.Name -contains 'Health') {
    throw 'An all-Unknown report column should be omitted.'
}
if (-not (Get-Command New-IloSession).Parameters.ContainsKey('IgnoreCertificateErrors')) {
    throw 'New-IloSession is missing the internal certificate-control parameter.'
}

$emptyNotes = [System.Collections.Generic.List[string]]::new()
$emptyResult = @(Get-SafeCollection `
    -Session ([PSCustomObject]@{}) `
    -Uri $null `
    -Notes $emptyNotes `
    -Label 'test resource')
Assert-Equal $emptyResult.Count 0 'Missing collection should return no records'
Assert-Equal $emptyNotes.Count 1 'An empty notes collection did not accept a collection note'
Add-CollectionNote -Notes $emptyNotes -Message $emptyNotes[0]
Assert-Equal $emptyNotes.Count 1 'Duplicate collection notes should be suppressed'

$assessmentData = [PSCustomObject]@{
    GeneratedAt = '2026-07-24T12:00:00Z'
    ServerStatus = [ordered]@{ Health = 'OK'; State = 'Enabled' }
    Temperatures = @([PSCustomObject]@{ 'Upper critical (C)' = 42; Health = 'OK'; State = 'Enabled' })
    Fans = @([PSCustomObject]@{ Health = 'OK'; State = 'Enabled' })
    PowerSupplies = @([PSCustomObject]@{ Health = 'OK'; State = 'Enabled' })
    Memory = @([PSCustomObject]@{ Health = 'OK'; State = 'Enabled' })
    Processors = @([PSCustomObject]@{ Health = 'OK'; State = 'Enabled' })
    Firmware = @([PSCustomObject]@{ Health = 'OK'; State = 'Enabled' })
    Management = @([PSCustomObject]@{ Health = 'OK'; State = 'Enabled' })
    RemoteSupportRegistration = @([PSCustomObject]@{ Registration = 'Registered'; 'Last transmission error' = 'None' })
    IloDedicatedNetworkPort = $dedicatedInterface.Rows
    IloSharedNetworkPort = $sharedInterface.Rows
    EventLogs = @(
        [PSCustomObject]@{ Created = '2026-07-20T10:00:00Z'; Severity = 'Critical'; Log = 'Integrated Management Log'; Message = 'Recent critical event'; Repaired = $false }
        [PSCustomObject]@{ Created = '2026-07-20T11:00:00Z'; Severity = 'Warning'; Log = 'Integrated Management Log'; Message = 'Recent warning event'; Repaired = $false }
        [PSCustomObject]@{ Created = '2026-05-01T10:00:00Z'; Severity = 'Critical'; Log = 'Integrated Management Log'; Message = 'Old critical event'; Repaired = $false }
    )
    SecurityDashboard = @([PSCustomObject]@{ 'Security status' = 'Ok'; Ignored = $false })
}
$assessment = @(New-AssessmentSummary $assessmentData)
$expectedSections = @(
    'Information', 'System Information', 'Firmware & OS Software',
    'Power & Thermal',
    'Performance', 'iLO Dedicated Network Port', 'iLO Shared Network Port',
    'Remote Support', 'Security Dashboard'
)
Assert-Equal $assessment.Count 9 'Assessment summary row count is incorrect'
Assert-Equal (($assessment.Section -join '|')) ($expectedSections -join '|') 'Assessment summary order is incorrect'
Assert-Equal $assessment[0].Status 'HEALTHY' 'Healthy assessment evidence was not recognized'
Assert-Equal $assessment[3].Status 'HEALTHY' 'A column name containing critical must not create a critical assessment'
Assert-Equal $assessment[5].Status 'IGNORED' 'Unconfigured dedicated iLO NIC assessment should be ignored'
Assert-Equal $assessment[6].Status 'HEALTHY' 'Configured shared iLO NIC assessment should be healthy'
Assert-Equal $assessment[7].Status 'HEALTHY' 'Registered Remote Support should be healthy'
Assert-Equal $assessment[8].Status 'HEALTHY' 'A healthy Security Dashboard should produce a healthy Security Dashboard assessment'
$criticalRecentEvents = @(Get-CriticalRecentEventLogs $assessmentData)
Assert-Equal $criticalRecentEvents.Count 1 'Event-log report filtering should retain only recent critical events'
Assert-Equal $criticalRecentEvents[0].Message 'Recent critical event' 'The wrong event survived report filtering'
$recommendedAction = Get-RecommendedActionText -Data $assessmentData -Assessment $assessment
if ($recommendedAction -notmatch 'critical iLO event-log entries') {
    throw 'Recommended Action did not include the recent critical event guidance.'
}

$nativeReportData = [PSCustomObject]@{
    Target = 'https://192.0.2.10'
    GeneratedAt = '2026-07-24T12:00:00Z'
    ServerStatus = [ordered]@{
        Name = 'Example ProLiant'
        Model = 'ProLiant DL380 Gen10'
        Health = 'OK'
        State = 'Enabled'
    }
    IloInformation = @([PSCustomObject][ordered]@{
        Name = 'iLO 5'
        Model = 'iLO 5'
        'Manager type' = 'BMC'
        'Firmware version' = '3.10'
        Health = 'OK'
    })
    StatusInformation = @([PSCustomObject][ordered]@{
        Component = 'Server'
        Health = 'OK'
        State = 'Enabled'
        Detail = 'Power: On'
    })
    ComputeOpsManagement = @([PSCustomObject][ordered]@{
        Manager = 'iLO 5'
        Supported = 'Yes'
        'Connection status' = 'Connected'
        'Workspace ID' = 'workspace-123'
        'Next retry time' = 'N/A'
    })
    Temperatures = @([PSCustomObject][ordered]@{
        Name = 'Ambient'
        'Reading (C)' = 22
        'Upper critical (C)' = 42
        Health = 'OK'
    })
    Fans = @([PSCustomObject][ordered]@{
        Name = 'Fan 1'
        'Reading (%)' = 23
        Health = 'OK'
    })
    PowerSupplies = @([PSCustomObject][ordered]@{
        Name = 'Power Supply 1'
        Model = 'Example PSU'
        Health = 'OK'
    })
    Storage = @([PSCustomObject][ordered]@{
        Type = 'Drive'
        Name = 'Disk 1'
        Capacity = '1.8 TB'
        Health = 'OK'
    })
    Memory = @([PSCustomObject][ordered]@{
        Name = 'DIMM 1'
        Capacity = '32 GB'
        Health = 'OK'
    })
    Processors = @([PSCustomObject][ordered]@{
        Name = 'CPU 1'
        Model = 'Example Processor'
        Health = 'OK'
    })
    SystemNetwork = @([PSCustomObject][ordered]@{
        Name = 'Embedded LOM 1'
        Type = 'Ethernet interface'
        'MAC address' = '00:11:22:33:44:66'
        'IP address' = '192.0.2.30'
        Link = 'LinkUp'
        'Speed (Mbps)' = 1000
        Health = 'OK'
    })
    DeviceInventory = @([PSCustomObject][ordered]@{
        Location = 'Embedded RAID'
        'Product name' = 'Smart Array'
        'Product version' = 'Unknown'
        'Firmware version' = '6.52'
        Status = 'Enabled'
    })
    Firmware = @([PSCustomObject][ordered]@{
        Name = 'iLO 5'
        Version = '3.10'
        Health = 'OK'
    })
    IloDedicatedNetworkPort = $dedicatedInterface.Rows
    IloSharedNetworkPort = $sharedInterface.Rows
    RemoteSupportRegistration = @([PSCustomObject][ordered]@{
        Registration = 'Registered'
        'Connection model' = 'CentralConnect'
        Destination = 'https://remote-support.example'
        'Last transmission error' = 'None'
    })
    Management = @([PSCustomObject][ordered]@{
        Name = 'iLO 5'
        Firmware = '3.10'
        State = 'Enabled'
    })
    SecurityDashboard = @([PSCustomObject][ordered]@{
        Name = 'Security State'
        'Security status' = 'Ok'
    })
    EventLogs = @(
        [PSCustomObject][ordered]@{
            Created = '2026-07-24T10:00:00Z'
            Severity = 'Critical'
            Log = 'Integrated Management Log'
            Message = 'Recent critical event for report validation.'
            Repaired = $false
        }
        [PSCustomObject][ordered]@{
            Created = '2026-07-23T10:00:00Z'
            Severity = 'Warning'
            Log = 'Integrated Management Log'
            Message = 'Recent warning event that must be omitted.'
            Repaired = $false
        }
        [PSCustomObject][ordered]@{
            Created = '2026-05-01T10:00:00Z'
            Severity = 'Critical'
            Log = 'Integrated Management Log'
            Message = 'Old critical event that must be omitted.'
            Repaired = $false
        }
    )
    CollectionNotes = @('Example report generated without Microsoft Word.')
}
$nativeReportPath = Join-Path ([IO.Path]::GetTempPath()) ('ilo-health-native-' + [guid]::NewGuid().ToString('N') + '.docx')
try {
    $createdReport = New-OpenXmlHealthReport -Data $nativeReportData -OutputPath $nativeReportPath -CustomerName 'Example Customer'
    Assert-Equal $createdReport $nativeReportPath 'Native DOCX generator returned the wrong path'
    if (-not (Test-Path $nativeReportPath)) { throw 'Native DOCX generator did not create a report.' }
    Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
    $zip = [IO.Compression.ZipFile]::OpenRead($nativeReportPath)
    try {
        $entryNames = @($zip.Entries.FullName)
        foreach ($requiredEntry in @(
            '[Content_Types].xml',
            '_rels/.rels',
            'word/document.xml',
            'word/styles.xml',
            'word/numbering.xml',
            'word/header1.xml',
            'word/footer1.xml'
        )) {
            if ($entryNames -notcontains $requiredEntry) {
                throw "Native DOCX is missing $requiredEntry."
            }
        }
        $documentEntry = $zip.GetEntry('word/document.xml')
        $reader = New-Object IO.StreamReader($documentEntry.Open())
        try { $documentText = $reader.ReadToEnd() }
        finally { $reader.Dispose() }
        foreach ($expectedText in @(
            'Recommended Action',
            'Assessment Summary',
            'Severity',
            'Information',
            'Server',
            'iLO',
            'Status',
            'HPE Compute Ops Management',
            'Remote Support',
            'Registration',
            'Registered',
            'Security Dashboard',
            'System Information',
            'Summary',
            'Processors',
            'Memory',
            'Network',
            'Device Inventory',
            'Storage',
            'Firmware &amp; OS Software',
            'Power &amp; Thermal',
            'Power Supply',
            'Fans',
            'Temperatures',
            'Connected',
            'Embedded LOM 1',
            'Smart Array'
        )) {
            if ($documentText -notmatch [regex]::Escape($expectedText)) {
                throw "Native DOCX is missing expected text: $expectedText."
            }
        }
        foreach ($unexpectedText in @(
            'Overall Health Score',
            'Administration - Event Logs',
            'Lifecycle Management',
            '>Administration<',
            '>Management<',
            'Collection Notes',
            'Failure reason',
            'Updateable',
            'Unknown',
            'Recent critical event for report validation.',
            'Recent warning event that must be omitted.',
            'Old critical event that must be omitted.'
        )) {
            if ($documentText -match [regex]::Escape($unexpectedText)) {
                throw "Native DOCX contains text that should have been omitted: $unexpectedText."
            }
        }
    }
    finally {
        $zip.Dispose()
    }
}
finally {
    Remove-Item $nativeReportPath -Force -ErrorAction SilentlyContinue
}

$originalRedfishGet = ${function:Invoke-RedfishGet}
$script:fakeResponses = @{
    '/empty-collection' = [PSCustomObject]@{ Members = @() }
    '/collection' = [PSCustomObject]@{
        Members = @([PSCustomObject]@{ '@odata.id' = '/items/1' })
        'Members@odata.nextLink' = '/collection?page=2'
    }
    '/collection?page=2' = [PSCustomObject]@{
        Members = @([PSCustomObject]@{ '@odata.id' = '/items/2' })
    }
    '/items/1' = [PSCustomObject]@{ Id = '1' }
    '/items/2' = [PSCustomObject]@{ Id = '2' }
}
function Invoke-RedfishGet {
    param($Session, [string]$Uri)
    if ($Uri -eq '/bad-collection') { throw 'The remote server returned an error: (400) Bad Request.' }
    return $script:fakeResponses[$Uri]
}
try {
    $members = @(Get-RedfishCollection -Session ([PSCustomObject]@{}) -Uri '/collection')
    Assert-Equal $members.Count 2 'Collection pagination failed'
    Assert-Equal $members[1].Id '2' 'Second collection page was not read'
    $fallbackNotes = [System.Collections.Generic.List[string]]::new()
    $fallbackMembers = @(Get-SafeCollectionFromUris `
        -Session ([PSCustomObject]@{}) `
        -Uris @('/bad-collection', '/collection') `
        -Notes $fallbackNotes `
        -Label 'storage')
    Assert-Equal $fallbackMembers.Count 2 'Collection URI fallback failed'
    Assert-Equal $fallbackNotes.Count 0 'Successful fallback should not leave a collection error note'
    $emptyFallbackMembers = @(Get-SafeCollectionFromUris `
        -Session ([PSCustomObject]@{}) `
        -Uris @('/empty-collection', '/collection') `
        -Notes $fallbackNotes `
        -Label 'systems')
    Assert-Equal $emptyFallbackMembers.Count 2 'An empty collection should not prevent endpoint fallback'
}
finally {
    Set-Item function:Invoke-RedfishGet $originalRedfishGet
}

$originalNewSession = ${function:New-IloSession}
$originalGetHealthData = ${function:Get-IloHealthData}
$originalNewReport = ${function:New-WordHealthReport}
$originalRemoveSession = ${function:Remove-IloSession}
$script:ignoreCertificateErrorsObserved = $false
$script:customerNameObserved = $null
$script:outputPathObserved = $null
function New-IloSession {
    param($BaseUri, $Credential, $TimeoutSec, [bool]$IgnoreCertificateErrors)
    $script:ignoreCertificateErrorsObserved = $IgnoreCertificateErrors
    return [PSCustomObject]@{ BaseUri = $BaseUri; SessionUri = $null }
}
function Get-IloHealthData { param($Session, $MaxLogEntries); return [PSCustomObject]@{} }
function New-WordHealthReport {
    param($Data, $OutputPath, $CustomerName)
    $script:customerNameObserved = $CustomerName
    $script:outputPathObserved = $OutputPath
    return $OutputPath
}
function Remove-IloSession { param($Session) }
try {
    $password = ConvertTo-SecureString 'smoke-test-only' -AsPlainText -Force
    $credential = [PSCredential]::new('test-user', $password)
    Invoke-IloHealthReport `
        -IloAddress '192.0.2.10' `
        -Credential $credential `
        -CustomerName 'Example Customer' `
        -OutputPath 'smoke-test.docx' `
        -SkipCertificateCheck
    Assert-Equal $script:ignoreCertificateErrorsObserved $true 'Certificate-skip forwarding failed'
    Assert-Equal $script:customerNameObserved 'Example Customer' 'Customer name forwarding failed'

    function Read-Host {
        param([string]$Prompt)
        if ($Prompt -eq 'Enter customer name') { return 'Prompted Customer' }
        throw "Unexpected Read-Host prompt: $Prompt"
    }
    try {
        Invoke-IloHealthReport `
            -IloAddress '192.0.2.10' `
            -Credential $credential
    }
    finally {
        Remove-Item function:Read-Host -Force
    }
    Assert-Equal $script:customerNameObserved 'Prompted Customer' 'Missing customer name should be prompted for'
    Assert-Equal (Split-Path -Parent $script:outputPathObserved) (Split-Path -Parent $scriptPath) 'Default report path should use the script directory'
    if ((Split-Path -Leaf $script:outputPathObserved) -notmatch '^ilo-health-192\.0\.2\.10-\d{8}-\d{6}\.docx$') {
        throw "Default report filename is incorrect: $script:outputPathObserved"
    }
}
finally {
    Set-Item function:New-IloSession $originalNewSession
    Set-Item function:Get-IloHealthData $originalGetHealthData
    Set-Item function:New-WordHealthReport $originalNewReport
    Set-Item function:Remove-IloSession $originalRemoveSession
}

Write-Host 'All PowerShell smoke tests passed.' -ForegroundColor Green
