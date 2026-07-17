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
}
Assert-Equal (Get-HealthValue $resource) 'OK' 'Health extraction failed'
Assert-Equal (Get-StateValue $resource) 'Enabled' 'State extraction failed'
Assert-Equal (Get-RedfishLink $resource 'Memory') '/redfish/v1/Systems/1/Memory' 'Link extraction failed'
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

$server = Convert-ServerStatus ([PSCustomObject]@{
    Name = 'Server'
    Model = 'ProLiant DL380 Gen10'
    Status = [PSCustomObject]@{ HealthRollup = 'Warning' }
})
Assert-Equal $server.Health 'Warning' 'Server health conversion failed'
Assert-Equal $server.Model 'ProLiant DL380 Gen10' 'Server model conversion failed'
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
    ServerStatus = [ordered]@{ Health = 'OK'; State = 'Enabled' }
    Temperatures = @([PSCustomObject]@{ 'Upper critical (C)' = 42; Health = 'OK'; State = 'Enabled' })
    Fans = @([PSCustomObject]@{ Health = 'OK'; State = 'Enabled' })
    PowerSupplies = @([PSCustomObject]@{ Health = 'OK'; State = 'Enabled' })
    Memory = @([PSCustomObject]@{ Health = 'OK'; State = 'Enabled' })
    Processors = @([PSCustomObject]@{ Health = 'OK'; State = 'Enabled' })
    Firmware = @([PSCustomObject]@{ Health = 'OK'; State = 'Enabled' })
    Management = @([PSCustomObject]@{ Health = 'OK'; State = 'Enabled' })
    EventLogs = @()
}
$assessment = @(New-AssessmentSummary $assessmentData)
$expectedSections = @(
    'Information', 'System Information', 'Firmware & OS Software',
    'iLO Federation', 'Remote Console & Media', 'Power & Thermal',
    'Performance', 'iLO Dedicated Network Port', 'iLO Shared Network Port',
    'Remote Support', 'Administration', 'Security', 'Management',
    'Lifecycle Management'
)
Assert-Equal $assessment.Count 14 'Assessment summary row count is incorrect'
Assert-Equal (($assessment.Section -join '|')) ($expectedSections -join '|') 'Assessment summary order is incorrect'
Assert-Equal $assessment[0].Status 'HEALTHY' 'Healthy assessment evidence was not recognized'
Assert-Equal $assessment[5].Status 'HEALTHY' 'A column name containing critical must not create a critical assessment'
Assert-Equal $assessment[3].Status 'RECOMMENDED' 'Unavailable assessment evidence should be recommended'
Assert-Equal (Get-OverallHealthScore $assessment) 65 'Overall health score calculation failed'

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
function New-IloSession {
    param($BaseUri, $Credential, $TimeoutSec, [bool]$IgnoreCertificateErrors)
    $script:ignoreCertificateErrorsObserved = $IgnoreCertificateErrors
    return [PSCustomObject]@{ BaseUri = $BaseUri; SessionUri = $null }
}
function Get-IloHealthData { param($Session, $MaxLogEntries); return [PSCustomObject]@{} }
function New-WordHealthReport {
    param($Data, $OutputPath, $CustomerName)
    $script:customerNameObserved = $CustomerName
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
}
finally {
    Set-Item function:New-IloSession $originalNewSession
    Set-Item function:Get-IloHealthData $originalGetHealthData
    Set-Item function:New-WordHealthReport $originalNewReport
    Set-Item function:Remove-IloSession $originalRemoveSession
}

Write-Host 'All PowerShell smoke tests passed.' -ForegroundColor Green
