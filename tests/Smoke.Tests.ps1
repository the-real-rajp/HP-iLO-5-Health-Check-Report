$ErrorActionPreference = 'Stop'
$scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'HP-iLO5-HealthReport.ps1'

$tokens = $null
$errors = $null
[void][Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$errors)
if ($errors.Count -gt 0) {
    throw "PowerShell parser errors: $($errors.Message -join '; ')"
}

. $scriptPath

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

$originalRedfishGet = ${function:Invoke-RedfishGet}
$script:fakeResponses = @{
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
    return $script:fakeResponses[$Uri]
}
try {
    $members = @(Get-RedfishCollection -Session ([PSCustomObject]@{}) -Uri '/collection')
    Assert-Equal $members.Count 2 'Collection pagination failed'
    Assert-Equal $members[1].Id '2' 'Second collection page was not read'
}
finally {
    Set-Item function:Invoke-RedfishGet $originalRedfishGet
}

$originalNewSession = ${function:New-IloSession}
$originalGetHealthData = ${function:Get-IloHealthData}
$originalNewReport = ${function:New-WordHealthReport}
$originalRemoveSession = ${function:Remove-IloSession}
$script:ignoreCertificateErrorsObserved = $false
function New-IloSession {
    param($BaseUri, $Credential, $TimeoutSec, [bool]$IgnoreCertificateErrors)
    $script:ignoreCertificateErrorsObserved = $IgnoreCertificateErrors
    return [PSCustomObject]@{ BaseUri = $BaseUri; SessionUri = $null }
}
function Get-IloHealthData { param($Session, $MaxLogEntries); return [PSCustomObject]@{} }
function New-WordHealthReport { param($Data, $OutputPath); return $OutputPath }
function Remove-IloSession { param($Session) }
try {
    $password = ConvertTo-SecureString 'smoke-test-only' -AsPlainText -Force
    $credential = [PSCredential]::new('test-user', $password)
    Invoke-IloHealthReport `
        -IloAddress '192.0.2.10' `
        -Credential $credential `
        -OutputPath 'smoke-test.docx' `
        -SkipCertificateCheck
    Assert-Equal $script:ignoreCertificateErrorsObserved $true 'Certificate-skip forwarding failed'
}
finally {
    Set-Item function:New-IloSession $originalNewSession
    Set-Item function:Get-IloHealthData $originalGetHealthData
    Set-Item function:New-WordHealthReport $originalNewReport
    Set-Item function:Remove-IloSession $originalRemoveSession
}

Write-Host 'All PowerShell smoke tests passed.' -ForegroundColor Green
