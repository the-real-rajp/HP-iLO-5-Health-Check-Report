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

Write-Host 'All PowerShell smoke tests passed.' -ForegroundColor Green
