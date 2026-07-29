# Generates the checked-in placeholder report without contacting an iLO.

. (Join-Path $PSScriptRoot '..\HP-iLO5-HealthReport.ps1')

$reportData = [PSCustomObject]@{
    Target = 'https://ilo.example.com'
    GeneratedAt = '2026-07-27T14:30:00Z'
    ServerStatus = [ordered]@{
        'Product name' = 'HPE ProLiant DL380 Gen10'
        'Serial number' = 'EXAMPLE123'
        Health = 'OK'
    }
    IloInformation = @([PSCustomObject][ordered]@{
        Name = 'iLO 5'
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
        'Connection status' = 'NotEnabled'
        'Workspace ID' = 'Not reported'
        'Next retry time' = 'N/A'
    })
    RemoteSupportRegistration = @([PSCustomObject][ordered]@{
        Registration = 'Registered'
        'Connection model' = 'CentralConnect'
        Destination = 'https://support.example.com'
        'Last transmission error' = 'None'
    })
    SecurityDashboard = @([PSCustomObject][ordered]@{
        Name = 'Overall Security Status'
        'Security status' = 'OK'
    })
    Processors = @([PSCustomObject][ordered]@{
        Name = 'CPU 1'
        Model = 'Intel(R) Xeon(R) Gold 6130 CPU'
        Health = 'OK'
    })
    Memory = @(
        [PSCustomObject][ordered]@{ Name = 'PROC 1 DIMM 8'; 'Capacity (MiB)' = 16384; Type = 'DDR4'; 'Speed (MHz)' = 2400; Health = 'OK' },
        [PSCustomObject][ordered]@{ Name = 'PROC 1 DIMM 10'; 'Capacity (MiB)' = 16384; Type = 'DDR4'; 'Speed (MHz)' = 2400; Health = 'OK' },
        [PSCustomObject][ordered]@{ Name = 'PROC 1 DIMM 12'; 'Capacity (MiB)' = 0; Type = 'N/A'; 'Speed (MHz)' = 'N/A'; Health = 'OK' }
    )
    SystemNetwork = @(
        [PSCustomObject][ordered]@{ Name = 'Embedded LOM 1'; Type = 'Ethernet interface'; 'MAC address' = '00:11:22:33:44:55'; 'IP address' = '192.0.2.20'; Link = 'LinkUp'; 'Speed (Mbps)' = 1000; Health = 'OK' },
        [PSCustomObject][ordered]@{ Name = 'Embedded LOM 2'; Type = 'Ethernet interface'; 'MAC address' = '00:11:22:33:44:56'; 'IP address' = 'N/A'; Link = 'N/A'; 'Speed (Mbps)' = 'N/A'; Health = 'OK' }
    )
    DeviceInventory = @([PSCustomObject][ordered]@{
        Location = 'Embedded LOM'
        'Product name' = 'HPE Ethernet 1Gb 4-port 331i Adapter'
        'Product version' = '1.0'
        'Firmware version' = '20.33.41'
        Status = 'OK'
    })
    Storage = @([PSCustomObject][ordered]@{
        Type = 'Drive'
        Name = 'Disk 1'
        Capacity = '1.8 TB'
        Health = 'OK'
    })
    Firmware = @([PSCustomObject][ordered]@{
        Name = 'iLO 5'
        Version = '3.10'
        Health = 'OK'
    })
    PowerSupplies = @([PSCustomObject][ordered]@{
        Name = 'Power Supply 1'
        Model = 'HPE 800W Flex Slot Platinum Hot Plug Power Supply Kit'
        Health = 'OK'
    })
    Fans = @([PSCustomObject][ordered]@{
        Name = 'Fan 1'
        Reading = '24%'
        Health = 'OK'
    })
    Temperatures = @(
        [PSCustomObject][ordered]@{ Name = '01-Inlet Ambient'; 'Reading (C)' = 22; 'Upper critical (C)' = 42; Health = 'OK' },
        [PSCustomObject][ordered]@{ Name = '02-System Board'; 'Reading (C)' = 0; 'Upper critical (C)' = 'N/A'; Health = 'OK' }
    )
    IloDedicatedNetworkPort = @(
        [PSCustomObject][ordered]@{ Setting = 'Configured for iLO'; Value = 'True' },
        [PSCustomObject][ordered]@{ Setting = 'Health'; Value = 'OK' },
        [PSCustomObject][ordered]@{ Setting = 'Link status'; Value = 'LinkUp' }
    )
    IloSharedNetworkPort = @([PSCustomObject][ordered]@{
        Setting = 'Configured for iLO'
        Value = 'False'
    })
    EventLogs = @()
    CollectionNotes = @()
}

$outputPath = Join-Path $PSScriptRoot 'sample-health-report.docx'
New-OpenXmlHealthReport -Data $reportData -OutputPath $outputPath -CustomerName 'Example Customer' -LogoPath (Join-Path (Split-Path -Parent $PSScriptRoot) 'images\winslow-technology-group-logo.png')
