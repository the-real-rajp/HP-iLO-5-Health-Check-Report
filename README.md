# HP iLO 5 Health Check Report

Generate a Microsoft Word health report for an HPE iLO 5 server with
PowerShell and the Redfish API.

The report covers:

- server status
- temperatures
- fans
- power supplies
- storage
- memory
- processors
- firmware
- Integrated Management Log, iLO Event Log, and other advertised log services

The script starts at `/redfish/v1/` and follows the links advertised by iLO,
so it does not assume that every server uses the same system, chassis, or
manager identifier.

## Requirements

- Windows 10, Windows 11, or Windows Server
- PowerShell 7.3 or newer
- Desktop Microsoft Word (used to create the `.docx` report)
- HTTPS network access to the iLO management interface
- An iLO account with read access to the requested Redfish resources

No extra PowerShell modules are required.

## Run

Open PowerShell in the repository directory:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
./HP-iLO5-HealthReport.ps1
```

The script prompts for an iLO IP address or FQDN and then displays the standard
Windows credential prompt. The password is not echoed or saved in the report.

Parameters can also be supplied directly:

```powershell
./HP-iLO5-HealthReport.ps1 `
    -IloAddress 'ilo.example.com' `
    -OutputPath '.\reports\server-01-health.docx'
```

For an iLO with a self-signed certificate in a trusted lab:

```powershell
./HP-iLO5-HealthReport.ps1 -IloAddress '192.0.2.10' -SkipCertificateCheck
```

Certificate verification remains enabled by default. Other options:

```text
-Credential            PSCredential to use instead of prompting
-TimeoutSec             Per-request timeout; default 30
-MaxLogEntries          Maximum entries collected from each log; default 100
-SkipCertificateCheck   Disable TLS validation for a trusted lab only
```

## Security notes

- Prefer a dedicated, least-privilege iLO account.
- Keep certificate validation enabled in production.
- Do not place passwords in scripts, command history, or source control.
- Reports can contain hostnames, serial numbers, firmware versions, and event
  messages. Handle them as operationally sensitive data.

## Validation

Run the cross-platform smoke tests with PowerShell 7:

```powershell
pwsh -NoProfile -File ./tests/Smoke.Tests.ps1
```

The smoke tests validate script parsing and pure data-shaping helpers without
contacting an iLO or starting Microsoft Word.

## References

- [HPE iLO 5 Redfish API reference](https://hewlettpackard.github.io/ilo-rest-api-docs/ilo5/)
- [DMTF Redfish standard](https://www.dmtf.org/standards/redfish)

