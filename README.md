# HP iLO 5 Health Check Report

Generate a Word-compatible health report for an HPE iLO 5 server with
PowerShell and the Redfish API. Microsoft Word is optional.

The report covers:

- server status
- temperatures
- fans
- power supplies
- storage
- memory
- processors
- firmware
- iLO Dedicated and Shared Network Port configuration
- Integrated Management Log, iLO Event Log, and other advertised log services
- iLO Security Dashboard overall status and risks
- empty detail sections and hardware records reported as `Absent` are omitted
- transient iLO connection failures are retried and duplicate notes are suppressed

The Word output follows a two-part health-assessment format: a branded cover
page followed by an Executive Overview with a Recommended Action,
fourteen-section Assessment Summary, and detailed evidence tables. Storage
evidence represents controllers, physical drives, and logical volumes when iLO
advertises them. Administration event-log evidence is limited to entries with
`Critical` severity from the previous month. A Dedicated or Shared Network Port
that is not configured for iLO is marked `IGNORED`; the Shared Network Port
section still lists its advertised interface, IP, VLAN, NIC, and port settings.
An `Ignored` Overall Security Status is treated as healthy, and individual
Security Dashboard findings with `Ignored = True` display `Ignored`.

The script starts at `/redfish/v1/` and follows the links advertised by iLO,
so it does not assume that every server uses the same system, chassis, or
manager identifier.

## Requirements

- Windows 10, Windows 11, or Windows Server
- Windows PowerShell 5.1 or PowerShell 7
- Desktop Microsoft Word is optional. When it is unavailable, the script uses
  its built-in Open XML generator to create the `.docx` directly.
- HTTPS network access to the iLO management interface
- An iLO account with read access to the requested Redfish resources

No extra PowerShell modules are required.

## Run

Open PowerShell in the repository directory:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
./HP-iLO5-HealthReport.ps1
```

The script prompts for an iLO IP address or FQDN, customer name, and then
displays the standard Windows credential prompt. The password is not echoed or
saved in the report.
When `-OutputPath` is omitted, the report is saved in the same folder as
`HP-iLO5-HealthReport.ps1`, regardless of PowerShell's current directory.

Parameters can also be supplied directly:

```powershell
./HP-iLO5-HealthReport.ps1 `
    -IloAddress 'ilo.example.com' `
    -CustomerName 'Example Customer' `
    -OutputPath '.\reports\server-01-health.docx'
```

For an iLO with a self-signed certificate in a trusted lab:

```powershell
./HP-iLO5-HealthReport.ps1 -IloAddress '192.0.2.10' -SkipCertificateCheck
```

Certificate verification remains enabled by default. Other options:

```text
-Credential            PSCredential to use instead of prompting
-CustomerName          Customer name displayed in the cover, header, and overview;
                       prompted when omitted
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

Run the smoke tests with PowerShell:

```powershell
pwsh -NoProfile -File ./tests/Smoke.Tests.ps1
```

The smoke tests validate script parsing, pure data-shaping helpers, and the
built-in DOCX package without contacting an iLO or starting Microsoft Word.

## References

- [HPE iLO 5 Redfish API reference](https://hewlettpackard.github.io/ilo-rest-api-docs/ilo5/)
- [DMTF Redfish standard](https://www.dmtf.org/standards/redfish)
