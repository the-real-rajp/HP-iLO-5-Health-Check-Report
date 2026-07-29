# HP iLO 5 Health Check Report

Generate a Word-compatible health report for an HPE iLO 5 server with
PowerShell 7 and the Redfish API. Microsoft Word is optional.

The report covers:

- Information: server identity, iLO details, component status, and HPE Compute
  Ops Management connection state
- Remote Support: registration, connection model, destination, and recent
  transmission state
- Security Dashboard: overall status and individual security findings
- System Information: summary, processors, memory, host network interfaces,
  HPE server device inventory, and storage
- Firmware & OS Software
- Power & Thermal: power supplies, fans, and temperatures
- Integrated Management Log, iLO Event Log, and other advertised log services
- empty detail sections, `Unknown` or `Absent` fields, and hardware records
  reported as `Absent` are omitted
- Memory and temperature rows containing `N/A` are omitted. Network rows with
  no usable data are omitted, while adapters with an IP address remain visible.
  Fan readings include a percent sign when applicable and do not show a Units
  column.
- transient iLO connection failures are retried and duplicate notes are suppressed

## Report layout

- A branded cover page is followed by an Executive Overview, Recommended
  Action, Assessment Summary, and grouped evidence tables.
- A Health Check Status/Severity guide appears above Recommended Action.
- Sections with only `Unknown` or unavailable status data are not shown or
  counted.
- The report uses a 0.75-inch top margin and 0.5-inch margins on the other
  sides. Body text uses Calibri 10 pt; evidence tables use Calibri 9 pt with
  auto-fit columns.
- The Winslow Tech Group logo appears in the upper-left header, and the report
  identifier appears in the upper-right header.
- The footer has `Confidential` on the left, the Winslow Tech Group copyright
  notice in the center, and page numbering on the right.
- The script downloads the WTG logo from winslowtg.com for each report and
  uses the bundled logo if the download is unavailable.

## Assessment behavior

- Lifecycle Management, Management, and Administration summary rows are not
  included.
- Storage evidence includes advertised controllers, physical drives, and
  logical volumes.
- Only `Critical` event-log entries from the previous month contribute to the
  recommended-action logic.
- An unconfigured Dedicated Network Port is `IGNORED`; an unconfigured Shared
  Network Port is omitted.
- Remote Support is `HEALTHY` when registered, `RECOMMENDED` when explicitly
  unregistered or reporting a transmission error, and omitted when its
  registration is unknown.
- An `Ignored` Overall Security Status or individual Security Dashboard finding
  is assessed as `WARNING`; a Security Status of `Risk` is assessed as
  `CRITICAL`. Ignored findings do not add a separate Ignored column.
- HPE Compute Ops Management uses iLO's HPE `CloudConnect` status and never
  exposes the GreenLake activation key. It is optional: `NotEnabled` remains
  visible in Information but does not make that assessment `RECOMMENDED`.

The script starts at `/redfish/v1/` and follows the links advertised by iLO,
so it does not assume that every server uses the same system, chassis, or
manager identifier.

## Requirements

- Windows 10, Windows 11, or Windows Server
- PowerShell 7.0 or later
- Desktop Microsoft Word is optional. The script uses its built-in Open XML
  generator by default for consistent branded output; set
  `HP_ILO_USE_WORD_COM=1` only to opt into the legacy Word automation path.
- HTTPS network access to the iLO management interface
- An iLO account with read access to the requested Redfish resources

No extra PowerShell modules are required.

## Example report

The current report layout and table filtering are illustrated in
[examples/sample-health-report.docx](examples/sample-health-report.docx).
It uses placeholder infrastructure data only and can be regenerated with
`examples/Generate-SampleReport.ps1`.

## Run

Open PowerShell 7 in the repository directory:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
pwsh -NoProfile -File ./HP-iLO5-HealthReport.ps1
```

The script prompts for an iLO IP address or FQDN, customer name, and then
displays the standard Windows credential prompt. The password is not echoed or
saved in the report or the log. During the run, the console shows connection,
collection, and report-generation progress. Collection activity is grouped so
you can see the area being queried and its current step, for example:

```text
Collecting iLO
  - System
  - Chassis
  - Managers
Collecting Network Configuration
  - iLO Ethernet interfaces
Collecting System Information
  - Memory
  - Processors
```

When `-OutputPath` is omitted, the report is saved in the same folder as
`HP-iLO5-HealthReport.ps1`, regardless of PowerShell's current directory. Its
default filename is `iLO Health Check - Customer Name - IP-or-FQDN.docx`.
A detailed log is written for every run. By default it is saved in the script's
`logs` folder as `iLO Health Check - Customer Name - IP-or-FQDN -
YYYYMMDD-HHMMSS.log` and records progress, collection notes, Redfish request
activity, and errors without recording credentials or session tokens.

Parameters can also be supplied directly:

```powershell
pwsh -NoProfile -File ./HP-iLO5-HealthReport.ps1 `
    -IloAddress 'ilo.example.com' `
    -CustomerName 'Example Customer' `
    -OutputPath '.\reports\server-01-health.docx'
```

For an iLO with a self-signed certificate in a trusted lab:

```powershell
pwsh -NoProfile -File ./HP-iLO5-HealthReport.ps1 -IloAddress '192.0.2.10' -SkipCertificateCheck
```

Certificate verification remains enabled by default. Other options:

```text
-Credential            PSCredential to use instead of prompting
-CustomerName          Customer name displayed in the cover, header, and overview;
                       prompted when omitted
-LogPath               Detailed log-file path; defaults to the timestamped logs folder
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

The smoke tests validate script parsing, report data shaping, timestamped log
creation, grouped progress formatting, and the built-in DOCX package without
contacting an iLO or starting Microsoft Word.

## Maintainer guide

See [AGENT.md](AGENT.md) for the project conventions, validation steps, and
reporting rules used when modifying this repository.

## References

- [HPE iLO 5 Redfish API reference](https://hewlettpackard.github.io/ilo-rest-api-docs/ilo5/)
- [DMTF Redfish standard](https://www.dmtf.org/standards/redfish)
