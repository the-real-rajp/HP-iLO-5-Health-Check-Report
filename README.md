# HP iLO 5 Health Check Report

Generate a Microsoft Word health report for an HPE iLO 5 server through its
Redfish API. The report covers:

- server status
- temperatures
- fans
- power supplies
- storage
- memory
- processors
- firmware
- Integrated Management Log and other available system event logs

The tool discovers Redfish resource links instead of assuming that every iLO
uses the same system or chassis identifier.

## Requirements

- Python 3.13 or newer (designed for Python 3.13.14)
- Network access to the iLO management interface over HTTPS
- An iLO account with read access to the requested Redfish resources

## Install

```bash
python3.13 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -e .
```

For development and tests:

```bash
python -m pip install -e '.[dev]'
pytest
```

## Run

Run without a host argument to be prompted for an IP address or FQDN:

```bash
ilo-health-report
```

Or provide it directly:

```bash
ilo-health-report --host ilo.example.com
```

The username and password are requested securely at the terminal. The tool
creates a Redfish session and does not save credentials in the report.

By default, HTTPS certificates are verified. For an iLO with a self-signed
certificate in a trusted lab, verification can be disabled explicitly:

```bash
ilo-health-report --host 192.0.2.10 --insecure
```

Other useful options:

```text
--output PATH         Destination .docx path
--timeout SECONDS     Per-request timeout (default: 30)
--max-log-entries N   Maximum entries per event log (default: 100)
```

## Security notes

- Prefer a dedicated, least-privilege iLO account.
- Keep certificate verification enabled in production.
- Do not put credentials in command-line arguments, source files, or inventory
  files. Passwords supplied at the prompt are not echoed.
- The report may contain serial numbers, firmware versions, hostnames, and log
  messages. Handle it as operationally sensitive data.

## Redfish coverage

The collector begins at `/redfish/v1/` and follows advertised links to systems,
chassis, thermal, power, storage, memory, processors, firmware inventory, and
log services. Availability varies by server model, installed hardware, iLO
firmware, privileges, and storage-controller Redfish support. An unavailable
resource is recorded in the report instead of aborting the entire collection.

Official references:

- [HPE iLO 5 Redfish API reference](https://hewlettpackard.github.io/ilo-rest-api-docs/ilo5/)
- [DMTF Redfish schemas and specifications](https://www.dmtf.org/standards/redfish)

