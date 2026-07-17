"""Generate a synthetic report for layout review without contacting an iLO."""

from pathlib import Path

from ilo_health_report.report import write_report


SAMPLE_DATA = {
    "generated_at": "2026-07-17T12:00:00+00:00",
    "target": "https://ilo.example.com",
    "server_status": {
        "Name": "server-01",
        "Model": "ProLiant DL380 Gen10",
        "Manufacturer": "HPE",
        "Serial number": "REDACTED",
        "Power state": "On",
        "BIOS version": "U30 v2.90",
        "Health": "OK",
    },
    "temperatures": [
        {"Name": "Ambient", "Reading (°C)": 22, "Upper critical (°C)": 42, "Health": "OK", "State": "Enabled"},
        {"Name": "CPU 1", "Reading (°C)": 41, "Upper critical (°C)": 85, "Health": "OK", "State": "Enabled"},
    ],
    "fans": [{"Name": "Fan 1", "Reading": 18, "Units": "Percent", "Health": "OK", "State": "Enabled"}],
    "power_supplies": [{"Name": "Power Supply 1", "Model": "800W", "Serial number": "REDACTED", "Capacity (W)": 800, "Health": "OK", "State": "Enabled"}],
    "storage": [{"Name": "Smart Array", "Description": "Controller", "Health": "OK", "State": "Enabled"}],
    "memory": [{"Name": "PROC 1 DIMM 1", "Capacity (MiB)": 32768, "Type": "DDR4", "Speed (MHz)": 2933, "Health": "OK", "State": "Enabled"}],
    "processors": [{"Name": "CPU 1", "Model": "Intel Xeon", "Cores": 16, "Threads": 32, "Health": "OK", "State": "Enabled"}],
    "firmware": [{"Name": "iLO 5", "Version": "3.10", "Updateable": True, "Health": "OK", "State": "Enabled"}],
    "event_logs": [{"Log": "Integrated Management Log", "Created": "2026-07-16T10:00:00Z", "Severity": "Warning", "Message": "Example maintenance event", "Repaired": False}],
    "collection_notes": [],
}


if __name__ == "__main__":
    output = Path(__file__).with_name("sample-health-report.docx")
    write_report(SAMPLE_DATA, output)
    print(output)
