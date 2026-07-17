"""Command-line entry point."""

from __future__ import annotations

import argparse
import getpass
import sys
from datetime import datetime
from pathlib import Path

import urllib3

from .collector import HealthCollector
from .redfish import RedfishClient, RedfishError
from .report import write_report


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Generate a Word health report for an HPE iLO 5 server."
    )
    parser.add_argument("--host", help="iLO IP address or FQDN")
    parser.add_argument("--output", type=Path, help="Destination .docx file")
    parser.add_argument("--timeout", type=float, default=30.0)
    parser.add_argument("--max-log-entries", type=int, default=100)
    parser.add_argument(
        "--insecure",
        action="store_true",
        help="Disable TLS certificate verification (trusted labs only)",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    host = args.host or input("Enter the iLO 5 IP address or FQDN: ").strip()
    username = input("iLO username: ").strip()
    password = getpass.getpass("iLO password: ")
    if args.insecure:
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    if args.output:
        destination = args.output
    else:
        stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        safe_host = host.replace("://", "-").replace("/", "-").replace(":", "-")
        destination = Path("reports") / f"ilo-health-{safe_host}-{stamp}.docx"

    try:
        with RedfishClient(
            host,
            username,
            password,
            verify=not args.insecure,
            timeout=args.timeout,
        ) as client:
            data = HealthCollector(
                client, max_log_entries=max(0, args.max_log_entries)
            ).collect()
        output = write_report(data, destination)
    except (RedfishError, ValueError, OSError) as error:
        print(f"Error: {error}", file=sys.stderr)
        return 1

    print(f"Word report created: {output.resolve()}")
    return 0

