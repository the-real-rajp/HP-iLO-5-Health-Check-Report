"""Collect health data by traversing links advertised by Redfish."""

from __future__ import annotations

from datetime import UTC, datetime
from typing import Any

from .redfish import RedfishClient, RedfishError


def _link(resource: dict[str, Any], key: str) -> str | None:
    value = resource.get(key)
    return value.get("@odata.id") if isinstance(value, dict) else None


def _status(resource: dict[str, Any]) -> str:
    status = resource.get("Status") or {}
    return str(status.get("HealthRollup") or status.get("Health") or "Unknown")


def _safe_collection(
    client: RedfishClient,
    uri: str | None,
    *,
    limit: int | None = None,
) -> tuple[list[dict[str, Any]], str | None]:
    if not uri:
        return [], "Resource is not advertised by this system."
    try:
        return list(client.members(uri, limit=limit)), None
    except RedfishError as error:
        return [], str(error)


class HealthCollector:
    def __init__(self, client: RedfishClient, *, max_log_entries: int = 100) -> None:
        self.client = client
        self.max_log_entries = max_log_entries

    def collect(self) -> dict[str, Any]:
        root = self.client.get("/redfish/v1/")
        systems, system_error = _safe_collection(self.client, _link(root, "Systems"))
        chassis, chassis_error = _safe_collection(self.client, _link(root, "Chassis"))
        managers, manager_error = _safe_collection(self.client, _link(root, "Managers"))
        if not systems:
            raise RedfishError(system_error or "No ComputerSystem resource was found.")

        system = systems[0]
        chassis_item = chassis[0] if chassis else {}
        report: dict[str, Any] = {
            "generated_at": datetime.now(UTC).isoformat(timespec="seconds"),
            "target": self.client.base_url.rstrip("/"),
            "server_status": self._server_status(system),
            "temperatures": [],
            "fans": [],
            "power_supplies": [],
            "storage": [],
            "memory": [],
            "processors": [],
            "firmware": [],
            "event_logs": [],
            "collection_notes": [],
        }
        if chassis_error:
            report["collection_notes"].append(chassis_error)
        if manager_error:
            report["collection_notes"].append(manager_error)

        self._collect_thermal_and_power(chassis_item, report)
        self._collect_storage(system, report)
        self._collect_memory(system, report)
        self._collect_processors(system, report)
        self._collect_firmware(root, report)
        self._collect_logs(system, report)
        for manager in managers:
            self._collect_logs(manager, report)
        return report

    @staticmethod
    def _server_status(system: dict[str, Any]) -> dict[str, Any]:
        return {
            "Name": system.get("HostName") or system.get("Name") or "Unknown",
            "Model": system.get("Model", "Unknown"),
            "Manufacturer": system.get("Manufacturer", "Unknown"),
            "Serial number": system.get("SerialNumber", "Unknown"),
            "Power state": system.get("PowerState", "Unknown"),
            "BIOS version": system.get("BiosVersion", "Unknown"),
            "Health": _status(system),
        }

    def _collect_thermal_and_power(
        self, chassis: dict[str, Any], report: dict[str, Any]
    ) -> None:
        chassis_uri = chassis.get("@odata.id")
        thermal_uri = _link(chassis, "Thermal")
        power_uri = _link(chassis, "Power")
        if chassis_uri:
            thermal_uri = thermal_uri or f"{chassis_uri.rstrip('/')}/Thermal"
            power_uri = power_uri or f"{chassis_uri.rstrip('/')}/Power"

        try:
            thermal = self.client.get(thermal_uri) if thermal_uri else {}
            report["temperatures"] = [
                {
                    "Name": item.get("Name", "Unknown"),
                    "Reading (°C)": item.get("ReadingCelsius", "N/A"),
                    "Upper critical (°C)": item.get("UpperThresholdCritical", "N/A"),
                    "Health": _status(item),
                    "State": (item.get("Status") or {}).get("State", "Unknown"),
                }
                for item in thermal.get("Temperatures", [])
            ]
            report["fans"] = [
                {
                    "Name": item.get("Name", "Unknown"),
                    "Reading": item.get("Reading", "N/A"),
                    "Units": item.get("ReadingUnits", "N/A"),
                    "Health": _status(item),
                    "State": (item.get("Status") or {}).get("State", "Unknown"),
                }
                for item in thermal.get("Fans", [])
            ]
        except RedfishError as error:
            report["collection_notes"].append(str(error))

        try:
            power = self.client.get(power_uri) if power_uri else {}
            report["power_supplies"] = [
                {
                    "Name": item.get("Name", "Unknown"),
                    "Model": item.get("Model", "Unknown"),
                    "Serial number": item.get("SerialNumber", "Unknown"),
                    "Capacity (W)": item.get("PowerCapacityWatts", "N/A"),
                    "Health": _status(item),
                    "State": (item.get("Status") or {}).get("State", "Unknown"),
                }
                for item in power.get("PowerSupplies", [])
            ]
        except RedfishError as error:
            report["collection_notes"].append(str(error))

    def _collect_storage(self, system: dict[str, Any], report: dict[str, Any]) -> None:
        resources, error = _safe_collection(self.client, _link(system, "Storage"))
        if error:
            report["collection_notes"].append(error)
        for storage in resources:
            base = {
                "Name": storage.get("Name", storage.get("Id", "Unknown")),
                "Description": storage.get("Description", ""),
                "Health": _status(storage),
                "State": (storage.get("Status") or {}).get("State", "Unknown"),
            }
            controllers = storage.get("StorageControllers") or []
            if controllers:
                base["Controllers"] = "; ".join(
                    str(item.get("Model") or item.get("Name") or "Unknown")
                    for item in controllers
                )
            report["storage"].append(base)

            for relation, label in (("Drives", "Drive"), ("Volumes", "Volume")):
                linked = storage.get(relation)
                if isinstance(linked, list):
                    uris = [item.get("@odata.id") for item in linked if isinstance(item, dict)]
                    children = []
                    for uri in uris:
                        try:
                            children.append(self.client.get(uri))
                        except RedfishError as child_error:
                            report["collection_notes"].append(str(child_error))
                else:
                    children, child_error = _safe_collection(
                        self.client, _link(storage, relation)
                    )
                    if child_error and _link(storage, relation):
                        report["collection_notes"].append(child_error)
                for item in children:
                    report["storage"].append(
                        {
                            "Name": f"{label}: {item.get('Name', item.get('Id', 'Unknown'))}",
                            "Description": item.get("Model") or item.get("VolumeType") or "",
                            "Health": _status(item),
                            "State": (item.get("Status") or {}).get("State", "Unknown"),
                        }
                    )

    def _collect_memory(self, system: dict[str, Any], report: dict[str, Any]) -> None:
        resources, error = _safe_collection(self.client, _link(system, "Memory"))
        if error:
            report["collection_notes"].append(error)
        report["memory"] = [
            {
                "Name": item.get("DeviceLocator") or item.get("Name", "Unknown"),
                "Capacity (MiB)": item.get("CapacityMiB", "N/A"),
                "Type": item.get("MemoryDeviceType", "Unknown"),
                "Speed (MHz)": item.get("OperatingSpeedMhz", "N/A"),
                "Health": _status(item),
                "State": (item.get("Status") or {}).get("State", "Unknown"),
            }
            for item in resources
        ]

    def _collect_processors(self, system: dict[str, Any], report: dict[str, Any]) -> None:
        resources, error = _safe_collection(self.client, _link(system, "Processors"))
        if error:
            report["collection_notes"].append(error)
        report["processors"] = [
            {
                "Name": item.get("Socket") or item.get("Name", "Unknown"),
                "Model": item.get("Model", "Unknown"),
                "Cores": item.get("TotalCores", "N/A"),
                "Threads": item.get("TotalThreads", "N/A"),
                "Health": _status(item),
                "State": (item.get("Status") or {}).get("State", "Unknown"),
            }
            for item in resources
        ]

    def _collect_firmware(self, root: dict[str, Any], report: dict[str, Any]) -> None:
        update_uri = _link(root, "UpdateService") or "/redfish/v1/UpdateService"
        try:
            update_service = self.client.get(update_uri)
            inventory_uri = _link(update_service, "FirmwareInventory")
            resources, error = _safe_collection(self.client, inventory_uri)
            if error:
                report["collection_notes"].append(error)
            report["firmware"] = [
                {
                    "Name": item.get("Name", item.get("Id", "Unknown")),
                    "Version": item.get("Version", "Unknown"),
                    "Updateable": item.get("Updateable", "Unknown"),
                    "Health": _status(item),
                    "State": (item.get("Status") or {}).get("State", "Unknown"),
                }
                for item in resources
            ]
        except RedfishError as error:
            report["collection_notes"].append(str(error))

    def _collect_logs(self, system: dict[str, Any], report: dict[str, Any]) -> None:
        services, error = _safe_collection(self.client, _link(system, "LogServices"))
        if error:
            report["collection_notes"].append(error)
        for service in services:
            entries_uri = _link(service, "Entries")
            entries, entries_error = _safe_collection(
                self.client, entries_uri, limit=self.max_log_entries
            )
            if entries_error:
                report["collection_notes"].append(entries_error)
            service_name = service.get("Name", service.get("Id", "Event log"))
            for item in entries:
                hpe = (item.get("Oem") or {}).get("Hpe") or {}
                report["event_logs"].append(
                    {
                        "Log": service_name,
                        "Created": item.get("Created", "Unknown"),
                        "Severity": item.get("Severity") or hpe.get("Severity") or "Unknown",
                        "Message": item.get("Message", ""),
                        "Repaired": hpe.get("Repaired", "N/A"),
                    }
                )
