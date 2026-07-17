from ilo_health_report.collector import HealthCollector


class FakeClient:
    base_url = "https://ilo.example.com/"

    resources = {
        "/redfish/v1/": {
            "Systems": {"@odata.id": "/systems"},
            "Chassis": {"@odata.id": "/chassis"},
            "Managers": {"@odata.id": "/managers"},
            "UpdateService": {"@odata.id": "/update"},
        },
        "/systems": {"Members": [{"@odata.id": "/systems/1"}]},
        "/systems/1": {
            "@odata.id": "/systems/1",
            "Name": "Server",
            "Status": {"Health": "OK"},
            "Memory": {"@odata.id": "/memory"},
            "Processors": {"@odata.id": "/processors"},
            "Storage": {"@odata.id": "/storage"},
            "LogServices": {"@odata.id": "/logs"},
        },
        "/chassis": {"Members": [{"@odata.id": "/chassis/1"}]},
        "/managers": {"Members": []},
        "/chassis/1": {"@odata.id": "/chassis/1"},
        "/chassis/1/Thermal": {"Temperatures": [], "Fans": []},
        "/chassis/1/Power": {"PowerSupplies": []},
        "/memory": {"Members": []},
        "/processors": {"Members": []},
        "/storage": {"Members": []},
        "/logs": {"Members": []},
        "/update": {"FirmwareInventory": {"@odata.id": "/firmware"}},
        "/firmware": {"Members": []},
    }

    def get(self, uri):
        return self.resources[uri]

    def members(self, uri, *, limit=None):
        members = self.get(uri).get("Members", [])
        count = 0
        for member in members:
            if limit is not None and count >= limit:
                return
            yield self.get(member["@odata.id"])
            count += 1


def test_collector_returns_requested_sections() -> None:
    data = HealthCollector(FakeClient()).collect()
    assert data["server_status"]["Health"] == "OK"
    for key in (
        "temperatures",
        "fans",
        "power_supplies",
        "storage",
        "memory",
        "processors",
        "firmware",
        "event_logs",
    ):
        assert key in data
