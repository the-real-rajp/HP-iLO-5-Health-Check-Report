from unittest.mock import Mock

from ilo_health_report.redfish import RedfishClient


def test_host_is_normalized_to_https() -> None:
    client = RedfishClient("ilo.example.com", "user", "secret")
    assert client.base_url == "https://ilo.example.com/"


def test_members_follows_pagination() -> None:
    client = RedfishClient("ilo.example.com", "user", "secret")
    client.get = Mock(
        side_effect=[
            {"Members": [{"@odata.id": "/item/1"}], "Members@odata.nextLink": "/page/2"},
            {"Id": "1"},
            {"Members": [{"@odata.id": "/item/2"}]},
            {"Id": "2"},
        ]
    )
    assert [item["Id"] for item in client.members("/collection")] == ["1", "2"]


def test_members_honors_limit() -> None:
    client = RedfishClient("ilo.example.com", "user", "secret")
    client.get = Mock(
        side_effect=[
            {"Members": [{"@odata.id": "/item/1"}, {"@odata.id": "/item/2"}]},
            {"Id": "1"},
        ]
    )
    assert [item["Id"] for item in client.members("/collection", limit=1)] == ["1"]

