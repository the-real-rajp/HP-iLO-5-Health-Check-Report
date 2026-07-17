"""Small read-only Redfish client with session authentication."""

from __future__ import annotations

from collections.abc import Iterator
from typing import Any
from urllib.parse import urljoin

import requests


class RedfishError(RuntimeError):
    """Raised when a Redfish request cannot be completed."""


class RedfishClient:
    def __init__(
        self,
        host: str,
        username: str,
        password: str,
        *,
        verify: bool = True,
        timeout: float = 30.0,
    ) -> None:
        normalized = host.strip().rstrip("/")
        if not normalized:
            raise ValueError("An iLO IP address or FQDN is required.")
        if "://" not in normalized:
            normalized = f"https://{normalized}"
        if not normalized.lower().startswith("https://"):
            raise ValueError("Only HTTPS iLO endpoints are supported.")

        self.base_url = normalized + "/"
        self.username = username
        self.password = password
        self.verify = verify
        self.timeout = timeout
        self.session = requests.Session()
        self.session.headers.update(
            {"Accept": "application/json", "OData-Version": "4.0"}
        )
        self.session_uri: str | None = None

    def __enter__(self) -> "RedfishClient":
        self.login()
        return self

    def __exit__(self, exc_type: object, exc: object, traceback: object) -> None:
        self.logout()

    def _url(self, uri: str) -> str:
        return urljoin(self.base_url, uri)

    def login(self) -> None:
        response = self.session.post(
            self._url("/redfish/v1/SessionService/Sessions"),
            json={"UserName": self.username, "Password": self.password},
            verify=self.verify,
            timeout=self.timeout,
        )
        if response.status_code not in (200, 201):
            raise RedfishError(
                f"Redfish login failed with HTTP {response.status_code}."
            )
        token = response.headers.get("X-Auth-Token")
        if not token:
            raise RedfishError("iLO did not return a Redfish session token.")
        self.session.headers["X-Auth-Token"] = token
        self.session_uri = response.headers.get("Location")

    def logout(self) -> None:
        if not self.session_uri:
            return
        try:
            self.session.delete(
                self._url(self.session_uri),
                verify=self.verify,
                timeout=self.timeout,
            )
        finally:
            self.session_uri = None
            self.session.headers.pop("X-Auth-Token", None)

    def get(self, uri: str) -> dict[str, Any]:
        response = self.session.get(
            self._url(uri), verify=self.verify, timeout=self.timeout
        )
        if not response.ok:
            raise RedfishError(f"GET {uri} failed with HTTP {response.status_code}.")
        try:
            payload = response.json()
        except requests.JSONDecodeError as error:
            raise RedfishError(f"GET {uri} returned invalid JSON.") from error
        if not isinstance(payload, dict):
            raise RedfishError(f"GET {uri} did not return a JSON object.")
        return payload

    def members(self, uri: str, *, limit: int | None = None) -> Iterator[dict[str, Any]]:
        """Yield collection members, following Redfish pagination links."""
        next_uri: str | None = uri
        yielded = 0
        while next_uri:
            collection = self.get(next_uri)
            for member in collection.get("Members", []):
                if limit is not None and yielded >= limit:
                    return
                if isinstance(member, dict) and member.get("@odata.id"):
                    yield self.get(member["@odata.id"])
                elif isinstance(member, dict):
                    yield member
                yielded += 1
            next_uri = collection.get("Members@odata.nextLink")

