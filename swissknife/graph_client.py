from typing import Any, Dict, Optional, Union

import requests

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"


class GraphClient:
    """
    Примитивный клиент для Microsoft Graph.
    """

    def __init__(self, access_token: str, base_url: str = GRAPH_BASE_URL) -> None:
        self.access_token = access_token
        self.base_url = base_url.rstrip("/")

    def _make_url(self, path: str) -> str:
        if path.startswith("http://") or path.startswith("https://"):
            return path
        path = path.lstrip("/")
        return f"{self.base_url}/{path}"

    def request(
        self,
        method: str,
        path: str,
        params: Optional[Dict[str, Any]] = None,
        json: Optional[Dict[str, Any]] = None,
    ) -> Union[Dict[str, Any], str]:
        url = self._make_url(path)
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Accept": "application/json",
        }
        if json is not None:
            headers["Content-Type"] = "application/json"

        resp = requests.request(
            method=method.upper(),
            url=url,
            headers=headers,
            params=params,
            json=json,
        )

        try:
            data = resp.json()
        except ValueError:
            data = resp.text

        if not resp.ok:
            raise RuntimeError(f"Graph API error {resp.status_code}: {data}")

        return data

    def get(self, path: str, params: Optional[Dict[str, Any]] = None) -> Any:
        return self.request("GET", path, params=params)

    def post(self, path: str, json: Optional[Dict[str, Any]] = None) -> Any:
        return self.request("POST", path, json=json)

    def delete(self, path: str) -> Any:
        return self.request("DELETE", path)

    def patch(self, path: str, json: Optional[Dict[str, Any]] = None) -> Any:
        return self.request("PATCH", path, json=json)

    def put(self, path: str, json: Optional[Dict[str, Any]] = None) -> Any:
        """
        PUT-запрос (нужен, например, для /groups/{id}/team).
        """
        return self.request("PUT", path, json=json)
