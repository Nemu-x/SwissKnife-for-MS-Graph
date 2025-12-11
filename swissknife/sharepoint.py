from typing import Any, Dict, List
from pathlib import Path

import requests

from .graph_client import GraphClient, GRAPH_BASE_URL


def list_sites(
    client: GraphClient,
    search: str = "",
    top: int = 20,
) -> List[Dict[str, Any]]:
    """
    Список сайтов SharePoint.
    Если search не пустой, использует ?search=.
    """
    params: Dict[str, Any] = {"$top": top}
    if search:
        params["search"] = search

    result = client.get("/sites", params=params)
    return result.get("value", [])


def list_site_root(
    client: GraphClient,
    site_id: str,
    top: int = 50,
) -> List[Dict[str, Any]]:
    """
    Список элементов в корне drive сайта.
    GET /sites/{site-id}/drive/root/children
    """
    path = f"/sites/{site_id}/drive/root/children"
    result = client.get(path, params={"$top": top})
    return result.get("value", [])


def _auth_headers(client: GraphClient) -> Dict[str, str]:
    return {
        "Authorization": f"Bearer {client.access_token}",
    }


def download_site_item(
    client: GraphClient,
    site_id: str,
    item_id: str,
    local_path: str,
) -> None:
    """
    Скачать файл из drive сайта.
    GET /sites/{site-id}/drive/items/{item-id}/content
    """
    url = f"{GRAPH_BASE_URL}/sites/{site_id}/drive/items/{item_id}/content"
    headers = _auth_headers(client)

    resp = requests.get(url, headers=headers, stream=True)
    if not resp.ok:
        raise RuntimeError(f"Download failed: {resp.status_code} {resp.text}")

    dest = Path(local_path)
    dest.parent.mkdir(parents=True, exist_ok=True)

    with dest.open("wb") as f:
        for chunk in resp.iter_content(chunk_size=8192):
            if chunk:
                f.write(chunk)


def upload_site_file(
    client: GraphClient,
    site_id: str,
    local_path: str,
    remote_path: str,
) -> Dict[str, Any]:
    """
    Загрузка файла в drive сайта по пути.
    PUT /sites/{site-id}/drive/root:/{remote_path}:/content
    """
    url = f"{GRAPH_BASE_URL}/sites/{site_id}/drive/root:/{remote_path}:/content"
    headers = _auth_headers(client)

    with open(local_path, "rb") as f:
        resp = requests.put(url, headers=headers, data=f)

    try:
        data = resp.json()
    except ValueError:
        data = resp.text

    if not resp.ok:
        raise RuntimeError(f"Upload failed: {resp.status_code} {data}")

    return data


def delete_site_item(
    client: GraphClient,
    site_id: str,
    item_id: str,
) -> None:
    """
    Удалить элемент из drive сайта.
    """
    path = f"/sites/{site_id}/drive/items/{item_id}"
    client.delete(path)


def create_site_link(
    client: GraphClient,
    site_id: str,
    item_id: str,
    link_type: str = "view",
    scope: str = "organization",
) -> Dict[str, Any]:
    """
    Создать sharing link для элемента drive сайта.
    """
    body = {
        "type": link_type,
        "scope": scope,
    }
    path = f"/sites/{site_id}/drive/items/{item_id}/createLink"
    return client.post(path, json=body)
