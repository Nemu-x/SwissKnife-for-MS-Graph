from typing import Any, Dict, List, Optional
import os
from pathlib import Path

import requests

from .graph_client import GraphClient, GRAPH_BASE_URL


def list_root(
    client: GraphClient,
    user: str,
    top: int = 50,
) -> List[Dict[str, Any]]:
    """
    Список файлов/папок в корне OneDrive пользователя.
    GET /users/{user}/drive/root/children
    """
    path = f"/users/{user}/drive/root/children"
    params = {"$top": top}
    result = client.get(path, params=params)
    return result.get("value", [])


def list_children(
    client: GraphClient,
    user: str,
    item_id: str,
    top: int = 50,
) -> List[Dict[str, Any]]:
    """
    Список детей конкретной папки/элемента.
    GET /users/{user}/drive/items/{item-id}/children
    """
    path = f"/users/{user}/drive/items/{item_id}/children"
    params = {"$top": top}
    result = client.get(path, params=params)
    return result.get("value", [])


def search_files(
    client: GraphClient,
    user: str,
    query: str,
    top: int = 25,
) -> List[Dict[str, Any]]:
    """
    Поиск файлов в OneDrive.
    GET /users/{user}/drive/root/search(q='{query}')
    """
    path = f"/users/{user}/drive/root/search(q='{query}')"
    params = {"$top": top}
    result = client.get(path, params=params)
    return result.get("value", [])


def _auth_headers(client: GraphClient) -> Dict[str, str]:
    return {
        "Authorization": f"Bearer {client.access_token}",
    }


def download_item(
    client: GraphClient,
    user: str,
    item_id: str,
    local_path: str,
) -> None:
    """
    Скачивает файл по item_id во временную/локальную папку.
    GET /users/{user}/drive/items/{item-id}/content
    """
    url = f"{GRAPH_BASE_URL}/users/{user}/drive/items/{item_id}/content"
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


def upload_file_to_path(
    client: GraphClient,
    user: str,
    local_path: str,
    remote_path: str,
    overwrite: bool = True,
) -> Dict[str, Any]:
    """
    Загружает файл в OneDrive по указанному пути.
    PUT /users/{user}/drive/root:/{remote_path}:/content
    """
    url = f"{GRAPH_BASE_URL}/users/{user}/drive/root:/{remote_path}:/content"

    headers = _auth_headers(client)
    # overwrite зависит от поведения Graph, но как правило последний wins

    with open(local_path, "rb") as f:
        resp = requests.put(url, headers=headers, data=f)

    try:
        data = resp.json()
    except ValueError:
        data = resp.text

    if not resp.ok:
        raise RuntimeError(f"Upload failed: {resp.status_code} {data}")

    return data


def delete_item(
    client: GraphClient,
    user: str,
    item_id: str,
) -> None:
    """
    Удаляет файл/папку.
    DELETE /users/{user}/drive/items/{item-id}
    """
    path = f"/users/{user}/drive/items/{item_id}"
    client.delete(path)


def create_link(
    client: GraphClient,
    user: str,
    item_id: str,
    link_type: str = "view",      # view | edit | embed
    scope: str = "organization",  # anonymous | organization
) -> Dict[str, Any]:
    """
    Создаёт шаринг-ссылку.
    POST /users/{user}/drive/items/{item-id}/createLink
    """
    body = {
        "type": link_type,
        "scope": scope,
    }
    path = f"/users/{user}/drive/items/{item_id}/createLink"
    return client.post(path, json=body)


def clone_root(
    client: GraphClient,
    source_user: str,
    target_user: str,
    overwrite: bool = False,
    tmp_dir: str = "/tmp/swissknife_onedrive_clone",
) -> Dict[str, Any]:
    """
    Клонирует ВСЕ файлы из корня OneDrive source_user в корень target_user.
    ВНИМАНИЕ: пока только верхний уровень (без рекурсии по подпапкам).
    """
    result: Dict[str, Any] = {
        "copied": [],
        "skipped": [],
        "failed": [],
    }

    os.makedirs(tmp_dir, exist_ok=True)

    # список файлов источника
    src_items = list_root(client, source_user, top=999)

    # список файлов назначения (для проверки overwrite)
    tgt_items = list_root(client, target_user, top=999)
    tgt_names = {item.get("name", ""): item for item in tgt_items}

    for item in src_items:
        name = item.get("name", "")
        if not name:
            continue

        # папки пока пропускаем (MVP)
        if "folder" in item:
            result["skipped"].append(
                {"name": name, "reason": "folder_not_supported_yet"}
            )
            continue

        if (not overwrite) and name in tgt_names:
            result["skipped"].append(
                {"name": name, "reason": "exists_in_target"}
            )
            continue

        item_id = item.get("id")
        if not item_id:
            continue

        local_file = os.path.join(tmp_dir, name)

        try:
            download_item(client, source_user, item_id, local_file)
            upload_file_to_path(client, target_user, local_file, name)
            result["copied"].append({"name": name})
        except Exception as e:
            result["failed"].append({"name": name, "error": str(e)})

    return result
