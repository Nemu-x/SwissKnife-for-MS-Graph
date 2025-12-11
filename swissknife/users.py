from typing import Any, Dict, List

from .graph_client import GraphClient


def list_users(client: GraphClient, top: int = 25) -> List[Dict[str, Any]]:
    """
    Возвращает список первых top пользователей.
    Требуется User.Read.All (Application).
    """
    params = {"$top": top}
    result = client.get("/users", params=params)
    return result.get("value", [])


def get_user(client: GraphClient, user: str) -> Dict[str, Any]:
    """
    Получить одного пользователя по UPN или objectId.
    """
    path = f"/users/{user}"
    return client.get(path)


def get_user_member_of(client: GraphClient, user: str) -> List[Dict[str, Any]]:
    """
    Группы/роли, в которых состоит пользователь.
    """
    path = f"/users/{user}/memberOf"
    result = client.get(path)
    return result.get("value", [])


def get_user_license_details(client: GraphClient, user: str) -> List[Dict[str, Any]]:
    """
    Лицензии пользователя.
    Требуется Directory.Read.All или чуть выше.
    """
    path = f"/users/{user}/licenseDetails"
    result = client.get(path)
    return result.get("value", [])
