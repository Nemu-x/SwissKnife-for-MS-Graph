from typing import Any, Dict, List

from .graph_client import GraphClient


def list_groups(client: GraphClient, top: int = 25) -> List[Dict[str, Any]]:
    """
    Список первых top групп (AAD / M365 Groups / Teams backing groups).
    """
    params = {"$top": top}
    result = client.get("/groups", params=params)
    return result.get("value", [])


def get_group(client: GraphClient, group_id: str) -> Dict[str, Any]:
    """
    Получить одну группу по id.
    """
    path = f"/groups/{group_id}"
    return client.get(path)


def list_group_members(client: GraphClient, group_id: str) -> List[Dict[str, Any]]:
    """
    Список участников группы (users / groups / service principals).
    """
    path = f"/groups/{group_id}/members"
    result = client.get(path)
    return result.get("value", [])


def add_group_owner_by_upn(
    client: GraphClient,
    group_id: str,
    user_upn: str,
) -> Dict[str, Any]:
    """
    Добавить owner'а в группу по UPN.

    1) Получаем пользователя по /users/{user_upn}
    2) Добавляем в /groups/{group_id}/owners/$ref

    Требуются права: Group.ReadWrite.All + User.Read.All / Directory.Read.All
    """
    user = client.get(f"/users/{user_upn}")
    user_id = user.get("id")
    if not user_id:
        raise RuntimeError(f"Не найден пользователь {user_upn}")

    body = {
        "@odata.id": f"https://graph.microsoft.com/v1.0/users/{user_id}"
    }

    path = f"/groups/{group_id}/owners/$ref"
    result = client.post(path, json=body)
    return result
