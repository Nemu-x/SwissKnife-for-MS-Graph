from typing import Any, Dict, List

from .graph_client import GraphClient


def list_user_chats(client: GraphClient, user: str) -> List[Dict[str, Any]]:
    """
    Возвращает список чатов пользователя.
    user — UPN или id пользователя (например, user1@example.com).
    """
    path = f"/users/{user}/chats"
    result = client.get(path)
    return result.get("value", [])


def add_user_to_chat(
    client: GraphClient,
    chat_id: str,
    user_upn: str,
    as_owner: bool = False,
) -> Dict[str, Any]:
    """
    Добавляет пользователя в чат по chat_id.
    user_upn — UPN пользователя (user1@example.com).
    as_owner — добавить как owner (True) или обычный участник (False).
    Требуются соответствующие Chat.* Application-permissions.
    """
    roles: List[str] = ["owner"] if as_owner else []

    body = {
        "@odata.type": "#microsoft.graph.aadUserConversationMember",
        "roles": roles,
        "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{user_upn}')",
    }

    path = f"/chats/{chat_id}/members"
    result = client.post(path, json=body)
    return result


def get_chat_messages(
    client: GraphClient,
    chat_id: str,
    top: int = 50,
) -> List[Dict[str, Any]]:
    """
    Получает последние top сообщений из чата.
    Требуется Chat.Read.All или Chat.ReadWrite.All (Application).
    """
    params = {"$top": top}
    path = f"/chats/{chat_id}/messages"
    result = client.get(path, params=params)
    return result.get("value", [])


def list_chat_members(
    client: GraphClient,
    chat_id: str,
) -> List[Dict[str, Any]]:
    """
    Список участников чата.
    GET /chats/{chat-id}/members
    """
    path = f"/chats/{chat_id}/members"
    result = client.get(path)
    return result.get("value", [])


def remove_user_from_chat(
    client: GraphClient,
    chat_id: str,
    user_upn: str,
) -> None:
    """
    Удаляет пользователя из чата по UPN.

    Внутри:
      - получаем список участников,
      - ищем по email==user_upn (без регистра),
      - удаляем по membership-id.
    """
    members = list_chat_members(client, chat_id)
    target_member_id = None

    upn_lower = user_upn.lower()

    for m in members:
        email = (m.get("email") or "").lower()
        if email == upn_lower:
            target_member_id = m.get("id")
            break

    if not target_member_id:
        raise RuntimeError(
            f"Не найден участник с UPN/email {user_upn} в чате {chat_id}"
        )

    path = f"/chats/{chat_id}/members/{target_member_id}"
    client.delete(path)
