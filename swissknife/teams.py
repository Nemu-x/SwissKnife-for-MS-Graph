from typing import Any, Dict, List

from .graph_client import GraphClient


def list_user_joined_teams(
    client: GraphClient,
    user: str,
) -> List[Dict[str, Any]]:
    """
    Список команд (Teams), в которых состоит пользователь.

    user — UPN или id пользователя.
    GET /users/{id}/joinedTeams
    """
    path = f"/users/{user}/joinedTeams"
    result = client.get(path)
    return result.get("value", [])


def list_team_channels(
    client: GraphClient,
    team_id: str,
) -> List[Dict[str, Any]]:
    """
    Список каналов в команде.
    GET /teams/{team-id}/channels
    """
    path = f"/teams/{team_id}/channels"
    result = client.get(path)
    return result.get("value", [])


def add_member_to_team(
    client: GraphClient,
    team_id: str,
    user_upn: str,
    as_owner: bool = False,
) -> Dict[str, Any]:
    """
    Добавляет пользователя в команду (Team).
    POST /teams/{team-id}/members
    """

    roles: List[str] = ["owner"] if as_owner else []

    body = {
        "@odata.type": "#microsoft.graph.aadUserConversationMember",
        "roles": roles,
        "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{user_upn}')",
    }

    path = f"/teams/{team_id}/members"
    result = client.post(path, json=body)
    return result


def add_member_to_channel(
    client: GraphClient,
    team_id: str,
    channel_id: str,
    user_upn: str,
    as_owner: bool = False,
) -> Dict[str, Any]:
    """
    Добавляет пользователя в канал.

    POST /teams/{team-id}/channels/{channel-id}/members
    """

    roles: List[str] = ["owner"] if as_owner else []

    body = {
        "@odata.type": "#microsoft.graph.aadUserConversationMember",
        "roles": roles,
        "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{user_upn}')",
    }

    path = f"/teams/{team_id}/channels/{channel_id}/members"
    result = client.post(path, json=body)
    return result


def list_team_members(
    client: GraphClient,
    team_id: str,
) -> List[Dict[str, Any]]:
    """
    Список участников команды (Team).
    GET /teams/{team-id}/members
    """
    path = f"/teams/{team_id}/members"
    result = client.get(path)
    return result.get("value", [])


def list_channel_members(
    client: GraphClient,
    team_id: str,
    channel_id: str,
) -> List[Dict[str, Any]]:
    """
    Список участников канала.
    GET /teams/{team-id}/channels/{channel-id}/members
    """
    path = f"/teams/{team_id}/channels/{channel_id}/members"
    result = client.get(path)
    return result.get("value", [])


def remove_member_from_team(
    client: GraphClient,
    team_id: str,
    user_upn: str,
) -> None:
    """
    Удаляет пользователя из Team по UPN.
    """
    members = list_team_members(client, team_id)
    upn_lower = user_upn.lower()
    membership_id = None

    for m in members:
        email = (m.get("email") or "").lower()
        if email == upn_lower:
            membership_id = m.get("id")
            break

    if not membership_id:
        raise RuntimeError(
            f"Не найден участник с UPN/email {user_upn} в Team {team_id}"
        )

    path = f"/teams/{team_id}/members/{membership_id}"
    client.delete(path)


def remove_member_from_channel(
    client: GraphClient,
    team_id: str,
    channel_id: str,
    user_upn: str,
) -> None:
    """
    Удаляет пользователя из канала по UPN.
    """
    members = list_channel_members(client, team_id, channel_id)
    upn_lower = user_upn.lower()
    membership_id = None

    for m in members:
        email = (m.get("email") or "").lower()
        if email == upn_lower:
            membership_id = m.get("id")
            break

    if not membership_id:
        raise RuntimeError(
            f"Не найден участник с UPN/email {user_upn} в канале {channel_id}"
        )

    path = f"/teams/{team_id}/channels/{channel_id}/members/{membership_id}"
    client.delete(path)
