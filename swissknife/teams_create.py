from typing import Any, Dict, Optional

from .graph_client import GraphClient


def create_m365_group(
    client: GraphClient,
    display_name: str,
    description: str,
    mail_nickname: str,
    owner_upn: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Создать M365 группу (Unified group).

    Если указан owner_upn — сразу добавляем его как owner и member группы.
    Требуются Group.ReadWrite.All + User.Read.All / Directory.Read.All.
    """
    body: Dict[str, Any] = {
        "displayName": display_name,
        "description": description,
        "groupTypes": ["Unified"],
        "mailEnabled": True,
        "securityEnabled": False,
        "mailNickname": mail_nickname,
    }

    if owner_upn:
        user_url = f"https://graph.microsoft.com/v1.0/users('{owner_upn}')"
        body["owners@odata.bind"] = [user_url]
        body["members@odata.bind"] = [user_url]

    return client.post("/groups", json=body)


def create_team_from_group(
    client: GraphClient,
    group_id: str,
) -> Dict[str, Any]:
    """
    Превратить M365 группу в Team.
    POST /groups/{id}/team
    Требуются Team.Create (Application).
    """
    body = {
        "memberSettings": {
            "allowCreateUpdateChannels": True
        },
        "messagingSettings": {
            "allowUserEditMessages": True,
            "allowUserDeleteMessages": True
        },
        "funSettings": {
            "allowGiphy": True,
            "giphyContentRating": "strict"
        }
    }

    return client.put(f"/groups/{group_id}/team", json=body)


def create_channel(
    client: GraphClient,
    team_id: str,
    display_name: str,
    description: str,
    channel_type: str = "standard",   # standard | private | shared
    owner_upn: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Создать канал в Teams.

    standard — обычный канал (без members, все члены команды его видят).
    private / shared — при app-only нужно указать owner_upn,
    Graph требует, чтобы при создании был ровно один owner.
    """
    body: Dict[str, Any] = {
        "displayName": display_name,
        "description": description,
        "membershipType": channel_type,  # standard / private / shared
    }

    if channel_type in ("private", "shared"):
        if not owner_upn:
            raise RuntimeError(
                "Для private/shared канала нужно указать --owner (UPN владельца)."
            )

        member = {
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            "roles": ["owner"],
            "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{owner_upn}')",
        }

        body["members"] = [member]

    return client.post(f"/teams/{team_id}/channels", json=body)
