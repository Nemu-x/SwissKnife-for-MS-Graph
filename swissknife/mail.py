from typing import Any, Dict, List

from .graph_client import GraphClient


def list_messages(
    client: GraphClient,
    user: str,
    top: int = 25,
    folder: str = "inbox",
) -> List[Dict[str, Any]]:
    """
    Список писем из папки пользователя.
    user  — UPN или id.
    folder — системная папка: inbox, sentitems, drafts и т.п.
    """
    path = f"/users/{user}/mailFolders/{folder}/messages"
    params = {
        "$top": top,
        "$select": "id,subject,from,receivedDateTime,isRead",
        "$orderby": "receivedDateTime DESC",
    }
    result = client.get(path, params=params)
    return result.get("value", [])


def send_mail(
    client: GraphClient,
    user: str,
    subject: str,
    body_text: str,
    to_recipients: List[str],
) -> None:
    """
    Отправить письмо от имени пользователя.
    Требуется Mail.Send (Application).
    """
    message = {
        "subject": subject,
        "body": {
            "contentType": "Text",
            "content": body_text,
        },
        "toRecipients": [
            {"emailAddress": {"address": addr}} for addr in to_recipients
        ],
    }

    payload = {
        "message": message,
        "saveToSentItems": True,
    }

    path = f"/users/{user}/sendMail"
    client.post(path, json=payload)
