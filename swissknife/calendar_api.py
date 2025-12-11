from typing import Any, Dict, List

from .graph_client import GraphClient


def list_events(
    client: GraphClient,
    user: str,
    top: int = 25,
) -> List[Dict[str, Any]]:
    """
    Список событий пользователя.
    Простейший вариант: /users/{id}/events, последние top.
    Требуется Calendars.Read или Calendars.ReadBasic.All (Application).
    """
    path = f"/users/{user}/events"
    params = {
        "$top": top,
        "$orderby": "start/dateTime DESC",
        "$select": "id,subject,start,end,location,organizer",
    }
    result = client.get(path, params=params)
    return result.get("value", [])


def create_event(
    client: GraphClient,
    user: str,
    subject: str,
    body_text: str,
    start_iso: str,
    end_iso: str,
    timezone: str,
    attendees: List[str],
) -> Dict[str, Any]:
    """
    Создать событие в календаре пользователя.
    Даты в ISO-формате: '2025-12-11T10:00:00'
    timezone, например: 'UTC' или 'Russian Standard Time' и т.п.
    Требуется Calendars.ReadWrite (Application).
    """
    event = {
        "subject": subject,
        "body": {
            "contentType": "Text",
            "content": body_text,
        },
        "start": {
            "dateTime": start_iso,
            "timeZone": timezone,
        },
        "end": {
            "dateTime": end_iso,
            "timeZone": timezone,
        },
        "attendees": [
            {
                "emailAddress": {"address": addr},
                "type": "required",
            }
            for addr in attendees
        ],
    }

    path = f"/users/{user}/events"
    return client.post(path, json=event)
