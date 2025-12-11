from typing import Any, Dict

from .graph_client import GraphClient


def get_user_basic(client: GraphClient, user: str) -> Dict[str, Any]:
    """
    Базовая инфа о пользователе: id, displayName, upn, accountEnabled.
    GET /users/{user}?$select=...
    """
    path = f"/users/{user}"
    params = {
        "$select": "id,displayName,userPrincipalName,mail,accountEnabled"
    }
    return client.get(path, params=params)


def block_user(client: GraphClient, user: str) -> Dict[str, Any]:
    """
    Заблокировать пользователя (accountEnabled = false).
    PATCH /users/{user}
    Требуются Directory.ReadWrite.All или User.ReadWrite.All (Application).
    """
    path = f"/users/{user}"
    body = {"accountEnabled": False}
    return client.patch(path, json=body)


def unblock_user(client: GraphClient, user: str) -> Dict[str, Any]:
    """
    Разблокировать пользователя (accountEnabled = true).
    """
    path = f"/users/{user}"
    body = {"accountEnabled": True}
    return client.patch(path, json=body)


def reset_password(
    client: GraphClient,
    user: str,
    new_password: str,
    force_change_next_signin: bool = True,
) -> Dict[str, Any]:
    """
    Сброс пароля пользователю.
    ВАЖНО: работает только для cloud-only / управляемых в Entra ID аккаунтов.
    Для синхронизированных из локального AD может не сработать (ошибка Graph).
    """
    path = f"/users/{user}"
    body = {
        "passwordProfile": {
            "forceChangePasswordNextSignIn": force_change_next_signin,
            "password": new_password,
        }
    }
    return client.patch(path, json=body)


def revoke_sessions(client: GraphClient, user: str) -> Dict[str, Any]:
    """
    Принудительный logout / revoke refresh tokens.
    POST /users/{user}/revokeSignInSessions
    Требует Directory.AccessAsUser.All (delegated) или Directory.ReadWrite.All/Directory.Read.All в зависимости от сценария,
    но в ряде тенантов позволяет отозвать сессии и c Application токеном.
    """
    path = f"/users/{user}/revokeSignInSessions"
    return client.post(path, json={})
