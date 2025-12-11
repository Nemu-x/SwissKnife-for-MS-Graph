import msal
from typing import List, Optional

GRAPH_AUTHORITY_TEMPLATE = "https://login.microsoftonline.com/{tenant_id}"


def get_confidential_client(
    tenant_id: str,
    client_id: str,
    client_secret: str,
) -> msal.ConfidentialClientApplication:
    """
    Создаёт MSAL ConfidentialClientApplication для client credentials flow.
    """
    authority = GRAPH_AUTHORITY_TEMPLATE.format(tenant_id=tenant_id)
    app = msal.ConfidentialClientApplication(
        client_id=client_id,
        client_credential=client_secret,
        authority=authority,
    )
    return app


def acquire_token_client_credentials(
    tenant_id: str,
    client_id: str,
    client_secret: str,
    scopes: Optional[List[str]] = None,
) -> str:
    """
    Получает access token для приложения (app-only).
    По умолчанию scope: https://graph.microsoft.com/.default
    """
    if scopes is None:
        scopes = ["https://graph.microsoft.com/.default"]

    app = get_confidential_client(tenant_id, client_id, client_secret)

    # Попытка взять токен из кэша (на время жизни процесса)
    result = app.acquire_token_silent(scopes, account=None)

    if not result:
        # Берём новый токен
        result = app.acquire_token_for_client(scopes=scopes)

    if "access_token" not in result:
        raise RuntimeError(
            f"Не удалось получить токен: {result.get('error')}: {result.get('error_description')}"
        )

    return result["access_token"]
