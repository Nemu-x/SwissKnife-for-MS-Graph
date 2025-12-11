from typing import Any, Dict, List

from .graph_client import GraphClient


def list_skus(client: GraphClient) -> List[Dict[str, Any]]:
    """
    Список подписок (SKU) тенанта.
    GET /subscribedSkus
    Требуются Directory.Read.All / Organization.Read.All (Application).
    """
    result = client.get("/subscribedSkus")
    return result.get("value", [])


def assign_licenses(
    client: GraphClient,
    user: str,
    add_sku_ids: List[str],
    remove_sku_ids: List[str],
) -> Dict[str, Any]:
    """
    Выдать / снять лицензии пользователя.
    POST /users/{user}/assignLicense

    add_sku_ids  — список GUID SKU, которые нужно ДОБАВИТЬ.
    remove_sku_ids — список GUID SKU, которые нужно СНЯТЬ.
    """
    body = {
        "addLicenses": [{"skuId": sku} for sku in add_sku_ids],
        "removeLicenses": remove_sku_ids,
    }

    path = f"/users/{user}/assignLicense"
    return client.post(path, json=body)
