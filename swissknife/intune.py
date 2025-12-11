from typing import Any, Dict, List

from .graph_client import GraphClient


def list_managed_devices(
    client: GraphClient,
    top: int = 50,
) -> List[Dict[str, Any]]:
    """
    Список управляемых устройств Intune.
    GET /deviceManagement/managedDevices
    Требуются: DeviceManagementManagedDevices.Read.All (или ReadWrite.All).
    """
    params = {"$top": top}
    result = client.get("/deviceManagement/managedDevices", params=params)
    return result.get("value", [])


def get_managed_device(
    client: GraphClient,
    device_id: str,
) -> Dict[str, Any]:
    """
    Информация об одном устройстве.
    """
    path = f"/deviceManagement/managedDevices/{device_id}"
    return client.get(path)


def wipe_device(
    client: GraphClient,
    device_id: str,
    keep_enrollment_data: bool = False,
    keep_user_data: bool = False,
) -> Any:
    """
    Wipe устройства (полный/частичный).
    POST /deviceManagement/managedDevices/{id}/wipe
    """
    path = f"/deviceManagement/managedDevices/{device_id}/wipe"
    body = {
        "keepEnrollmentData": keep_enrollment_data,
        "keepUserData": keep_user_data,
        "macOsUnlockCode": None,
    }
    return client.post(path, json=body)


def retire_device(
    client: GraphClient,
    device_id: str,
) -> Any:
    """
    Retire устройства.
    POST /deviceManagement/managedDevices/{id}/retire
    """
    path = f"/deviceManagement/managedDevices/{device_id}/retire"
    return client.post(path, json={})


def remote_lock_device(
    client: GraphClient,
    device_id: str,
) -> Any:
    """
    Remote lock устройства.
    POST /deviceManagement/managedDevices/{id}/remoteLock
    """
    path = f"/deviceManagement/managedDevices/{device_id}/remoteLock"
    return client.post(path, json={})
