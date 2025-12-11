from typing import Any, Dict, List
from .graph_client import GraphClient


def audit_logs(client: GraphClient, top: int = 50) -> List[Dict[str, Any]]:
    """
    Читает auditLogs/directoryAudits
    Требуются: AuditLog.Read.All (Application)
    """
    result = client.get("/auditLogs/directoryAudits", params={"$top": top})
    return result.get("value", [])


def sign_in_logs(client: GraphClient, top: int = 50) -> List[Dict[str, Any]]:
    """
    Читает auditLogs/signIns
    Требуются: AuditLog.Read.All (Application)
    """
    result = client.get("/auditLogs/signIns", params={"$top": top})
    return result.get("value", [])
