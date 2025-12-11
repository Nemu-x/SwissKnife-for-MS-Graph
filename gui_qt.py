import sys
import json
import traceback
import requests
from dataclasses import dataclass
from typing import Any, Dict, List, Optional
from pathlib import Path
from PySide6.QtCore import Qt, QRegularExpression
from PySide6.QtGui import QIcon, QColor, QTextCharFormat, QSyntaxHighlighter

from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QGridLayout,
    QGroupBox,
    QLabel,
    QLineEdit,
    QPushButton,
    QTabWidget,
    QPlainTextEdit,
    QComboBox,
    QFileDialog,
    QSpinBox,
    QCheckBox,
    QTableWidget,
    QTableWidgetItem,
    QTreeWidget,
    QTreeWidgetItem,
    QSplitter,
)


GRAPH_TOKEN_URL = "https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"


@dataclass
class GraphConfig:
    tenant_id: str
    client_id: str
    client_secret: str


class GraphClient:
    def __init__(self, config: GraphConfig) -> None:
        self.config = config
        self._access_token: Optional[str] = None

    # ----- auth -----
    def authenticate(self) -> None:
        data = {
            "client_id": self.config.client_id,
            "client_secret": self.config.client_secret,
            "grant_type": "client_credentials",
            "scope": "https://graph.microsoft.com/.default",
        }
        url = GRAPH_TOKEN_URL.format(tenant_id=self.config.tenant_id)
        resp = requests.post(url, data=data, timeout=30)
        if not resp.ok:
            raise RuntimeError(f"Auth failed: {resp.status_code} {resp.text}")
        self._access_token = resp.json()["access_token"]

    # ----- low-level request -----
    def request(
        self,
        method: str,
        path_or_url: str,
        *,
        params: Optional[Dict[str, Any]] = None,
        json_data: Optional[Dict[str, Any]] = None,
    ) -> Any:
        if not self._access_token:
            raise RuntimeError("Client is not authenticated")

        if path_or_url.startswith("http"):
            url = path_or_url
        else:
            url = GRAPH_BASE_URL + path_or_url

        headers = {
            "Authorization": f"Bearer {self._access_token}",
            "Accept": "application/json",
        }
        if json_data is not None:
            headers["Content-Type"] = "application/json"

        resp = requests.request(
            method.upper(),
            url,
            headers=headers,
            params=params,
            json=json_data,
            timeout=60,
        )
        if resp.status_code == 204:
            return None
        try:
            data = resp.json()
        except Exception:
            text = resp.text
            if resp.ok:
                return text
            raise RuntimeError(f"Graph error {resp.status_code}: {text}") from None

        if not resp.ok:
            raise RuntimeError(f"Graph error {resp.status_code}: {json.dumps(data)}")
        return data

    def get(self, path: str, params: Optional[Dict[str, Any]] = None) -> Any:
        return self.request("GET", path, params=params)

    def post(self, path: str, json_data: Optional[Dict[str, Any]] = None) -> Any:
        return self.request("POST", path, json_data=json_data)

    def patch(self, path: str, json_data: Optional[Dict[str, Any]] = None) -> Any:
        return self.request("PATCH", path, json_data=json_data)

    def put(self, path: str, json_data: Optional[Dict[str, Any]] = None) -> Any:
        return self.request("PUT", path, json_data=json_data)

    def delete(self, path: str) -> Any:
        return self.request("DELETE", path)


def resource_path(relative: str) -> Path:

    if hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / relative  # type: ignore[attr-defined]
    return Path(__file__).with_name(relative)
    
class JsonHighlighter(QSyntaxHighlighter):
    def __init__(self, document):
        super().__init__(document)

        self.fmt_key = QTextCharFormat()
        self.fmt_key.setForeground(QColor("#60a5fa"))   # ключи

        self.fmt_string = QTextCharFormat()
        self.fmt_string.setForeground(QColor("#22c55e"))  # строки

        self.fmt_number = QTextCharFormat()
        self.fmt_number.setForeground(QColor("#facc15"))  # числа

        self.fmt_bool = QTextCharFormat()
        self.fmt_bool.setForeground(QColor("#f97316"))    # true/false

        self.fmt_null = QTextCharFormat()
        self.fmt_null.setForeground(QColor("#e5e7eb"))    # null

        self.re_key = QRegularExpression(r'"([^"\\]|\\.)*"(?=\s*:)')
        self.re_string = QRegularExpression(r'(?::\s*)("(?:[^"\\]|\\.)*")')
        self.re_number = QRegularExpression(r'\b-?\d+(?:\.\d+)?\b')
        self.re_bool = QRegularExpression(r'\b(true|false)\b')
        self.re_null = QRegularExpression(r'\bnull\b')

    def highlightBlock(self, text: str) -> None:
        # ключи
        it = self.re_key.globalMatch(text)
        while it.hasNext():
            m = it.next()
            self.setFormat(m.capturedStart(), m.capturedLength(), self.fmt_key)

        # строки-значения
        it = self.re_string.globalMatch(text)
        while it.hasNext():
            m = it.next()
            start = m.capturedStart(1)
            length = m.capturedLength(1)
            self.setFormat(start, length, self.fmt_string)

        # числа
        it = self.re_number.globalMatch(text)
        while it.hasNext():
            m = it.next()
            self.setFormat(m.capturedStart(), m.capturedLength(), self.fmt_number)

        # true/false
        it = self.re_bool.globalMatch(text)
        while it.hasNext():
            m = it.next()
            self.setFormat(m.capturedStart(), m.capturedLength(), self.fmt_bool)

        # null
        it = self.re_null.globalMatch(text)
        while it.hasNext():
            m = it.next()
            self.setFormat(m.capturedStart(), m.capturedLength(), self.fmt_null)


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()

        self.setWindowTitle("SwissKnife for Microsoft Graph")
        self.resize(1300, 850)

        icon_path = resource_path("swissknife_graph.ico")
        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))

        self.client: Optional[GraphClient] = None

        central = QWidget()
        self.setCentralWidget(central)

        root_layout = QVBoxLayout(central)

        # Auth
        auth_group = self._build_auth_group()
        root_layout.addWidget(auth_group)

        # Tabs + result area via splitter
        self.tabs = QTabWidget()

        self.view_tabs = QTabWidget()
        self.view_tabs.setTabPosition(QTabWidget.South)

        # Result views
        self.table_view = QTableWidget()
        self.details_view = QPlainTextEdit()
        self.details_view.setReadOnly(True)
        self.tree_view = QTreeWidget()
        self.tree_view.setHeaderLabels(["Key", "Value"])
        self.json_view = QPlainTextEdit()
        self.json_view.setReadOnly(True)

        self.details_highlighter = JsonHighlighter(self.details_view.document())
        self.json_highlighter = JsonHighlighter(self.json_view.document())

        self.view_tabs.addTab(self.table_view, "Table")
        self.view_tabs.addTab(self.details_view, "Details")
        self.view_tabs.addTab(self.tree_view, "Tree")
        self.view_tabs.addTab(self.json_view, "Raw JSON")

        # Build all tabs
        self._build_teams_tab()
        self._build_groups_tab()
        self._build_chats_tab()
        self._build_mail_calendar_tab()
        self._build_onedrive_tab()
        self._build_sharepoint_tab()
        self._build_admin_tab()
        self._build_intune_tab()
        self._build_audit_tab()
        self._build_raw_tab()
        self._build_about_tab()

        splitter = QSplitter(Qt.Vertical)
        splitter.addWidget(self.tabs)
        splitter.addWidget(self.view_tabs)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 2)
        splitter.setSizes([350, 500])

        root_layout.addWidget(splitter, stretch=1)

    # ---------- Auth group ----------
    def _build_auth_group(self) -> QGroupBox:
        group = QGroupBox("Auth (client credentials)")
        layout = QGridLayout(group)

        self.ed_tenant = QLineEdit()
        self.ed_client = QLineEdit()
        self.ed_secret = QLineEdit()
        self.ed_secret.setEchoMode(QLineEdit.Password)

        btn_connect = QPushButton("Connect")
        btn_connect.clicked.connect(self.on_connect_clicked)

        self.lbl_status = QLabel("Not connected")

        layout.addWidget(QLabel("Tenant ID:"), 0, 0)
        layout.addWidget(self.ed_tenant, 0, 1)
        layout.addWidget(QLabel("Client ID:"), 1, 0)
        layout.addWidget(self.ed_client, 1, 1)
        layout.addWidget(QLabel("Client Secret:"), 2, 0)
        layout.addWidget(self.ed_secret, 2, 1)
        layout.addWidget(btn_connect, 0, 2, 2, 1)
        layout.addWidget(self.lbl_status, 2, 2)

        # Theme switcher
        self.cb_theme = QComboBox()
        self.cb_theme.addItems(["Dark", "Light"])
        self.cb_theme.setCurrentText("Dark")
        self.cb_theme.currentTextChanged.connect(self.on_theme_changed)
        layout.addWidget(QLabel("Theme:"), 3, 0)
        layout.addWidget(self.cb_theme, 3, 1)

        return group

    # ---------- Tabs builders ----------
    def _build_teams_tab(self) -> None:
        tab = QWidget()
        layout = QVBoxLayout(tab)

        grid = QGridLayout()
        self.ed_teams_user = QLineEdit()
        btn_list_teams = QPushButton("List joined teams")
        btn_list_teams.clicked.connect(self.on_list_joined_teams)

        grid.addWidget(QLabel("User UPN:"), 0, 0)
        grid.addWidget(self.ed_teams_user, 0, 1)
        grid.addWidget(btn_list_teams, 0, 2)

        self.ed_team_id = QLineEdit()
        btn_list_channels = QPushButton("List channels")
        btn_list_channels.clicked.connect(self.on_list_team_channels)

        grid.addWidget(QLabel("Team (Group) ID:"), 1, 0)
        grid.addWidget(self.ed_team_id, 1, 1)
        grid.addWidget(btn_list_channels, 1, 2)

        layout.addLayout(grid)

        # Create channel group
        group_channel = QGroupBox("Create channel in Team")
        grid_ch = QGridLayout(group_channel)

        self.ed_ch_name = QLineEdit()
        self.ed_ch_desc = QLineEdit()
        self.cb_ch_type = QComboBox()
        self.cb_ch_type.addItems(["standard", "private", "shared"])
        self.ed_ch_owner = QLineEdit()

        btn_create_channel = QPushButton("Create channel")
        btn_create_channel.clicked.connect(self.on_create_channel)

        grid_ch.addWidget(QLabel("Name:"), 0, 0)
        grid_ch.addWidget(self.ed_ch_name, 0, 1)
        grid_ch.addWidget(btn_create_channel, 0, 2)
        grid_ch.addWidget(QLabel("Description:"), 1, 0)
        grid_ch.addWidget(self.ed_ch_desc, 1, 1, 1, 2)
        grid_ch.addWidget(QLabel("Type:"), 2, 0)
        grid_ch.addWidget(self.cb_ch_type, 2, 1)
        grid_ch.addWidget(QLabel("Owner UPN (private/shared):"), 3, 0)
        grid_ch.addWidget(self.ed_ch_owner, 3, 1, 1, 2)

        layout.addWidget(group_channel)

        # Add member group
        group_member = QGroupBox("Add member to Team / Channel")
        grid_m = QGridLayout(group_member)

        self.ed_member_channel_id = QLineEdit()
        self.ed_member_user_upn = QLineEdit()
        self.cb_member_role = QComboBox()
        self.cb_member_role.addItems(["member", "owner"])

        btn_add_member = QPushButton("Add member")
        btn_add_member.clicked.connect(self.on_add_member_team_channel)

        grid_m.addWidget(QLabel("Channel ID (optional, empty → Team):"), 0, 0)
        grid_m.addWidget(self.ed_member_channel_id, 0, 1)
        grid_m.addWidget(btn_add_member, 0, 2)
        grid_m.addWidget(QLabel("User UPN:"), 1, 0)
        grid_m.addWidget(self.ed_member_user_upn, 1, 1)
        grid_m.addWidget(QLabel("Role:"), 2, 0)
        grid_m.addWidget(self.cb_member_role, 2, 1)

        layout.addWidget(group_member)
        layout.addStretch(1)

        self.tabs.addTab(tab, "Teams")

    def _build_groups_tab(self) -> None:
        tab = QWidget()
        layout = QVBoxLayout(tab)

        group_create = QGroupBox("Create Microsoft 365 group")
        grid_c = QGridLayout(group_create)

        self.ed_group_display = QLineEdit()
        self.ed_group_mailnick = QLineEdit()
        self.ed_group_desc = QLineEdit()
        self.cb_group_teamify = QCheckBox("Create Team (teamify) after group is created")

        btn_create_group = QPushButton("Create group")
        btn_create_group.clicked.connect(self.on_create_group)

        grid_c.addWidget(QLabel("Display name:"), 0, 0)
        grid_c.addWidget(self.ed_group_display, 0, 1)
        grid_c.addWidget(btn_create_group, 0, 2)
        grid_c.addWidget(QLabel("Mail nickname:"), 1, 0)
        grid_c.addWidget(self.ed_group_mailnick, 1, 1)
        grid_c.addWidget(QLabel("Description:"), 2, 0)
        grid_c.addWidget(self.ed_group_desc, 2, 1, 1, 2)
        grid_c.addWidget(self.cb_group_teamify, 3, 0, 1, 3)

        layout.addWidget(group_create)

        group_manage = QGroupBox("Manage group")
        grid_m = QGridLayout(group_manage)

        self.ed_group_id = QLineEdit()
        self.ed_group_owner_upn = QLineEdit()
        self.ed_group_member_upn = QLineEdit()

        btn_add_owner = QPushButton("Add owner")
        btn_add_owner.clicked.connect(self.on_add_group_owner)
        btn_add_member = QPushButton("Add member")
        btn_add_member.clicked.connect(self.on_add_group_member)
        btn_teamify = QPushButton("Teamify → create Team")
        btn_teamify.clicked.connect(self.on_teamify_group)

        grid_m.addWidget(QLabel("Group ID:"), 0, 0)
        grid_m.addWidget(self.ed_group_id, 0, 1, 1, 2)
        grid_m.addWidget(QLabel("Owner UPN:"), 1, 0)
        grid_m.addWidget(self.ed_group_owner_upn, 1, 1)
        grid_m.addWidget(btn_add_owner, 1, 2)
        grid_m.addWidget(QLabel("Member UPN:"), 2, 0)
        grid_m.addWidget(self.ed_group_member_upn, 2, 1)
        grid_m.addWidget(btn_add_member, 2, 2)
        grid_m.addWidget(btn_teamify, 3, 0, 1, 3)

        layout.addWidget(group_manage)
        layout.addStretch(1)

        self.tabs.addTab(tab, "Groups")

    def _build_chats_tab(self) -> None:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        label = QLabel("Chats functionality can be used via Raw tab for now.")
        layout.addWidget(label)
        layout.addStretch(1)
        self.tabs.addTab(tab, "Chats")

    def _build_mail_calendar_tab(self) -> None:
        tab = QWidget()
        layout = QVBoxLayout(tab)

        grid = QGridLayout()
        self.ed_mail_user = QLineEdit()
        self.sp_mail_top = QSpinBox()
        self.sp_mail_top.setRange(1, 200)
        self.sp_mail_top.setValue(20)
        btn_list_mail = QPushButton("List messages")
        btn_list_mail.clicked.connect(self.on_list_messages)

        grid.addWidget(QLabel("User UPN:"), 0, 0)
        grid.addWidget(self.ed_mail_user, 0, 1)
        grid.addWidget(QLabel("Top:"), 0, 2)
        grid.addWidget(self.sp_mail_top, 0, 3)
        grid.addWidget(btn_list_mail, 0, 4)

        self.ed_mail_to = QLineEdit()
        self.ed_mail_subject = QLineEdit()
        self.ed_mail_body = QPlainTextEdit()
        btn_send_mail = QPushButton("Send mail")
        btn_send_mail.clicked.connect(self.on_send_mail)

        grid2 = QGridLayout()
        grid2.addWidget(QLabel("To:"), 0, 0)
        grid2.addWidget(self.ed_mail_to, 0, 1, 1, 3)
        grid2.addWidget(QLabel("Subject:"), 1, 0)
        grid2.addWidget(self.ed_mail_subject, 1, 1, 1, 3)
        grid2.addWidget(QLabel("Body:"), 2, 0)
        grid2.addWidget(self.ed_mail_body, 2, 1, 1, 3)
        grid2.addWidget(btn_send_mail, 3, 3)

        layout.addLayout(grid)
        layout.addLayout(grid2)
        layout.addStretch(1)

        self.tabs.addTab(tab, "Mail/Calendar")

    def _build_onedrive_tab(self) -> None:
        tab = QWidget()
        layout = QVBoxLayout(tab)

        grid = QGridLayout()
        self.ed_od_user = QLineEdit()
        self.ed_od_item = QLineEdit()
        self.ed_od_local = QLineEdit()
        self.ed_od_remote = QLineEdit()

        btn_list_root = QPushButton("List root")
        btn_list_root.clicked.connect(self.on_od_list_root)
        btn_browse = QPushButton("Browse...")
        btn_browse.clicked.connect(lambda: self._choose_file(self.ed_od_local))
        btn_download = QPushButton("Download")
        btn_download.clicked.connect(self.on_od_download)
        btn_upload = QPushButton("Upload")
        btn_upload.clicked.connect(self.on_od_upload)

        grid.addWidget(QLabel("User UPN:"), 0, 0)
        grid.addWidget(self.ed_od_user, 0, 1)
        grid.addWidget(btn_list_root, 0, 2)

        grid.addWidget(QLabel("Item ID:"), 1, 0)
        grid.addWidget(self.ed_od_item, 1, 1, 1, 2)

        grid.addWidget(QLabel("Local file:"), 2, 0)
        grid.addWidget(self.ed_od_local, 2, 1)
        grid.addWidget(btn_browse, 2, 2)

        grid.addWidget(QLabel("Remote path (upload):"), 3, 0)
        grid.addWidget(self.ed_od_remote, 3, 1, 1, 2)

        layout.addLayout(grid)

        grid2 = QGridLayout()
        grid2.addWidget(btn_download, 0, 0)
        grid2.addWidget(btn_upload, 0, 1)
        layout.addLayout(grid2)
        layout.addStretch(1)

        self.tabs.addTab(tab, "OneDrive")

    def _build_sharepoint_tab(self) -> None:
        tab = QWidget()
        layout = QVBoxLayout(tab)

        grid = QGridLayout()
        self.ed_sp_search = QLineEdit()
        btn_list_sites = QPushButton("List sites")
        btn_list_sites.clicked.connect(self.on_sp_list_sites)

        self.ed_sp_site_id = QLineEdit()
        self.ed_sp_item_id = QLineEdit()
        self.ed_sp_local = QLineEdit()
        self.ed_sp_remote = QLineEdit()

        btn_sp_browse = QPushButton("Browse...")
        btn_sp_browse.clicked.connect(lambda: self._choose_file(self.ed_sp_local))
        btn_sp_list_root = QPushButton("List root drive")
        btn_sp_list_root.clicked.connect(self.on_sp_list_root)
        btn_sp_download = QPushButton("Download")
        btn_sp_download.clicked.connect(self.on_sp_download)
        btn_sp_upload = QPushButton("Upload")
        btn_sp_upload.clicked.connect(self.on_sp_upload)

        grid.addWidget(QLabel("Search (optional):"), 0, 0)
        grid.addWidget(self.ed_sp_search, 0, 1)
        grid.addWidget(btn_list_sites, 0, 2)

        grid.addWidget(QLabel("Site ID:"), 1, 0)
        grid.addWidget(self.ed_sp_site_id, 1, 1, 1, 2)
        grid.addWidget(QLabel("Item ID:"), 2, 0)
        grid.addWidget(self.ed_sp_item_id, 2, 1, 1, 2)
        grid.addWidget(QLabel("Local file:"), 3, 0)
        grid.addWidget(self.ed_sp_local, 3, 1)
        grid.addWidget(btn_sp_browse, 3, 2)
        grid.addWidget(QLabel("Remote path (upload):"), 4, 0)
        grid.addWidget(self.ed_sp_remote, 4, 1, 1, 2)

        layout.addLayout(grid)

        grid2 = QGridLayout()
        grid2.addWidget(btn_sp_list_root, 0, 0)
        grid2.addWidget(btn_sp_download, 0, 1)
        grid2.addWidget(btn_sp_upload, 0, 2)
        layout.addLayout(grid2)

        layout.addStretch(1)
        self.tabs.addTab(tab, "SharePoint")

    def _build_admin_tab(self) -> None:
        tab = QWidget()
        layout = QVBoxLayout(tab)

        grid = QGridLayout()
        self.ed_admin_user = QLineEdit()

        btn_user_info = QPushButton("User info")
        btn_user_info.clicked.connect(self.on_admin_user_info)
        btn_block = QPushButton("Block")
        btn_block.clicked.connect(lambda: self.on_admin_block(True))
        btn_unblock = QPushButton("Unblock")
        btn_unblock.clicked.connect(lambda: self.on_admin_block(False))

        grid.addWidget(QLabel("User UPN:"), 0, 0)
        grid.addWidget(self.ed_admin_user, 0, 1, 1, 3)
        grid.addWidget(btn_user_info, 1, 0)
        grid.addWidget(btn_block, 1, 1)
        grid.addWidget(btn_unblock, 1, 2)

        layout.addLayout(grid)
        layout.addStretch(1)
        self.tabs.addTab(tab, "Admin")

    def _build_intune_tab(self) -> None:
        tab = QWidget()
        layout = QVBoxLayout(tab)

        grid = QGridLayout()
        self.sp_intune_top = QSpinBox()
        self.sp_intune_top.setRange(1, 200)
        self.sp_intune_top.setValue(20)
        btn_list_devices = QPushButton("List devices")
        btn_list_devices.clicked.connect(self.on_intune_list_devices)

        self.ed_intune_device_id = QLineEdit()
        self.cb_intune_keep_enrollment = QCheckBox("Keep enrollment data (wipe)")
        self.cb_intune_keep_user = QCheckBox("Keep user data (wipe)")

        btn_device_info = QPushButton("Device info")
        btn_device_info.clicked.connect(self.on_intune_device_info)
        btn_wipe = QPushButton("Wipe")
        btn_wipe.clicked.connect(self.on_intune_wipe)
        btn_retire = QPushButton("Retire")
        btn_retire.clicked.connect(self.on_intune_retire)

        grid.addWidget(QLabel("Top:"), 0, 0)
        grid.addWidget(self.sp_intune_top, 0, 1)
        grid.addWidget(btn_list_devices, 0, 2)
        grid.addWidget(QLabel("Device ID:"), 1, 0)
        grid.addWidget(self.ed_intune_device_id, 1, 1, 1, 2)
        grid.addWidget(self.cb_intune_keep_enrollment, 2, 0, 1, 2)
        grid.addWidget(self.cb_intune_keep_user, 2, 2, 1, 2)
        grid.addWidget(btn_device_info, 3, 0)
        grid.addWidget(btn_wipe, 3, 1)
        grid.addWidget(btn_retire, 3, 2)

        layout.addLayout(grid)
        layout.addStretch(1)

        self.tabs.addTab(tab, "Intune")

    def _build_audit_tab(self) -> None:
        tab = QWidget()
        layout = QVBoxLayout(tab)

        grid = QGridLayout()
        self.sp_audit_signins = QSpinBox()
        self.sp_audit_signins.setRange(1, 200)
        self.sp_audit_signins.setValue(20)
        self.sp_audit_dir = QSpinBox()
        self.sp_audit_dir.setRange(1, 200)
        self.sp_audit_dir.setValue(20)

        btn_signins = QPushButton("List sign-in logs")
        btn_signins.clicked.connect(self.on_audit_signins)
        btn_dir = QPushButton("List directory audit logs")
        btn_dir.clicked.connect(self.on_audit_directory)

        grid.addWidget(QLabel("Top sign-ins:"), 0, 0)
        grid.addWidget(self.sp_audit_signins, 0, 1)
        grid.addWidget(btn_signins, 0, 2)
        grid.addWidget(QLabel("Top directory audits:"), 1, 0)
        grid.addWidget(self.sp_audit_dir, 1, 1)
        grid.addWidget(btn_dir, 1, 2)

        layout.addLayout(grid)
        layout.addStretch(1)
        self.tabs.addTab(tab, "Audit")

    def _build_raw_tab(self) -> None:
        tab = QWidget()
        layout = QVBoxLayout(tab)

        grid = QGridLayout()
        self.cb_raw_method = QComboBox()
        self.cb_raw_method.addItems(["GET", "POST", "PATCH", "PUT", "DELETE"])
        self.ed_raw_path = QLineEdit()

        grid.addWidget(QLabel("Method:"), 0, 0)
        grid.addWidget(self.cb_raw_method, 0, 1)
        grid.addWidget(QLabel("Path / URL:"), 0, 2)
        grid.addWidget(self.ed_raw_path, 0, 3)

        self.ed_raw_body = QPlainTextEdit()
        self.ed_raw_body.setPlaceholderText('{"example": "value"} (optional)')

        btn_send = QPushButton("Send raw request")
        btn_send.clicked.connect(self.on_raw_send)

        examples = QPlainTextEdit()
        examples.setReadOnly(True)
        examples.setPlainText(
            "Examples:\n"
            "GET /organization\n"
            "GET /users\n"
            "GET /teams\n"
            "GET /users/user@domain.com/drive/root/children\n"
            "GET /auditLogs/signIns?$top=10\n"
        )

        layout.addLayout(grid)
        layout.addWidget(QLabel("JSON body (for POST / PATCH):"))
        layout.addWidget(self.ed_raw_body)
        layout.addWidget(btn_send)
        layout.addWidget(QLabel("Examples:"))
        layout.addWidget(examples)
        layout.addStretch(1)

        self.tabs.addTab(tab, "Raw")

    def _build_about_tab(self) -> None:
        tab = QWidget()
        layout = QVBoxLayout(tab)

        # Заголовок
        title = QLabel("🗡️ SwissKnife for Microsoft Graph")
        title.setStyleSheet("font-size: 12pt; font-weight: bold;")
        layout.addWidget(title)

        subtitle = QLabel(
            "Lightweight desktop toolkit around Microsoft Graph.\n"
            "Made for IT admins who prefer buttons over endless PowerShell scripts."
        )
        subtitle.setWordWrap(True)
        layout.addWidget(subtitle)

        sep = QLabel("────────────────────────────────────────")
        layout.addWidget(sep)

        # Список возможностей
        features = QLabel(
            "🧰 What it can help with:\n"
            "  • Teams / Groups / Channels management\n"
            "  • OneDrive & SharePoint files (list / download / upload)\n"
            "  • User admin (block / unblock / basic info)\n"
            "  • Intune managed devices (info / wipe / retire)\n"
            "  • Audit logs & raw Graph requests for advanced scenarios\n"
        )
        features.setWordWrap(True)
        layout.addWidget(features)

        # GitHub
        github = QLabel(
            "🔗 Project:\n"
            "  GitHub: https://github.com/Nemu-x/SwissKnife-for-MS-Graph"
        )
        github.setWordWrap(True)
        layout.addWidget(github)

        # Блок донатов — отдельным маленьким текстовым полем, чтобы удобно копировать
        donate_label = QLabel("💰 Support the project (optional):")
        layout.addWidget(donate_label)

        donate_box = QPlainTextEdit()
        donate_box.setReadOnly(True)
        donate_box.setFixedHeight(90)  # чтобы не было гигантского поля и скролла
        donate_box.setPlainText(
            "USDT (TRC20): 0xD9333e859Fb74D885d22E27568589de61E4433b5\n"
            "BTC:          bc1qkkcgpqym967k2x73al6f7fpvkx52q4rzkut3we\n"
            "ETH:          0xD9333e859Fb74D885d22E27568589de61E4433b5\n"
        )
        layout.addWidget(donate_box)

        footer = QLabel("Feedback, issues and PRs are very welcome 🙌")
        footer.setWordWrap(True)
        layout.addWidget(footer)

        layout.addStretch(1)

        self.tabs.addTab(tab, "About")


    # ---------- Theme switcher ----------
    def on_theme_changed(self, text: str) -> None:
        self.apply_theme(text.lower())

    def apply_theme(self, theme: str) -> None:
        app = QApplication.instance()
        if app is None:
            return
        if theme == "light":
            app.setStyleSheet(self._light_qss())
        else:
            app.setStyleSheet(self._dark_qss())

    @staticmethod
    def _dark_qss() -> str:
        return """
QWidget {
    background-color: #2E2E2E;
    color: #F2F2F2;
    font-family: "Segoe UI", sans-serif;
    font-size: 9pt;
}

/* GroupBox – рамки секций */
QGroupBox {
    border: 1px solid #555555;
    border-radius: 6px;
    margin-top: 10px;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 10px;
    padding: 0 4px;
    color: #DDDDDD;
    background-color: #2E2E2E;
}

/* Инпуты */
QLineEdit, QPlainTextEdit, QTextEdit, QComboBox {
    background-color: #262626;
    border: 1px solid #606060;
    border-radius: 4px;
    padding: 4px 6px;
    selection-background-color: #707070;
    selection-color: #FFFFFF;
}
QLineEdit:focus, QPlainTextEdit:focus, QTextEdit:focus, QComboBox:focus {
    border: 1px solid #AAAAAA;
}

/* Кнопки */
QPushButton {
    background-color: #383838;
    border: 1px solid #5A5A5A;
    border-radius: 4px;
    padding: 4px 10px;
    color: #F2F2F2;
}
QPushButton:hover {
    background-color: #444444;
}
QPushButton:pressed {
    background-color: #202020;
}
QPushButton:disabled {
    background-color: #2A2A2A;
    border-color: #404040;
    color: #888888;
}

/* Табы */
QTabWidget::pane {
    border: 1px solid #555555;
    top: -1px;
}
QTabBar::tab {
    background: #2E2E2E;
    border: 1px solid #555555;
    padding: 4px 10px;
    border-top-left-radius: 4px;
    border-top-right-radius: 4px;
    margin-right: 2px;
}
QTabBar::tab:selected {
    background: #383838;
    border-color: #AAAAAA;
}
QTabBar::tab:hover {
    background: #353535;
}

/* Таблица */
QTableView {
    gridline-color: #555555;
    background-color: #262626;
    alternate-background-color: #262626;
    selection-background-color: #505050;
    selection-color: #FFFFFF;
}
QHeaderView::section {
    background-color: #383838;
    color: #F2F2F2;
    padding: 4px;
    border: 1px solid #555555;
}

/* Дерево */
QTreeView {
    background-color: #262626;
    alternate-background-color: #262626;
    selection-background-color: #505050;
    selection-color: #FFFFFF;
    border: 1px solid #555555;
}

/* Скроллбары */
QScrollBar:vertical {
    background: #252525;
    width: 12px;
    margin: 0px;
}
QScrollBar::handle:vertical {
    background: #5C5C5C;
    min-height: 20px;
    border-radius: 4px;
}
QScrollBar::handle:vertical:hover {
    background: #777777;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0px;
}

QScrollBar:horizontal {
    background: #252525;
    height: 12px;
    margin: 0px;
}
QScrollBar::handle:horizontal {
    background: #5C5C5C;
    min-width: 20px;
    border-radius: 4px;
}
QScrollBar::handle:horizontal:hover {
    background: #777777;
}
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
    width: 0px;
}
"""

    @staticmethod
    def _light_qss() -> str:
        return """
QWidget {
    background-color: #f9fafb;
    color: #111827;
    font-family: "Segoe UI", sans-serif;
    font-size: 9pt;
}

QGroupBox {
    border: 1px solid #d1d5db;
    border-radius: 6px;
    margin-top: 10px;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 10px;
    padding: 0 4px;
    color: #6b7280;
    background-color: #f9fafb;
}

QLineEdit, QPlainTextEdit, QTextEdit, QComboBox {
    background-color: #ffffff;
    border: 1px solid #d1d5db;
    border-radius: 4px;
    padding: 4px 6px;
    selection-background-color: #2563eb;
}
QLineEdit:focus, QPlainTextEdit:focus, QTextEdit:focus, QComboBox:focus {
    border: 1px solid #2563eb;
}

QPushButton {
    background-color: #2563eb;
    border: 1px solid #1d4ed8;
    border-radius: 4px;
    padding: 4px 10px;
    color: #f9fafb;
}
QPushButton:hover {
    background-color: #1d4ed8;
}
QPushButton:pressed {
    background-color: #1e40af;
}
QPushButton:disabled {
    background-color: #e5e7eb;
    border-color: #d1d5db;
    color: #9ca3af;
}

QTabWidget::pane {
    border: 1px solid #d1d5db;
    top: -1px;
}
QTabBar::tab {
    background: #e5e7eb;
    border: 1px solid #d1d5db;
    border-bottom-color: #d1d5db;
    padding: 4px 10px;
    border-top-left-radius: 4px;
    border-top-right-radius: 4px;
    margin-right: 2px;
}
QTabBar::tab:selected {
    background: #ffffff;
    border-color: #2563eb;
}
QTabBar::tab:hover {
    background: #f3f4f6;
}

QTableView {
    gridline-color: #e5e7eb;
    background-color: #ffffff;
    alternate-background-color: #f9fafb;
    selection-background-color: #2563eb;
    selection-color: #f9fafb;
}
QHeaderView::section {
    background-color: #e5e7eb;
    color: #111827;
    padding: 4px;
    border: 1px solid #d1d5db;
}

QTreeView {
    background-color: #ffffff;
    alternate-background-color: #f9fafb;
    selection-background-color: #2563eb;
    selection-color: #f9fafb;
    border: 1px solid #d1d5db;
}

/* Scrollbars */
QScrollBar:vertical {
    background: #f3f4f6;
    width: 12px;
    margin: 0px;
}
QScrollBar::handle:vertical {
    background: #9ca3af;
    min-height: 20px;
    border-radius: 4px;
}
QScrollBar::handle:vertical:hover {
    background: #6b7280;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0px;
}

QScrollBar:horizontal {
    background: #f3f4f6;
    height: 12px;
    margin: 0px;
}
QScrollBar::handle:horizontal {
    background: #9ca3af;
    min-width: 20px;
    border-radius: 4px;
}
QScrollBar::handle:horizontal:hover {
    background: #6b7280;
}
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
    width: 0px;
}
"""

    # ---------- Helpers ----------
    def on_connect_clicked(self) -> None:
        tenant = self.ed_tenant.text().strip()
        client_id = self.ed_client.text().strip()
        secret = self.ed_secret.text().strip()
        if not tenant or not client_id or not secret:
            self.lbl_status.setText("Please fill Tenant / Client / Secret")
            self.lbl_status.setStyleSheet("color: #f97316;")
            return
        try:
            cfg = GraphConfig(tenant_id=tenant, client_id=client_id, client_secret=secret)
            self.client = GraphClient(cfg)
            self.client.authenticate()

            # 🔹 селф-тест: /organization
            org = self.client.get("/organization")

            self.lbl_status.setText("Connected")
            self.lbl_status.setStyleSheet("color: #22c55e;")

            # 🔹 показать результат в нижней панели
            self._display_result({
                "message": "Connected successfully",
                "organization": org,
            })
        except Exception as exc:
            self.lbl_status.setText("Auth failed")
            self.lbl_status.setStyleSheet("color: #f97316;")
            self._display_error(exc)


    def _ensure_client(self) -> Optional[GraphClient]:
        if not self.client:
            self._display_error(RuntimeError("Not connected"))
            return None
        return self.client

    def _display_error(self, exc: Exception) -> None:
        data = {"error": str(exc)}
        self._display_result(data)

    def _display_result(self, data: Any) -> None:
        # JSON view
        self.json_view.setPlainText(json.dumps(data, indent=2, ensure_ascii=False))

        # Table view (best-effort)
        self.table_view.clear()
        self.table_view.setRowCount(0)
        self.table_view.setColumnCount(0)

        rows: List[Dict[str, Any]] = []

        if isinstance(data, dict) and "value" in data and isinstance(data["value"], list):
            rows = data["value"]
        elif isinstance(data, list):
            rows = data
        elif isinstance(data, dict):
            rows = [data]
        else:
            rows = [{"value": data}]

        if rows:
            # Collect columns
            cols: List[str] = []
            for row in rows:
                if isinstance(row, dict):
                    for k in row.keys():
                        if k not in cols:
                            cols.append(k)
            self.table_view.setColumnCount(len(cols))
            self.table_view.setHorizontalHeaderLabels(cols)
            self.table_view.setRowCount(len(rows))

            for r, row in enumerate(rows):
                for c, col in enumerate(cols):
                    value = row.get(col, "")
                    text = json.dumps(value, ensure_ascii=False) if isinstance(value, (dict, list)) else str(value)
                    item = QTableWidgetItem(text)
                    self.table_view.setItem(r, c, item)

        # Details view – pretty JSON
        self.details_view.setPlainText(json.dumps(data, indent=2, ensure_ascii=False))

        # Tree view
        self.tree_view.clear()
        self._fill_tree(self.tree_view.invisibleRootItem(), data)

    def _fill_tree(self, parent: QTreeWidgetItem, data: Any, key: str = "") -> None:
        if isinstance(data, dict):
            node = parent
            if key:
                node = QTreeWidgetItem(parent, [key, ""])
            for k, v in data.items():
                self._fill_tree(node, v, str(k))
        elif isinstance(data, list):
            node = parent
            if key:
                node = QTreeWidgetItem(parent, [key, f"[{len(data)}]"])
            for idx, v in enumerate(data):
                self._fill_tree(node, v, f"[{idx}]")
        else:
            text = "" if data is None else str(data)
            QTreeWidgetItem(parent, [key, text])

    def _choose_file(self, line: QLineEdit) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Choose file")
        if path:
            line.setText(path)

    # ---------- Teams actions ----------
    def on_list_joined_teams(self) -> None:
        client = self._ensure_client()
        if not client:
            return
        upn = self.ed_teams_user.text().strip()
        if not upn:
            self._display_error(RuntimeError("User UPN is required"))
            return
        try:
            data = client.get(f"/users/{upn}/joinedTeams")
            self._display_result(data)
        except Exception as exc:
            self._display_error(exc)

    def on_list_team_channels(self) -> None:
        client = self._ensure_client()
        if not client:
            return
        team_id = self.ed_team_id.text().strip()
        if not team_id:
            self._display_error(RuntimeError("Team ID is required"))
            return
        try:
            data = client.get(f"/teams/{team_id}/channels")
            self._display_result(data)
        except Exception as exc:
            self._display_error(exc)

    def on_create_channel(self) -> None:
        client = self._ensure_client()
        if not client:
            return
        team_id = self.ed_team_id.text().strip()
        name = self.ed_ch_name.text().strip()
        ch_type = self.cb_ch_type.currentText()
        owner_upn = self.ed_ch_owner.text().strip()
        desc = self.ed_ch_desc.text().strip()

        if not team_id or not name:
            self._display_error(RuntimeError("Team ID and channel name are required"))
            return
        body: Dict[str, Any] = {
            "displayName": name,
            "description": desc,
            "membershipType": ch_type if ch_type != "standard" else "standard",
        }
        if ch_type in ("private", "shared"):
            if not owner_upn:
                self._display_error(RuntimeError("Owner UPN is required for private/shared channel"))
                return
            body["members"] = [
                {
                    "@odata.type": "#microsoft.graph.aadUserConversationMember",
                    "roles": ["owner"],
                    "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{owner_upn}')",
                }
            ]
        try:
            data = client.post(f"/teams/{team_id}/channels", json_data=body)
            self._display_result(data)
        except Exception as exc:
            self._display_error(exc)

    def on_add_member_team_channel(self) -> None:
        client = self._ensure_client()
        if not client:
            return
        team_id = self.ed_team_id.text().strip()
        channel_id = self.ed_member_channel_id.text().strip()
        user_upn = self.ed_member_user_upn.text().strip()
        role = self.cb_member_role.currentText()
        if not team_id or not user_upn:
            self._display_error(RuntimeError("Team ID and User UPN required"))
            return
        member = {
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            "roles": [role] if role == "owner" else [],
            "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{user_upn}')",
        }
        try:
            if channel_id:
                path = f"/teams/{team_id}/channels/{channel_id}/members"
            else:
                path = f"/teams/{team_id}/members"
            data = client.post(path, json_data=member)
            self._display_result(data)
        except Exception as exc:
            self._display_error(exc)

    # ---------- Groups actions ----------
    def on_create_group(self) -> None:
        client = self._ensure_client()
        if not client:
            return
        display = self.ed_group_display.text().strip()
        mailnick = self.ed_group_mailnick.text().strip()
        desc = self.ed_group_desc.text().strip()
        if not display or not mailnick:
            self._display_error(RuntimeError("Display name and mail nickname are required"))
            return
        body = {
            "displayName": display,
            "mailNickname": mailnick,
            "description": desc,
            "groupTypes": ["Unified"],
            "mailEnabled": True,
            "securityEnabled": False,
        }
        try:
            group = client.post("/groups", json_data=body)
            self._display_result(group)
            if self.cb_group_teamify.isChecked():
                gid = group.get("id")
                if gid:
                    self.ed_group_id.setText(gid)
                    self.on_teamify_group()
        except Exception as exc:
            self._display_error(exc)

    def on_add_group_owner(self) -> None:
        client = self._ensure_client()
        if not client:
            return
        gid = self.ed_group_id.text().strip()
        owner_upn = self.ed_group_owner_upn.text().strip()
        if not gid or not owner_upn:
            self._display_error(RuntimeError("Group ID and owner UPN required"))
            return
        body = {
            "@odata.id": f"https://graph.microsoft.com/v1.0/users/{owner_upn}"
        }
        try:
            client.post(f"/groups/{gid}/owners/$ref", json_data=body)
            self._display_result({"status": "owner added"})
        except Exception as exc:
            self._display_error(exc)

    def on_add_group_member(self) -> None:
        client = self._ensure_client()
        if not client:
            return
        gid = self.ed_group_id.text().strip()
        member_upn = self.ed_group_member_upn.text().strip()
        if not gid or not member_upn:
            self._display_error(RuntimeError("Group ID and member UPN required"))
            return
        body = {
            "@odata.id": f"https://graph.microsoft.com/v1.0/users/{member_upn}"
        }
        try:
            client.post(f"/groups/{gid}/members/$ref", json_data=body)
            self._display_result({"status": "member added"})
        except Exception as exc:
            self._display_error(exc)

    def on_teamify_group(self) -> None:
        client = self._ensure_client()
        if not client:
            return
        gid = self.ed_group_id.text().strip()
        if not gid:
            self._display_error(RuntimeError("Group ID is required"))
            return
        body = {
            "memberSettings": {
                "allowCreateUpdateChannels": True,
            },
            "messagingSettings": {
                "allowUserEditMessages": True,
                "allowUserDeleteMessages": True,
            },
            "funSettings": {
                "allowGiphy": True,
                "giphyContentRating": "strict",
            },
        }
        try:
            data = client.put(f"/groups/{gid}/team", json_data=body)
            self._display_result(data)
        except Exception as exc:
            self._display_error(exc)

    # ---------- Mail / Calendar ----------
    def on_list_messages(self) -> None:
        client = self._ensure_client()
        if not client:
            return
        upn = self.ed_mail_user.text().strip()
        if not upn:
            self._display_error(RuntimeError("User UPN is required"))
            return
        top = self.sp_mail_top.value()
        try:
            data = client.get(f"/users/{upn}/messages", params={"$top": top})
            self._display_result(data)
        except Exception as exc:
            self._display_error(exc)

    def on_send_mail(self) -> None:
        client = self._ensure_client()
        if not client:
            return
        upn = self.ed_mail_user.text().strip()
        to = self.ed_mail_to.text().strip()
        subject = self.ed_mail_subject.text().strip()
        body_text = self.ed_mail_body.toPlainText()
        if not upn or not to or not subject:
            self._display_error(RuntimeError("User, To, Subject are required"))
            return
        msg = {
            "message": {
                "subject": subject,
                "body": {"contentType": "Text", "content": body_text},
                "toRecipients": [{"emailAddress": {"address": to}}],
            },
            "saveToSentItems": True,
        }
        try:
            client.post(f"/users/{upn}/sendMail", json_data=msg)
            self._display_result({"status": "mail sent"})
        except Exception as exc:
            self._display_error(exc)

    # ---------- OneDrive ----------
    def on_od_list_root(self) -> None:
        client = self._ensure_client()
        if not client:
            return
        upn = self.ed_od_user.text().strip()
        if not upn:
            self._display_error(RuntimeError("User UPN is required"))
            return
        try:
            data = client.get(f"/users/{upn}/drive/root/children")
            self._display_result(data)
        except Exception as exc:
            self._display_error(exc)

    def on_od_download(self) -> None:
        client = self._ensure_client()
        if not client:
            return
        upn = self.ed_od_user.text().strip()
        item_id = self.ed_od_item.text().strip()
        local = self.ed_od_local.text().strip()
        if not upn or not item_id or not local:
            self._display_error(RuntimeError("User, item ID and local file are required"))
            return
        try:
            url = f"/users/{upn}/drive/items/{item_id}/content"
            # use request to get raw bytes
            if not client._access_token:
                raise RuntimeError("Not authenticated")
            headers = {
                "Authorization": f"Bearer {client._access_token}",
            }
            full_url = GRAPH_BASE_URL + url
            resp = requests.get(full_url, headers=headers, timeout=120)
            if not resp.ok:
                raise RuntimeError(f"Download failed: {resp.status_code} {resp.text}")
            with open(local, "wb") as f:
                f.write(resp.content)
            self._display_result({"status": f"downloaded to {local}"})
        except Exception as exc:
            self._display_error(exc)

    def on_od_upload(self) -> None:
        client = self._ensure_client()
        if not client:
            return
        upn = self.ed_od_user.text().strip()
        local = self.ed_od_local.text().strip()
        remote = self.ed_od_remote.text().strip().lstrip("/")
        if not upn or not local or not remote:
            self._display_error(RuntimeError("User, local file and remote path are required"))
            return
        try:
            with open(local, "rb") as f:
                content = f.read()
            if not client._access_token:
                raise RuntimeError("Not authenticated")
            headers = {
                "Authorization": f"Bearer {client._access_token}",
                "Content-Type": "application/octet-stream",
            }
            url = GRAPH_BASE_URL + f"/users/{upn}/drive/root:/{remote}:/content"
            resp = requests.put(url, headers=headers, data=content, timeout=300)
            if not resp.ok:
                raise RuntimeError(f"Upload failed: {resp.status_code} {resp.text}")
            data = resp.json()
            self._display_result(data)
        except Exception as exc:
            self._display_error(exc)

    # ---------- SharePoint ----------
    def on_sp_list_sites(self) -> None:
        client = self._ensure_client()
        if not client:
            return
        search = self.ed_sp_search.text().strip()
        try:
            if search:
                data = client.get("/sites", params={"search": search})
            else:
                data = client.get("/sites?search=*")
            self._display_result(data)
        except Exception as exc:
            self._display_error(exc)

    def on_sp_list_root(self) -> None:
        client = self._ensure_client()
        if not client:
            return
        site_id = self.ed_sp_site_id.text().strip()
        if not site_id:
            self._display_error(RuntimeError("Site ID is required"))
            return
        try:
            data = client.get(f"/sites/{site_id}/drive/root/children")
            self._display_result(data)
        except Exception as exc:
            self._display_error(exc)

    def on_sp_download(self) -> None:
        client = self._ensure_client()
        if not client:
            return
        site_id = self.ed_sp_site_id.text().strip()
        item_id = self.ed_sp_item_id.text().strip()
        local = self.ed_sp_local.text().strip()
        if not site_id or not item_id or not local:
            self._display_error(RuntimeError("Site, item ID and local file are required"))
            return
        try:
            if not client._access_token:
                raise RuntimeError("Not authenticated")
            headers = {
                "Authorization": f"Bearer {client._access_token}",
            }
            url = GRAPH_BASE_URL + f"/sites/{site_id}/drive/items/{item_id}/content"
            resp = requests.get(url, headers=headers, timeout=120)
            if not resp.ok:
                raise RuntimeError(f"Download failed: {resp.status_code} {resp.text}")
            with open(local, "wb") as f:
                f.write(resp.content)
            self._display_result({"status": f"downloaded to {local}"})
        except Exception as exc:
            self._display_error(exc)

    def on_sp_upload(self) -> None:
        client = self._ensure_client()
        if not client:
            return
        site_id = self.ed_sp_site_id.text().strip()
        local = self.ed_sp_local.text().strip()
        remote = self.ed_sp_remote.text().strip().lstrip("/")
        if not site_id or not local or not remote:
            self._display_error(RuntimeError("Site, local file and remote path are required"))
            return
        try:
            with open(local, "rb") as f:
                content = f.read()
            if not client._access_token:
                raise RuntimeError("Not authenticated")
            headers = {
                "Authorization": f"Bearer {client._access_token}",
                "Content-Type": "application/octet-stream",
            }
            url = GRAPH_BASE_URL + f"/sites/{site_id}/drive/root:/{remote}:/content"
            resp = requests.put(url, headers=headers, data=content, timeout=300)
            if not resp.ok:
                raise RuntimeError(f"Upload failed: {resp.status_code} {resp.text}")
            data = resp.json()
            self._display_result(data)
        except Exception as exc:
            self._display_error(exc)

    # ---------- Admin ----------
    def on_admin_user_info(self) -> None:
        client = self._ensure_client()
        if not client:
            return
        upn = self.ed_admin_user.text().strip()
        if not upn:
            self._display_error(RuntimeError("User UPN required"))
            return
        try:
            data = client.get(f"/users/{upn}")
            self._display_result(data)
        except Exception as exc:
            self._display_error(exc)

    def on_admin_block(self, block: bool) -> None:
        client = self._ensure_client()
        if not client:
            return
        upn = self.ed_admin_user.text().strip()
        if not upn:
            self._display_error(RuntimeError("User UPN required"))
            return
        body = {"accountEnabled": not block}
        try:
            client.patch(f"/users/{upn}", json_data=body)
            self._display_result({"status": "blocked" if block else "unblocked"})
        except Exception as exc:
            self._display_error(exc)

    # ---------- Intune ----------
    def on_intune_list_devices(self) -> None:
        client = self._ensure_client()
        if not client:
            return
        top = self.sp_intune_top.value()
        try:
            data = client.get("/deviceManagement/managedDevices", params={"$top": top})
            self._display_result(data)
        except Exception as exc:
            self._display_error(exc)

    def on_intune_device_info(self) -> None:
        client = self._ensure_client()
        if not client:
            return
        did = self.ed_intune_device_id.text().strip()
        if not did:
            self._display_error(RuntimeError("Device ID required"))
            return
        try:
            data = client.get(f"/deviceManagement/managedDevices/{did}")
            self._display_result(data)
        except Exception as exc:
            self._display_error(exc)

    def on_intune_wipe(self) -> None:
        client = self._ensure_client()
        if not client:
            return
        did = self.ed_intune_device_id.text().strip()
        if not did:
            self._display_error(RuntimeError("Device ID required"))
            return
        body = {
            "keepEnrollmentData": self.cb_intune_keep_enrollment.isChecked(),
            "keepUserData": self.cb_intune_keep_user.isChecked(),
            "useProtectedWipe": False,
        }
        try:
            client.post(f"/deviceManagement/managedDevices/{did}/wipe", json_data=body)
            self._display_result({"status": "wipe requested"})
        except Exception as exc:
            self._display_error(exc)

    def on_intune_retire(self) -> None:
        client = self._ensure_client()
        if not client:
            return
        did = self.ed_intune_device_id.text().strip()
        if not did:
            self._display_error(RuntimeError("Device ID required"))
            return
        try:
            client.post(f"/deviceManagement/managedDevices/{did}/retire", json_data={})
            self._display_result({"status": "retire requested"})
        except Exception as exc:
            self._display_error(exc)

    # ---------- Audit ----------
    def on_audit_signins(self) -> None:
        client = self._ensure_client()
        if not client:
            return
        top = self.sp_audit_signins.value()
        try:
            data = client.get("/auditLogs/signIns", params={"$top": top})
            self._display_result(data)
        except Exception as exc:
            self._display_error(exc)

    def on_audit_directory(self) -> None:
        client = self._ensure_client()
        if not client:
            return
        top = self.sp_audit_dir.value()
        try:
            data = client.get("/auditLogs/directoryAudits", params={"$top": top})
            self._display_result(data)
        except Exception as exc:
            self._display_error(exc)

    # ---------- Raw ----------
    def on_raw_send(self) -> None:
        client = self._ensure_client()
        if not client:
            return
        method = self.cb_raw_method.currentText()
        path = self.ed_raw_path.text().strip()
        if not path:
            self._display_error(RuntimeError("Path / URL is required"))
            return
        body_text = self.ed_raw_body.toPlainText().strip()
        json_body = None
        if body_text:
            try:
                json_body = json.loads(body_text)
            except Exception as exc:
                self._display_error(RuntimeError(f"Invalid JSON body: {exc}"))
                return
        try:
            data = client.request(method, path, json_data=json_body)
            self._display_result(data if data is not None else {"status": "No content"})
        except Exception as exc:
            self._display_error(exc)


def main() -> None:
    app = QApplication(sys.argv)
    app.setApplicationName("SwissKnife Graph GUI")

    win = MainWindow()
    win.apply_theme("dark")
    win.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
