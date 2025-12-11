import json

import typer
from rich.console import Console
from rich.table import Table
from rich import print_json

from swissknife.auth import acquire_token_client_credentials
from swissknife.graph_client import GraphClient
from swissknife import chats as chats_module
from swissknife import users as users_module
from swissknife import groups as groups_module
from swissknife import teams as teams_module
from swissknife import licensing as licensing_module
from swissknife import mail as mail_module
from swissknife import calendar_api as calendar_module
from swissknife import teams_create as teams_create_module
from swissknife import audit as audit_module
from swissknife import onedrive as onedrive_module
from swissknife import sharepoint as sharepoint_module
from swissknife import admin as admin_module
from swissknife import intune as intune_module



app = typer.Typer(help="Swissknife for Microsoft Graph (MVP CLI)")
console = Console()


def build_graph_client(
    tenant_id: str,
    client_id: str,
    client_secret: str,
) -> GraphClient:
    """
    Получаем токен и создаём GraphClient.
    """
    token = acquire_token_client_credentials(
        tenant_id=tenant_id,
        client_id=client_id,
        client_secret=client_secret,
    )
    return GraphClient(access_token=token)


# --- AUTH TEST --- #


@app.command("auth-test")
def auth_test(
    tenant_id: str = typer.Option(
        ...,
        "--tenant-id",
        envvar="GRAPH_TENANT_ID",
        help="Tenant ID (GUID) вашего Microsoft 365 / Entra ID.",
    ),
    client_id: str = typer.Option(
        ...,
        "--client-id",
        envvar="GRAPH_CLIENT_ID",
        help="Application (client) ID зарегистрированного приложения.",
    ),
    client_secret: str = typer.Option(
        ...,
        "--client-secret",
        prompt=True,
        hide_input=True,
        envvar="GRAPH_CLIENT_SECRET",
        help="Client secret для приложения.",
    ),
):
    """
    Тест получения токена и простого запроса в Graph.
    Сейчас проверяем /chats?$top=1 (нужен Chat.Read.All как Application).
    """
    console.rule("[bold]Проверка авторизации[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    try:
        data = client.get("/chats?$top=1")
        console.print(
            "[bold green]OK[/bold green] Токен получен, запрос к /chats успешен."
        )
        print_json(data=data)
    except RuntimeError as e:
        console.print("[red]Ошибка при запросе к Graph[/red]")
        console.print(str(e))
        console.print(
            "\n[yellow]Проверь, что у приложения есть нужные Application permissions "
            "в разделе API Permissions и что им выдан admin consent.[/yellow]"
        )
        raise typer.Exit(code=1)

# --- ADMIN --- #

admin_app = typer.Typer(help="Админ-консоль (пользователи)")
app.add_typer(admin_app, name="admin")


@admin_app.command("user-info")
def admin_user_info(
    user: str = typer.Argument(..., help="UPN или id пользователя."),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(
        ...,
        "--client-secret",
        prompt=True,
        hide_input=True,
        envvar="GRAPH_CLIENT_SECRET",
    ),
):
    """
    Базовая инфа о пользователе для админов: upn, mail, accountEnabled.
    """
    console.rule(f"[bold]Admin: user info {user}[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    data = admin_module.get_user_basic(client, user)
    print_json(data=data)


@admin_app.command("block")
def admin_block_user(
    user: str = typer.Argument(..., help="UPN или id пользователя, которого блокируем."),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(
        ...,
        "--client-secret",
        prompt=True,
        hide_input=True,
        envvar="GRAPH_CLIENT_SECRET",
    ),
):
    """
    Заблокировать пользователя (accountEnabled = false).
    """
    console.rule(f"[bold red]Admin: BLOCK user {user}[/bold red]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    result = admin_module.block_user(client, user)
    console.print("[bold red]Пользователь заблокирован.[/bold red]")
    print_json(data=result)


@admin_app.command("unblock")
def admin_unblock_user(
    user: str = typer.Argument(..., help="UPN или id пользователя, которого разблокируем."),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(
        ...,
        "--client-secret",
        prompt=True,
        hide_input=True,
        envvar="GRAPH_CLIENT_SECRET",
    ),
):
    """
    Разблокировать пользователя (accountEnabled = true).
    """
    console.rule(f"[bold green]Admin: UNBLOCK user {user}[/bold green]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    result = admin_module.unblock_user(client, user)
    console.print("[bold green]Пользователь разблокирован.[/bold green]")
    print_json(data=result)


@admin_app.command("reset-password")
def admin_reset_password(
    user: str = typer.Argument(..., help="UPN или id пользователя."),
    new_password: str = typer.Option(
        ...,
        "--password",
        prompt=True,
        hide_input=True,
        confirmation_prompt=True,
        help="Новый пароль.",
    ),
    force_change: bool = typer.Option(
        True,
        "--force-change/--no-force-change",
        help="Требовать смену пароля при следующем входе.",
    ),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(
        ...,
        "--client-secret",
        prompt=True,
        hide_input=True,
        envvar="GRAPH_CLIENT_SECRET",
    ),
):
    """
    Сбросить пароль пользователю.
    Внимание: для синхронизированных из локального AD аккаунтов может не сработать.
    """
    console.rule(f"[bold]Admin: reset password for {user}[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    result = admin_module.reset_password(
        client,
        user,
        new_password=new_password,
        force_change_next_signin=force_change,
    )

    console.print("[bold green]Пароль сброшен (если Graph не вернул ошибку).[/bold green]")
    print_json(data=result)


@admin_app.command("revoke-sessions")
def admin_revoke_sessions(
    user: str = typer.Argument(..., help="UPN или id пользователя."),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(
        ...,
        "--client-secret",
        prompt=True,
        hide_input=True,
        envvar="GRAPH_CLIENT_SECRET",
    ),
):
    """
    Принудительно разлогинить пользователя (revokeSignInSessions).
    """
    console.rule(f"[bold]Admin: revoke sessions for {user}[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    result = admin_module.revoke_sessions(client, user)
    console.print("[bold yellow]Сессии пользователя отозваны (по данным Graph).[/bold yellow]")
    print_json(data=result)


# --- AUDIT --- #

audit_app = typer.Typer(help="Audit & Sign-in Logs")
app.add_typer(audit_app, name="audit")

@audit_app.command("logs")
def audit_logs(
    top: int = typer.Option(50, "--top"),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(..., "--client-secret", prompt=True, hide_input=True, envvar="GRAPH_CLIENT_SECRET"),
):
    console.rule("[bold]Audit Logs[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    logs = audit_module.audit_logs(client, top)
    print_json(data=logs)

@audit_app.command("signin")
def audit_signin(
    top: int = typer.Option(50, "--top"),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(..., "--client-secret", prompt=True, hide_input=True, envvar="GRAPH_CLIENT_SECRET"),
):
    console.rule("[bold]Sign-in Logs[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    logs = audit_module.sign_in_logs(client, top)
    print_json(data=logs)


# --- INTUNE --- #

intune_app = typer.Typer(help="Intune / managed devices")
app.add_typer(intune_app, name="intune")

@intune_app.command("devices")
def intune_devices(
    top: int = typer.Option(50, "--top", help="Сколько устройств показывать."),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(
        ...,
        "--client-secret",
        prompt=True,
        hide_input=True,
        envvar="GRAPH_CLIENT_SECRET",
    ),
):
    """
    Список управляемых Intune устройств.
    """
    console.rule("[bold]Intune: managed devices[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    devices = intune_module.list_managed_devices(client, top=top)
    if not devices:
        console.print("[yellow]Устройств не найдено или нет доступа.[/yellow]")
        raise typer.Exit(0)

    table = Table(show_header=True, header_style="bold blue")
    table.add_column("Index", style="dim", width=6)
    table.add_column("Device ID", overflow="fold")
    table.add_column("Device Name", overflow="fold")
    table.add_column("OS", overflow="fold")
    table.add_column("User", overflow="fold")

    for idx, d in enumerate(devices, start=1):
        table.add_row(
            str(idx),
            d.get("id", ""),
            d.get("deviceName", ""),
            d.get("operatingSystem", ""),
            d.get("userPrincipalName", "") or "",
        )

    console.print(table)


@intune_app.command("device")
def intune_device(
    device_id: str = typer.Argument(..., help="ID устройства."),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(
        ...,
        "--client-secret",
        prompt=True,
        hide_input=True,
        envvar="GRAPH_CLIENT_SECRET",
    ),
):
    """
    Подробная информация об одном устройстве.
    """
    console.rule(f"[bold]Intune: device {device_id}[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    data = intune_module.get_managed_device(client, device_id)
    print_json(data=data)


@intune_app.command("wipe")
def intune_wipe(
    device_id: str = typer.Argument(..., help="ID устройства."),
    keep_enrollment: bool = typer.Option(
        False,
        "--keep-enrollment/--drop-enrollment",
        help="Сохранить ли данные о регистрации.",
    ),
    keep_user_data: bool = typer.Option(
        False,
        "--keep-user-data/--drop-user-data",
        help="Сохранить ли пользовательские данные (если поддерживается).",
    ),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(
        ...,
        "--client-secret",
        prompt=True,
        hide_input=True,
        envvar="GRAPH_CLIENT_SECRET",
    ),
):
    """
    Wipe устройства (Intune).
    """
    console.rule(f"[bold red]Intune: WIPE device {device_id}[/bold red]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    result = intune_module.wipe_device(
        client,
        device_id,
        keep_enrollment_data=keep_enrollment,
        keep_user_data=keep_user_data,
    )
    console.print("[bold red]Команда wipe отправлена.[/bold red]")
    print_json(data=result)


@intune_app.command("retire")
def intune_retire(
    device_id: str = typer.Argument(..., help="ID устройства."),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(
        ...,
        "--client-secret",
        prompt=True,
        hide_input=True,
        envvar="GRAPH_CLIENT_SECRET",
    ),
):
    """
    Retire устройства.
    """
    console.rule(f"[bold]Intune: RETIRE device {device_id}[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    result = intune_module.retire_device(client, device_id)
    console.print("[bold yellow]Команда retire отправлена.[/bold yellow]")
    print_json(data=result)


@intune_app.command("lock")
def intune_lock(
    device_id: str = typer.Argument(..., help="ID устройства."),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(
        ...,
        "--client-secret",
        prompt=True,
        hide_input=True,
        envvar="GRAPH_CLIENT_SECRET",
    ),
):
    """
    Remote lock устройства.
    """
    console.rule(f"[bold]Intune: REMOTE LOCK device {device_id}[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    result = intune_module.remote_lock_device(client, device_id)
    console.print("[bold yellow]Команда remote lock отправлена.[/bold yellow]")
    print_json(data=result)


# --- CHATS --- #

chats_app = typer.Typer(help="Операции с чатами")
app.add_typer(chats_app, name="chats")


@chats_app.command("list")
def chats_list(
    user: str = typer.Argument(
        ...,
        help="UPN или id пользователя (например, user1@example.com).",
    ),
    tenant_id: str = typer.Option(
        ...,
        "--tenant-id",
        envvar="GRAPH_TENANT_ID",
    ),
    client_id: str = typer.Option(
        ...,
        "--client-id",
        envvar="GRAPH_CLIENT_ID",
    ),
    client_secret: str = typer.Option(
        ...,
        "--client-secret",
        prompt=True,
        hide_input=True,
        envvar="GRAPH_CLIENT_SECRET",
    ),
):
    """
    Вывести список чатов пользователя.
    """
    console.rule(f"[bold]Чаты пользователя {user}[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    chats = chats_module.list_user_chats(client, user)

    if not chats:
        console.print("[yellow]Чатов не найдено.[/yellow]")
        raise typer.Exit(code=0)

    table = Table(show_header=True, header_style="bold magenta")
    table.add_column("Index", style="dim", width=6)
    table.add_column("Chat ID", overflow="fold")
    table.add_column("Topic", overflow="fold")
    table.add_column("Chat type", overflow="fold")

    for idx, chat in enumerate(chats, start=1):
        table.add_row(
            str(idx),
            chat.get("id", ""),
            chat.get("topic", ""),
            chat.get("chatType", ""),
        )

    console.print(table)


@chats_app.command("add-member")
def chats_add_member(
    chat_id: str = typer.Argument(..., help="ID чата, куда добавляем участника."),
    user_upn: str = typer.Argument(..., help="UPN участника, например user2@example.com."),
    owner: bool = typer.Option(
        False,
        "--owner",
        help="Если указать, участник будет добавлен как owner.",
    ),
    tenant_id: str = typer.Option(
        ...,
        "--tenant-id",
        envvar="GRAPH_TENANT_ID",
    ),
    client_id: str = typer.Option(
        ...,
        "--client-id",
        envvar="GRAPH_CLIENT_ID",
    ),
    client_secret: str = typer.Option(
        ...,
        "--client-secret",
        prompt=True,
        hide_input=True,
        envvar="GRAPH_CLIENT_SECRET",
    ),
):
    """
    Добавить участника в чат по chat_id.
    """
    console.rule("[bold]Добавление участника в чат[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    result = chats_module.add_user_to_chat(
        client=client,
        chat_id=chat_id,
        user_upn=user_upn,
        as_owner=owner,
    )

    console.print("[bold green]Участник добавлен.[/bold green]")
    print_json(data=result)


@chats_app.command("remove-member")
def chats_remove_member(
    chat_id: str = typer.Argument(..., help="ID чата."),
    user_upn: str = typer.Argument(
        ..., help="UPN/email участника, которого нужно удалить."
    ),
    tenant_id: str = typer.Option(
        ...,
        "--tenant-id",
        envvar="GRAPH_TENANT_ID",
    ),
    client_id: str = typer.Option(
        ...,
        "--client-id",
        envvar="GRAPH_CLIENT_ID",
    ),
    client_secret: str = typer.Option(
        ...,
        "--client-secret",
        prompt=True,
        hide_input=True,
        envvar="GRAPH_CLIENT_SECRET",
    ),
):
    """
    Удалить участника из чата по UPN/email.
    """
    console.rule("[bold]Удаление участника из чата[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    chats_module.remove_user_from_chat(client, chat_id, user_upn)
    console.print("[bold green]Участник удалён из чата.[/bold green]")


@chats_app.command("members")
def chats_members(
    chat_id: str = typer.Argument(..., help="ID чата."),
    tenant_id: str = typer.Option(
        ...,
        "--tenant-id",
        envvar="GRAPH_TENANT_ID",
    ),
    client_id: str = typer.Option(
        ...,
        "--client-id",
        envvar="GRAPH_CLIENT_ID",
    ),
    client_secret: str = typer.Option(
        ...,
        "--client-secret",
        prompt=True,
        hide_input=True,
        envvar="GRAPH_CLIENT_SECRET",
    ),
):
    """
    Список участников чата.
    """
    console.rule(f"[bold]Участники чата {chat_id}[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    members = chats_module.list_chat_members(client, chat_id)
    if not members:
        console.print("[yellow]Участники не найдены.[/yellow]")
        raise typer.Exit(code=0)

    table = Table(show_header=True, header_style="bold cyan")
    table.add_column("Index", style="dim", width=6)
    table.add_column("Membership ID", overflow="fold")
    table.add_column("DisplayName", overflow="fold")
    table.add_column("Email", overflow="fold")
    table.add_column("Roles", overflow="fold")

    for idx, m in enumerate(members, start=1):
        roles = ", ".join(m.get("roles", []))
        table.add_row(
            str(idx),
            m.get("id", ""),
            m.get("displayName", ""),
            m.get("email", "") or "",
            roles,
        )

    console.print(table)


@chats_app.command("messages")
def chats_messages(
    chat_id: str = typer.Argument(..., help="ID чата."),
    top: int = typer.Option(50, "--top", help="Сколько последних сообщений забрать."),
    tenant_id: str = typer.Option(
        ...,
        "--tenant-id",
        envvar="GRAPH_TENANT_ID",
    ),
    client_id: str = typer.Option(
        ...,
        "--client-id",
        envvar="GRAPH_CLIENT_ID",
    ),
    client_secret: str = typer.Option(
        ...,
        "--client-secret",
        prompt=True,
        hide_input=True,
        envvar="GRAPH_CLIENT_SECRET",
    ),
):
    """
    Получить последние N сообщений из чата.
    """
    console.rule(f"[bold]Сообщения чата {chat_id}[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    messages = chats_module.get_chat_messages(client, chat_id, top=top)
    if not messages:
        console.print("[yellow]Сообщений не найдено.[/yellow]")
        raise typer.Exit(code=0)

    table = Table(show_header=True, header_style="bold cyan")
    table.add_column("Index", style="dim", width=6)
    table.add_column("From", overflow="fold")
    table.add_column("Created", overflow="fold")
    table.add_column("Preview", overflow="fold")

    for idx, msg in enumerate(messages, start=1):
        from_user = ""
        frm = msg.get("from", {})
        if frm:
            user = frm.get("user") or {}
            from_user = user.get("displayName") or user.get("id") or ""

        created = msg.get("createdDateTime", "")
        body = (msg.get("body") or {}).get("content", "") or ""
        preview = body.replace("\n", " ")
        if len(preview) > 80:
            preview = preview[:77] + "..."

        table.add_row(str(idx), from_user, created, preview)

    console.print(table)


# --- USERS --- #

users_app = typer.Typer(help="Операции с пользователями (Azure AD)")
app.add_typer(users_app, name="users")


@users_app.command("list")
def users_list(
    top: int = typer.Option(25, "--top", help="Сколько пользователей показать."),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(
        ...,
        "--client-secret",
        prompt=True,
        hide_input=True,
        envvar="GRAPH_CLIENT_SECRET",
    ),
):
    """
    Список пользователей (первые N).
    """
    console.rule("[bold]Пользователи[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    users = users_module.list_users(client, top=top)
    if not users:
        console.print("[yellow]Пользователи не найдены.[/yellow]")
        raise typer.Exit(code=0)

    table = Table(show_header=True, header_style="bold magenta")
    table.add_column("Index", style="dim", width=6)
    table.add_column("User ID", overflow="fold")
    table.add_column("UPN", overflow="fold")
    table.add_column("DisplayName", overflow="fold")

    for idx, u in enumerate(users, start=1):
        table.add_row(
            str(idx),
            u.get("id", ""),
            u.get("userPrincipalName", ""),
            u.get("displayName", ""),
        )

    console.print(table)


@users_app.command("get")
def users_get(
    user: str = typer.Argument(
        ...,
        help="UPN или id пользователя (например, user1@example.com).",
    ),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(
        ...,
        "--client-secret",
        prompt=True,
        hide_input=True,
        envvar="GRAPH_CLIENT_SECRET",
    ),
):
    """
    Получить подробную информацию о пользователе.
    """
    console.rule(f"[bold]Пользователь {user}[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    data = users_module.get_user(client, user)
    print_json(data=data)


@users_app.command("groups")
def users_groups(
    user: str = typer.Argument(..., help="UPN или id пользователя."),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(
        ...,
        "--client-secret",
        prompt=True,
        hide_input=True,
        envvar="GRAPH_CLIENT_SECRET",
    ),
):
    """
    Группы, в которых состоит пользователь.
    """
    console.rule(f"[bold]Группы пользователя {user}[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    items = users_module.get_user_member_of(client, user)
    if not items:
        console.print("[yellow]Группы не найдены.[/yellow]")
        raise typer.Exit(code=0)

    table = Table(show_header=True, header_style="bold magenta")
    table.add_column("Index", style="dim", width=6)
    table.add_column("Object ID", overflow="fold")
    table.add_column("Type", overflow="fold")
    table.add_column("DisplayName", overflow="fold")

    for idx, obj in enumerate(items, start=1):
        otype = obj.get("@odata.type", "")
        display_name = obj.get("displayName", "")
        oid = obj.get("id", "")
        table.add_row(str(idx), oid, otype, display_name)

    console.print(table)


@users_app.command("licenses")
def users_licenses(
    user: str = typer.Argument(..., help="UPN или id пользователя."),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(
        ...,
        "--client-secret",
        prompt=True,
        hide_input=True,
        envvar="GRAPH_CLIENT_SECRET",
    ),
):
    """
    Лицензии пользователя.
    """
    console.rule(f"[bold]Лицензии пользователя {user}[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    items = users_module.get_user_license_details(client, user)
    if not items:
        console.print("[yellow]У пользователя не найдено лицензий.[/yellow]")
        raise typer.Exit(code=0)

    table = Table(show_header=True, header_style="bold green")
    table.add_column("Index", style="dim", width=6)
    table.add_column("SKU ID", overflow="fold")
    table.add_column("SKU Part Number", overflow="fold")

    for idx, lic in enumerate(items, start=1):
        sku_id = lic.get("skuId", "")
        sku_part = lic.get("skuPartNumber", "")
        table.add_row(str(idx), str(sku_id), sku_part)

    console.print(table)


# --- GROUPS --- #

groups_app = typer.Typer(help="Операции с группами (AAD / M365 Groups)")
app.add_typer(groups_app, name="groups")


@groups_app.command("list")
def groups_list(
    top: int = typer.Option(25, "--top", help="Сколько групп показать."),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(
        ...,
        "--client-secret",
        prompt=True,
        hide_input=True,
        envvar="GRAPH_CLIENT_SECRET",
    ),
):
    """
    Список групп.
    """
    console.rule("[bold]Группы[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    groups = groups_module.list_groups(client, top=top)
    if not groups:
        console.print("[yellow]Группы не найдены.[/yellow]")
        raise typer.Exit(code=0)

    table = Table(show_header=True, header_style="bold yellow")
    table.add_column("Index", style="dim", width=6)
    table.add_column("Group ID", overflow="fold")
    table.add_column("DisplayName", overflow="fold")
    table.add_column("Mail", overflow="fold")

    for idx, g in enumerate(groups, start=1):
        table.add_row(
            str(idx),
            g.get("id", ""),
            g.get("displayName", ""),
            g.get("mail", ""),
        )

    console.print(table)


@groups_app.command("get")
def groups_get(
    group_id: str = typer.Argument(..., help="ID группы."),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(
        ...,
        "--client-secret",
        prompt=True,
        hide_input=True,
        envvar="GRAPH_CLIENT_SECRET",
    ),
):
    """
    Получить одну группу по ID.
    """
    console.rule(f"[bold]Группа {group_id}[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    data = groups_module.get_group(client, group_id)
    print_json(data=data)


@groups_app.command("members")
def groups_members(
    group_id: str = typer.Argument(..., help="ID группы."),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(
        ...,
        "--client-secret",
        prompt=True,
        hide_input=True,
        envvar="GRAPH_CLIENT_SECRET",
    ),
):
    """
    Состав группы.
    """
    console.rule(f"[bold]Участники группы {group_id}[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    members = groups_module.list_group_members(client, group_id)
    if not members:
        console.print("[yellow]Участники не найдены.[/yellow]")
        raise typer.Exit(code=0)

    table = Table(show_header=True, header_style="bold yellow")
    table.add_column("Index", style="dim", width=6)
    table.add_column("Object ID", overflow="fold")
    table.add_column("Type", overflow="fold")
    table.add_column("DisplayName", overflow="fold")

    for idx, m in enumerate(members, start=1):
        otype = m.get("@odata.type", "")
        oid = m.get("id", "")
        dn = m.get("displayName", "") or m.get("userPrincipalName", "")
        table.add_row(str(idx), oid, otype, dn)

    console.print(table)

@groups_app.command("add-owner")
def groups_add_owner(
    group_id: str = typer.Argument(..., help="ID группы."),
    user_upn: str = typer.Argument(..., help="UPN пользователя, которого сделать owner'ом."),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(
        ...,
        "--client-secret",
        prompt=True,
        hide_input=True,
        envvar="GRAPH_CLIENT_SECRET",
    ),
):
    """
    Добавить owner'а в группу по UPN.
    """
    console.rule("[bold]Добавление owner'а в группу[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    result = groups_module.add_group_owner_by_upn(client, group_id, user_upn)
    console.print("[bold green]Owner добавлен.[/bold green]")
    print_json(data=result)


# --- TEAMS --- #

teams_app = typer.Typer(help="Операции с Teams / каналами")
app.add_typer(teams_app, name="teams")


@teams_app.command("user-joined")
def teams_user_joined(
    user: str = typer.Argument(..., help="UPN или id пользователя."),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(
        ...,
        "--client-secret",
        prompt=True,
        hide_input=True,
        envvar="GRAPH_CLIENT_SECRET",
    ),
):
    """
    Команды (Teams), в которых состоит пользователь.
    """
    console.rule(f"[bold]Teams пользователя {user}[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    teams = teams_module.list_user_joined_teams(client, user)
    if not teams:
        console.print("[yellow]Команды не найдены.[/yellow]")
        raise typer.Exit(code=0)

    table = Table(show_header=True, header_style="bold blue")
    table.add_column("Index", style="dim", width=6)
    table.add_column("Team ID", overflow="fold")
    table.add_column("DisplayName", overflow="fold")
    table.add_column("Description", overflow="fold")

    for idx, t in enumerate(teams, start=1):
        table.add_row(
            str(idx),
            t.get("id", ""),
            t.get("displayName", ""),
            t.get("description", "") or "",
        )

    console.print(table)


@teams_app.command("channels")
def teams_channels(
    team_id: str = typer.Argument(..., help="ID команды (Team)."),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(
        ...,
        "--client-secret",
        prompt=True,
        hide_input=True,
        envvar="GRAPH_CLIENT_SECRET",
    ),
):
    """
    Каналы в команде.
    """
    console.rule(f"[bold]Каналы команды {team_id}[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    channels = teams_module.list_team_channels(client, team_id)
    if not channels:
        console.print("[yellow]Каналы не найдены.[/yellow]")
        raise typer.Exit(code=0)

    table = Table(show_header=True, header_style="bold blue")
    table.add_column("Index", style="dim", width=6)
    table.add_column("Channel ID", overflow="fold")
    table.add_column("DisplayName", overflow="fold")
    table.add_column("Type", overflow="fold")

    for idx, ch in enumerate(channels, start=1):
        mt_raw = (ch.get("membershipType") or "").lower()
        # нормализуем shared / unknownFutureValue
        if mt_raw == "unknownfuturevalue":
            mt_display = "shared"
        else:
            mt_display = mt_raw or "-"

        table.add_row(
            str(idx),
            ch.get("id", ""),
            ch.get("displayName", ""),
            mt_display,
        )

    console.print(table)


@teams_app.command("members")
def teams_members(
    team_id: str = typer.Argument(..., help="ID команды (Team)."),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(
        ...,
        "--client-secret",
        prompt=True,
        hide_input=True,
        envvar="GRAPH_CLIENT_SECRET",
    ),
):
    """
    Список участников Team.
    """
    console.rule(f"[bold]Участники Team {team_id}[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    members = teams_module.list_team_members(client, team_id)
    if not members:
        console.print("[yellow]Участники не найдены.[/yellow]")
        raise typer.Exit(code=0)

    table = Table(show_header=True, header_style="bold blue")
    table.add_column("Index", style="dim", width=6)
    table.add_column("Membership ID", overflow="fold")
    table.add_column("DisplayName", overflow="fold")
    table.add_column("Email", overflow="fold")
    table.add_column("Roles", overflow="fold")

    for idx, m in enumerate(members, start=1):
        roles = ", ".join(m.get("roles", []))
        table.add_row(
            str(idx),
            m.get("id", ""),
            m.get("displayName", ""),
            m.get("email", "") or "",
            roles,
        )

    console.print(table)


@teams_app.command("channel-members")
def teams_channel_members(
    team_id: str = typer.Argument(..., help="ID команды (Team)."),
    channel_id: str = typer.Argument(..., help="ID канала."),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(
        ...,
        "--client-secret",
        prompt=True,
        hide_input=True,
        envvar="GRAPH_CLIENT_SECRET",
    ),
):
    """
    Список участников канала.
    """
    console.rule(f"[bold]Участники канала {channel_id}[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    members = teams_module.list_channel_members(client, team_id, channel_id)
    if not members:
        console.print("[yellow]Участники не найдены.[/yellow]")
        raise typer.Exit(code=0)

    table = Table(show_header=True, header_style="bold blue")
    table.add_column("Index", style="dim", width=6)
    table.add_column("Membership ID", overflow="fold")
    table.add_column("DisplayName", overflow="fold")
    table.add_column("Email", overflow="fold")
    table.add_column("Roles", overflow="fold")

    for idx, m in enumerate(members, start=1):
        roles = ", ".join(m.get("roles", []))
        table.add_row(
            str(idx),
            m.get("id", ""),
            m.get("displayName", ""),
            m.get("email", "") or "",
            roles,
        )

    console.print(table)


@teams_app.command("add-member")
def teams_add_member(
    team_id: str = typer.Argument(..., help="ID команды (Team)."),
    user_upn: str = typer.Argument(..., help="UPN участника, например user2@example.com."),
    owner: bool = typer.Option(
        False,
        "--owner",
        help="Добавить как owner команды.",
    ),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(
        ...,
        "--client-secret",
        prompt=True,
        hide_input=True,
        envvar="GRAPH_CLIENT_SECRET",
    ),
):
    """
    Добавить участника в команду (Team).
    """
    console.rule("[bold]Добавление участника в Team[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    result = teams_module.add_member_to_team(
        client=client,
        team_id=team_id,
        user_upn=user_upn,
        as_owner=owner,
    )
    console.print("[bold green]Участник добавлен в Team.[/bold green]")
    print_json(data=result)


@teams_app.command("remove-member")
def teams_remove_member(
    team_id: str = typer.Argument(..., help="ID команды (Team)."),
    user_upn: str = typer.Argument(..., help="UPN/email участника."),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(
        ...,
        "--client-secret",
        prompt=True,
        hide_input=True,
        envvar="GRAPH_CLIENT_SECRET",
    ),
):
    """
    Удалить участника из Team по UPN/email.
    """
    console.rule("[bold]Удаление участника из Team[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    teams_module.remove_member_from_team(client, team_id, user_upn)
    console.print("[bold green]Участник удалён из Team.[/bold green]")


@teams_app.command("add-channel-member")
def teams_add_channel_member(
    team_id: str = typer.Argument(..., help="ID команды (Team)."),
    channel_id: str = typer.Argument(..., help="ID канала."),
    user_upn: str = typer.Argument(..., help="UPN участника."),
    owner: bool = typer.Option(
        False,
        "--owner",
        help="Добавить как owner канала (для приватных/общих каналов).",
    ),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(
        ...,
        "--client-secret",
        prompt=True,
        hide_input=True,
        envvar="GRAPH_CLIENT_SECRET",
    ),
):
    """
    Добавить участника в канал (обычно private/shared).
    """
    console.rule("[bold]Добавление участника в канал[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    result = teams_module.add_member_to_channel(
        client=client,
        team_id=team_id,
        channel_id=channel_id,
        user_upn=user_upn,
        as_owner=owner,
    )
    console.print("[bold green]Участник добавлен в канал.[/bold green]")
    print_json(data=result)


@teams_app.command("remove-channel-member")
def teams_remove_channel_member(
    team_id: str = typer.Argument(..., help="ID команды (Team)."),
    channel_id: str = typer.Argument(..., help="ID канала."),
    user_upn: str = typer.Argument(..., help="UPN/email участника."),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(
        ...,
        "--client-secret",
        prompt=True,
        hide_input=True,
        envvar="GRAPH_CLIENT_SECRET",
    ),
):
    """
    Удалить участника из канала по UPN/email.
    """
    console.rule("[bold]Удаление участника из канала[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    teams_module.remove_member_from_channel(client, team_id, channel_id, user_upn)
    console.print("[bold green]Участник удалён из канала.[/bold green]")


@teams_app.command("create-group")
def teams_create_group(
    display_name: str = typer.Argument(...),
    description: str = typer.Argument(...),
    mail_nickname: str = typer.Argument(...),
    owner_upn: str = typer.Option(
        None,
        "--owner",
        help="UPN пользователя, которого сразу сделать owner'ом группы/Team.",
    ),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(
        ...,
        "--client-secret",
        prompt=True,
        hide_input=True,
        envvar="GRAPH_CLIENT_SECRET",
    ),
):
    """
    Создать Microsoft 365 группу (Unified).
    Можно сразу указать owner'а.
    """
    console.rule("[bold green]Создание M365 группы[/bold green]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    result = teams_create_module.create_m365_group(
        client,
        display_name,
        description,
        mail_nickname,
        owner_upn=owner_upn,
    )

    console.print("[green]Группа создана[/green]")
    print_json(data=result)


@teams_app.command("teamify")
def teams_teamify(
    group_id: str = typer.Argument(..., help="ID группы, которую превращаем в Team"),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(..., "--client-secret", prompt=True, hide_input=True, envvar="GRAPH_CLIENT_SECRET"),
):
    console.rule("[bold blue]Создание Team из группы[/bold blue]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    result = teams_create_module.create_team_from_group(client, group_id)

    console.print("[bold green]Team создана[/bold green]")
    print_json(data=result)


@teams_app.command("create-channel")
def teams_create_channel(
    team_id: str = typer.Argument(..., help="ID Team (группа, у которой уже есть Team)."),
    display_name: str = typer.Argument(..., help="Имя канала."),
    description: str = typer.Argument(..., help="Описание канала."),
    channel_type: str = typer.Option(
        "standard",
        "--type",
        help="Тип канала: standard | private | shared",
    ),
    owner_upn: str = typer.Option(
        None,
        "--owner",
        help="UPN владельца для private/shared канала (обязателен для них).",
    ),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(
        ...,
        "--client-secret",
        prompt=True,
        hide_input=True,
        envvar="GRAPH_CLIENT_SECRET",
    ),
):
    console.rule("[bold cyan]Создание канала[/bold cyan]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    result = teams_create_module.create_channel(
        client,
        team_id,
        display_name,
        description,
        channel_type,
        owner_upn=owner_upn,
    )

    console.print("[green]Канал создан[/green]")
    print_json(data=result)

# --- ONEDRIVE --- #

onedrive_app = typer.Typer(help="OneDrive операции")
app.add_typer(onedrive_app, name="onedrive")

@onedrive_app.command("list-root")
def onedrive_list_root(
    user: str = typer.Argument(..., help="UPN/id пользователя."),
    top: int = typer.Option(50, "--top"),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(..., "--client-secret", prompt=True, hide_input=True, envvar="GRAPH_CLIENT_SECRET"),
):
    console.rule(f"[bold]OneDrive root {user}[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    items = onedrive_module.list_root(client, user, top=top)
    if not items:
        console.print("[yellow]Пусто.[/yellow]")
        raise typer.Exit(0)

    table = Table(show_header=True, header_style="bold cyan")
    table.add_column("Index", style="dim", width=6)
    table.add_column("Item ID", overflow="fold")
    table.add_column("Name", overflow="fold")
    table.add_column("Type", overflow="fold")

    for idx, it in enumerate(items, start=1):
        if "folder" in it:
            itype = "folder"
        elif "file" in it:
            itype = "file"
        else:
            itype = "other"

        table.add_row(
            str(idx),
            it.get("id", ""),
            it.get("name", ""),
            itype,
        )

    console.print(table)


@onedrive_app.command("children")
def onedrive_children(
    user: str = typer.Argument(...),
    item_id: str = typer.Argument(...),
    top: int = typer.Option(50, "--top"),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(..., "--client-secret", prompt=True, hide_input=True, envvar="GRAPH_CLIENT_SECRET"),
):
    console.rule(f"[bold]Children of item {item_id} ({user})[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    items = onedrive_module.list_children(client, user, item_id, top=top)
    if not items:
        console.print("[yellow]Пусто.[/yellow]")
        raise typer.Exit(0)

    table = Table(show_header=True, header_style="bold cyan")
    table.add_column("Index", style="dim", width=6)
    table.add_column("Item ID", overflow="fold")
    table.add_column("Name", overflow="fold")
    table.add_column("Type", overflow="fold")

    for idx, it in enumerate(items, start=1):
        if "folder" in it:
            itype = "folder"
        elif "file" in it:
            itype = "file"
        else:
            itype = "other"

        table.add_row(
            str(idx),
            it.get("id", ""),
            it.get("name", ""),
            itype,
        )

    console.print(table)


@onedrive_app.command("download")
def onedrive_download(
    user: str = typer.Argument(...),
    item_id: str = typer.Argument(...),
    dest: str = typer.Argument(..., help="Локальный путь, куда сохранить файл."),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(..., "--client-secret", prompt=True, hide_input=True, envvar="GRAPH_CLIENT_SECRET"),
):
    console.rule("[bold]OneDrive download[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    onedrive_module.download_item(client, user, item_id, dest)
    console.print(f"[green]Файл скачан в {dest}[/green]")


@onedrive_app.command("upload")
def onedrive_upload(
    user: str = typer.Argument(...),
    local_path: str = typer.Argument(..., help="Локальный файл."),
    remote_path: str = typer.Argument(..., help="Путь в OneDrive, напр. myfolder/file.txt"),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(..., "--client-secret", prompt=True, hide_input=True, envvar="GRAPH_CLIENT_SECRET"),
):
    console.rule("[bold]OneDrive upload[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    result = onedrive_module.upload_file_to_path(client, user, local_path, remote_path)
    console.print("[green]Файл загружен[/green]")
    print_json(data=result)


@onedrive_app.command("delete")
def onedrive_delete(
    user: str = typer.Argument(...),
    item_id: str = typer.Argument(...),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(..., "--client-secret", prompt=True, hide_input=True, envvar="GRAPH_CLIENT_SECRET"),
):
    console.rule("[bold]OneDrive delete[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    onedrive_module.delete_item(client, user, item_id)
    console.print("[green]Элемент удалён[/green]")


@onedrive_app.command("share-link")
def onedrive_share_link(
    user: str = typer.Argument(...),
    item_id: str = typer.Argument(...),
    link_type: str = typer.Option("view", "--type", help="view|edit|embed"),
    scope: str = typer.Option("organization", "--scope", help="organization|anonymous"),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(..., "--client-secret", prompt=True, hide_input=True, envvar="GRAPH_CLIENT_SECRET"),
):
    console.rule("[bold]OneDrive create link[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    result = onedrive_module.create_link(client, user, item_id, link_type, scope)
    console.print("[green]Ссылка создана[/green]")
    print_json(data=result)


@onedrive_app.command("search")
def onedrive_search(
    user: str = typer.Argument(...),
    query: str = typer.Argument(...),
    top: int = typer.Option(25, "--top"),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(..., "--client-secret", prompt=True, hide_input=True, envvar="GRAPH_CLIENT_SECRET"),
):
    console.rule(f"[bold]OneDrive search: {query}[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    items = onedrive_module.search_files(client, user, query, top=top)
    print_json(data=items)


@onedrive_app.command("clone-root")
def onedrive_clone_root(
    source_user: str = typer.Argument(..., help="UPN источника"),
    target_user: str = typer.Argument(..., help="UPN назначения"),
    overwrite: bool = typer.Option(False, "--overwrite", help="Перезаписывать файлы в целевом OneDrive"),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(..., "--client-secret", prompt=True, hide_input=True, envvar="GRAPH_CLIENT_SECRET"),
):
    """
    Клонирует все файлы из корня OneDrive source_user в корень target_user (только верхний уровень).
    """
    console.rule("[bold]OneDrive clone root[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    result = onedrive_module.clone_root(client, source_user, target_user, overwrite=overwrite)
    print_json(data=result)


# --- SHAREPOINT --- #

sharepoint_app = typer.Typer(help="SharePoint сайты и файлы")
app.add_typer(sharepoint_app, name="sp")

@sharepoint_app.command("sites")
def sp_sites(
    search: str = typer.Option("", "--search", help="Поиск по названию сайта"),
    top: int = typer.Option(20, "--top"),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(..., "--client-secret", prompt=True, hide_input=True, envvar="GRAPH_CLIENT_SECRET"),
):
    console.rule("[bold]SharePoint sites[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    sites = sharepoint_module.list_sites(client, search=search, top=top)
    if not sites:
        console.print("[yellow]Сайты не найдены.[/yellow]")
        raise typer.Exit(0)

    table = Table(show_header=True, header_style="bold green")
    table.add_column("Index", style="dim", width=6)
    table.add_column("Site ID", overflow="fold")
    table.add_column("Name", overflow="fold")
    table.add_column("WebUrl", overflow="fold")

    for idx, s in enumerate(sites, start=1):
        table.add_row(
            str(idx),
            s.get("id", ""),
            s.get("displayName", ""),
            s.get("webUrl", ""),
        )

    console.print(table)


@sharepoint_app.command("root")
def sp_root(
    site_id: str = typer.Argument(...),
    top: int = typer.Option(50, "--top"),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(..., "--client-secret", prompt=True, hide_input=True, envvar="GRAPH_CLIENT_SECRET"),
):
    console.rule(f"[bold]SP drive root {site_id}[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    items = sharepoint_module.list_site_root(client, site_id, top=top)
    if not items:
        console.print("[yellow]Пусто.[/yellow]")
        raise typer.Exit(0)

    table = Table(show_header=True, header_style="bold green")
    table.add_column("Index", style="dim", width=6)
    table.add_column("Item ID", overflow="fold")
    table.add_column("Name", overflow="fold")
    table.add_column("Type", overflow="fold")

    for idx, it in enumerate(items, start=1):
        if "folder" in it:
            itype = "folder"
        elif "file" in it:
            itype = "file"
        else:
            itype = "other"
        table.add_row(
            str(idx),
            it.get("id", ""),
            it.get("name", ""),
            itype,
        )

    console.print(table)


@sharepoint_app.command("download")
def sp_download(
    site_id: str = typer.Argument(...),
    item_id: str = typer.Argument(...),
    dest: str = typer.Argument(..., help="Локальный путь"),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(..., "--client-secret", prompt=True, hide_input=True, envvar="GRAPH_CLIENT_SECRET"),
):
    console.rule("[bold]SP download[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    sharepoint_module.download_site_item(client, site_id, item_id, dest)
    console.print(f"[green]Файл скачан в {dest}[/green]")


@sharepoint_app.command("upload")
def sp_upload(
    site_id: str = typer.Argument(...),
    local_path: str = typer.Argument(...),
    remote_path: str = typer.Argument(...),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(..., "--client-secret", prompt=True, hide_input=True, envvar="GRAPH_CLIENT_SECRET"),
):
    console.rule("[bold]SP upload[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    result = sharepoint_module.upload_site_file(client, site_id, local_path, remote_path)
    console.print("[green]Файл загружен[/green]")
    print_json(data=result)


@sharepoint_app.command("delete")
def sp_delete(
    site_id: str = typer.Argument(...),
    item_id: str = typer.Argument(...),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(..., "--client-secret", prompt=True, hide_input=True, envvar="GRAPH_CLIENT_SECRET"),
):
    console.rule("[bold]SP delete[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    sharepoint_module.delete_site_item(client, site_id, item_id)
    console.print("[green]Элемент удалён[/green]")


@sharepoint_app.command("share-link")
def sp_share_link(
    site_id: str = typer.Argument(...),
    item_id: str = typer.Argument(...),
    link_type: str = typer.Option("view", "--type"),
    scope: str = typer.Option("organization", "--scope"),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(..., "--client-secret", prompt=True, hide_input=True, envvar="GRAPH_CLIENT_SECRET"),
):
    console.rule("[bold]SP create link[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    result = sharepoint_module.create_site_link(client, site_id, item_id, link_type, scope)
    console.print("[green]Ссылка создана[/green]")
    print_json(data=result)



# --- LICENSING --- #

licensing_app = typer.Typer(help="Лицензии и SKU")
app.add_typer(licensing_app, name="licensing")


@licensing_app.command("skus")
def licensing_skus(
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(
        ...,
        "--client-secret",
        prompt=True,
        hide_input=True,
        envvar="GRAPH_CLIENT_SECRET",
    ),
):
    """
    Показать список лицензий (SKU) тенанта.
    """
    console.rule("[bold]Subscribed SKUs[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    skus = licensing_module.list_skus(client)
    if not skus:
        console.print("[yellow]SKU не найдены.[/yellow]")
        raise typer.Exit(code=0)

    table = Table(show_header=True, header_style="bold green")
    table.add_column("Index", style="dim", width=6)
    table.add_column("SKU ID", overflow="fold")
    table.add_column("SKU Part Number", overflow="fold")
    table.add_column("Consumed", overflow="fold")
    table.add_column("Total", overflow="fold")

    for idx, s in enumerate(skus, start=1):
        prepaid = s.get("prepaidUnits") or {}
        total = (
            prepaid.get("enabled", 0)
            + prepaid.get("suspended", 0)
            + prepaid.get("warning", 0)
        )
        table.add_row(
            str(idx),
            str(s.get("skuId", "")),
            s.get("skuPartNumber", ""),
            str(s.get("consumedUnits", 0)),
            str(total),
        )

    console.print(table)


@licensing_app.command("assign")
def licensing_assign(
    user: str = typer.Argument(..., help="UPN или ID пользователя."),
    add: list[str] = typer.Option(
        None,
        "--add",
        help="SKU ID (GUID) для выдачи. Можно указать несколько раз.",
    ),
    remove: list[str] = typer.Option(
        None,
        "--remove",
        help="SKU ID (GUID) для снятия. Можно указать несколько раз.",
    ),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(
        ...,
        "--client-secret",
        prompt=True,
        hide_input=True,
        envvar="GRAPH_CLIENT_SECRET",
    ),
):
    """
    Выдать/снять лицензии пользователю.
    Пример:
      python main.py licensing assign user1@example.com --add <SKU_GUID> --remove <SKU_GUID>
    """
    console.rule(f"[bold]Изменение лицензий пользователя {user}[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    add_list = list(add or [])
    remove_list = list(remove or [])

    if not add_list and not remove_list:
        console.print(
            "[red]Нечего делать: не указаны ни --add, ни --remove SKU ID.[/red]"
        )
        raise typer.Exit(code=1)

    result = licensing_module.assign_licenses(
        client=client,
        user=user,
        add_sku_ids=add_list,
        remove_sku_ids=remove_list,
    )

    console.print("[bold green]Лицензии обновлены.[/bold green]")
    print_json(data=result)


# --- MAIL --- #

mail_app = typer.Typer(help="Почта (Exchange Online)")
app.add_typer(mail_app, name="mail")


@mail_app.command("list")
def mail_list(
    user: str = typer.Argument(..., help="UPN или id пользователя."),
    top: int = typer.Option(20, "--top", help="Сколько писем показать."),
    folder: str = typer.Option(
        "inbox",
        "--folder",
        help="Системная папка: inbox, sentitems, drafts и т.п.",
    ),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(
        ...,
        "--client-secret",
        prompt=True,
        hide_input=True,
        envvar="GRAPH_CLIENT_SECRET",
    ),
):
    """
    Список последних писем пользователя из папки.
    """
    console.rule(f"[bold]Письма {user} ({folder})[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    messages = mail_module.list_messages(client, user, top=top, folder=folder)
    if not messages:
        console.print("[yellow]Писем не найдено.[/yellow]")
        raise typer.Exit(code=0)

    table = Table(show_header=True, header_style="bold orange1")
    table.add_column("Index", style="dim", width=6)
    table.add_column("Subject", overflow="fold")
    table.add_column("From", overflow="fold")
    table.add_column("Received", overflow="fold")
    table.add_column("IsRead", overflow="fold")

    for idx, m in enumerate(messages, start=1):
        frm = m.get("from") or {}
        addr = (frm.get("emailAddress") or {}).get("address", "")
        table.add_row(
            str(idx),
            m.get("subject", "") or "",
            addr,
            m.get("receivedDateTime", ""),
            str(m.get("isRead", False)),
        )

    console.print(table)


@mail_app.command("send")
def mail_send(
    user: str = typer.Argument(
        ..., help="UPN/id отправителя (от чьего имени шлём письмо)."
    ),
    subject: str = typer.Argument(..., help="Тема письма."),
    body: str = typer.Argument(..., help="Текст письма."),
    to: list[str] = typer.Argument(
        ..., help="Список получателей (email). Можно несколько аргументов подряд."
    ),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(
        ...,
        "--client-secret",
        prompt=True,
        hide_input=True,
        envvar="GRAPH_CLIENT_SECRET",
    ),
):
    """
    Отправить письмо от имени пользователя.
    Пример:
      python main.py mail send user1@example.com "Subj" "Body text" user2@example.com user3@example.com
    """
    console.rule(f"[bold]Отправка письма от {user}[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    mail_module.send_mail(client, user, subject, body, to)
    console.print("[bold green]Письмо отправлено.[/bold green]")


# --- CALENDAR --- #

calendar_app = typer.Typer(help="Календарь (Exchange Online)")
app.add_typer(calendar_app, name="calendar")


@calendar_app.command("list")
def calendar_list(
    user: str = typer.Argument(..., help="UPN или id пользователя."),
    top: int = typer.Option(20, "--top", help="Сколько событий показать."),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(
        ...,
        "--client-secret",
        prompt=True,
        hide_input=True,
        envvar="GRAPH_CLIENT_SECRET",
    ),
):
    """
    Список последних событий пользователя.
    """
    console.rule(f"[bold]События {user}[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    events = calendar_module.list_events(client, user, top=top)
    if not events:
        console.print("[yellow]События не найдены.[/yellow]")
        raise typer.Exit(code=0)

    table = Table(show_header=True, header_style="bold purple")
    table.add_column("Index", style="dim", width=6)
    table.add_column("Subject", overflow="fold")
    table.add_column("Start", overflow="fold")
    table.add_column("End", overflow="fold")
    table.add_column("Location", overflow="fold")

    for idx, ev in enumerate(events, start=1):
        start = (ev.get("start") or {}).get("dateTime", "")
        end = (ev.get("end") or {}).get("dateTime", "")
        loc = (ev.get("location") or {}).get("displayName", "")
        table.add_row(
            str(idx),
            ev.get("subject", "") or "",
            start,
            end,
            loc or "",
        )

    console.print(table)


@calendar_app.command("create")
def calendar_create(
    user: str = typer.Argument(..., help="UPN/id пользователя, чей календарь."),
    subject: str = typer.Argument(..., help="Тема встречи."),
    body: str = typer.Argument(..., help="Текст приглашения."),
    start: str = typer.Argument(
        ..., help="Начало в ISO формате, например 2025-12-11T10:00:00"
    ),
    end: str = typer.Argument(
        ..., help="Окончание в ISO формате, например 2025-12-11T11:00:00"
    ),
    timezone: str = typer.Option(
        "UTC",
        "--tz",
        help="Часовой пояс, напр. UTC или Russian Standard Time",
    ),
    attendee: list[str] = typer.Option(
        None,
        "--to",
        help="Участник встречи (email). Можно указать несколько раз.",
    ),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(
        ...,
        "--client-secret",
        prompt=True,
        hide_input=True,
        envvar="GRAPH_CLIENT_SECRET",
    ),
):
    """
    Создать событие в календаре пользователя.
    """
    console.rule(f"[bold]Создание события в календаре {user}[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    attendees = list(attendee or [])

    result = calendar_module.create_event(
        client=client,
        user=user,
        subject=subject,
        body_text=body,
        start_iso=start,
        end_iso=end,
        timezone=timezone,
        attendees=attendees,
    )

    console.print("[bold green]Событие создано.[/bold green]")
    print_json(data=result)

@calendar_app.command("quick-create")
def calendar_quick_create(
    payload: str = typer.Argument(
        ...,
        help="JSON строка с параметрами: "
             '{"user": "...", "subject": "...", "body": "...", '
             '"start": "ISO", "end": "ISO", "timezone":"UTC", "to":["a@b.com"]}',
    ),
    tenant_id: str = typer.Option(..., "--tenant-id", envvar="GRAPH_TENANT_ID"),
    client_id: str = typer.Option(..., "--client-id", envvar="GRAPH_CLIENT_ID"),
    client_secret: str = typer.Option(
        ...,
        "--client-secret",
        prompt=True,
        hide_input=True,
        envvar="GRAPH_CLIENT_SECRET",
    ),
):
    """
    Создать событие в календаре через один JSON-аргумент.
    Удобно для скриптов и интеграций.
    """
    import json

    console.rule("[bold]Quick Create Calendar Event[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    try:
        data = json.loads(payload)
    except json.JSONDecodeError as e:
        console.print("[red]Невозможно прочитать JSON[/red]")
        console.print(str(e))
        raise typer.Exit(code=1)

    required = ["user", "subject", "body", "start", "end"]
    for key in required:
        if key not in data:
            console.print(f"[red]Отсутствует обязательный параметр: {key}[/red]")
            raise typer.Exit(code=1)

    timezone = data.get("timezone", "UTC")
    attendees = data.get("to", [])

    result = calendar_module.create_event(
        client=client,
        user=data["user"],
        subject=data["subject"],
        body_text=data["body"],
        start_iso=data["start"],
        end_iso=data["end"],
        timezone=timezone,
        attendees=attendees,
    )

    console.print("[bold green]Событие создано успешно[/bold green]")
    print_json(data=result)


# --- RAW --- #


@app.command("raw")
def raw(
    method: str = typer.Argument(
        ...,
        help="HTTP метод: GET, POST, PATCH, DELETE",
    ),
    path: str = typer.Argument(
        ...,
        help="Путь Graph, например /users или /chats?$top=5. Можно и полный URL.",
    ),
    body: str = typer.Option(
        None,
        "--body",
        "-b",
        help="JSON-тело для POST/PATCH/PUT в виде строки. Например: "
             "'{\"topic\": \"New topic\"}'",
    ),
    tenant_id: str = typer.Option(
        ...,
        "--tenant-id",
        envvar="GRAPH_TENANT_ID",
    ),
    client_id: str = typer.Option(
        ...,
        "--client-id",
        envvar="GRAPH_CLIENT_ID",
    ),
    client_secret: str = typer.Option(
        ...,
        "--client-secret",
        prompt=True,
        hide_input=True,
        envvar="GRAPH_CLIENT_SECRET",
    ),
):
    """
    Универсальный низкоуровневый вызов Graph API.
    Удобно, когда нет отдельной команды, но нужно что-то быстро проверить.
    """
    console.rule("[bold]RAW Graph запрос[/bold]")
    client = build_graph_client(tenant_id, client_id, client_secret)

    method_upper = method.upper()
    if method_upper not in {"GET", "POST", "PATCH", "DELETE"}:
        console.print("[red]Поддерживаются только GET, POST, PATCH, DELETE[/red]")
        raise typer.Exit(code=1)

    json_body = None
    if body:
        try:
            json_body = json.loads(body)
        except json.JSONDecodeError as e:
            console.print("[red]Не удалось распарсить JSON из --body[/red]")
            console.print(str(e))
            raise typer.Exit(code=1)

    try:
        if method_upper == "GET":
            result = client.get(path)
        elif method_upper == "POST":
            result = client.post(path, json=json_body)
        elif method_upper == "PATCH":
            result = client.patch(path, json=json_body)
        elif method_upper == "DELETE":
            result = client.delete(path)
        else:
            raise RuntimeError("Неподдерживаемый метод")
    except RuntimeError as e:
        console.print("[red]Ошибка при запросе к Graph[/red]")
        console.print(str(e))
        raise typer.Exit(code=1)

    if isinstance(result, (dict, list)):
        print_json(data=result)
    else:
        console.print(result)


if __name__ == "__main__":
    app()
