"""Microsoft 365 mail MCP server (stdio). Delegated auth via MSAL device-code flow.

Sending is two-step only: create_draft / reply_draft, then send_draft.
Tokens are never logged or returned.
"""
from __future__ import annotations

import os
import re
import threading
from pathlib import Path
from typing import Any
from urllib.parse import quote

import httpx
import msal
from mcp.server.mcpserver import MCPServer
from mcp.server.mcpserver.exceptions import ToolError

GRAPH = "https://graph.microsoft.com/v1.0"
SCOPES = [
    "Mail.ReadWrite", "Mail.Send", "Mail.ReadWrite.Shared", "Mail.Send.Shared", "User.Read",
]  # offline_access is added automatically by MSAL (it's a reserved scope)
CACHE_PATH = Path(os.path.expanduser(os.environ.get("M365_TOKEN_CACHE", "~/.config/grok-bot-m365/token_cache.json")))


class MailError(ToolError):
    """Shown verbatim to the MCP client as the tool error."""


# ---------------------------------------------------------------- auth
class Auth:
    def __init__(self) -> None:
        self._app: msal.PublicClientApplication | None = None
        self.cache = msal.SerializableTokenCache()
        if CACHE_PATH.exists():
            self.cache.deserialize(CACHE_PATH.read_text())
        self._flow: dict | None = None
        self._flow_result: dict | None = None
        self._thread: threading.Thread | None = None
        self._lock = threading.Lock()

    @property
    def app(self) -> msal.PublicClientApplication:
        if self._app is None:
            cid = os.environ.get("M365_CLIENT_ID")
            if not cid:
                raise MailError("M365_CLIENT_ID env var is not set. See SETUP.md.")
            tenant = os.environ.get("M365_TENANT_ID", "organizations")
            self._app = msal.PublicClientApplication(
                cid, authority=f"https://login.microsoftonline.com/{tenant}", token_cache=self.cache
            )
        return self._app

    def _save(self) -> None:
        with self._lock:
            if self.cache.has_state_changed:
                CACHE_PATH.parent.mkdir(parents=True, exist_ok=True, mode=0o700)
                CACHE_PATH.touch(mode=0o600, exist_ok=True)
                os.chmod(CACHE_PATH, 0o600)
                CACHE_PATH.write_text(self.cache.serialize())
                os.chmod(CACHE_PATH, 0o600)
                self.cache.has_state_changed = False

    def token(self) -> str:
        accounts = self.app.get_accounts()
        if accounts:
            res = self.app.acquire_token_silent(SCOPES, account=accounts[0])
            self._save()
            if res and "access_token" in res:
                return res["access_token"]
        raise MailError("Not signed in to Microsoft 365. Call auth_start, then auth_status.")

    def start(self) -> dict:
        flow = self.app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise MailError(f"Could not start device flow: {flow.get('error_description', flow)}")
        self._flow, self._flow_result = flow, None

        def worker(f=flow):
            r = self.app.acquire_token_by_device_flow(f)  # blocks until done/expired
            self._save()
            if self._flow is f:
                self._flow_result = r

        self._thread = threading.Thread(target=worker, daemon=True)
        self._thread.start()
        return {
            "verification_uri": flow["verification_uri"],
            "user_code": flow["user_code"],
            "expires_in_seconds": flow.get("expires_in"),
            "message": flow.get("message"),
        }

    def status(self) -> dict:
        r = self._flow_result
        if r is not None:
            self._flow = None
            if "access_token" in r:
                self._flow_result = None
                claims = r.get("id_token_claims") or {}
                return {"signed_in": True, "account": claims.get("preferred_username")}
            return {"signed_in": False, "error": r.get("error"), "detail": r.get("error_description")}
        if self._flow is not None:
            return {"signed_in": False, "pending": True, "user_code": self._flow["user_code"],
                    "verification_uri": self._flow["verification_uri"]}
        accounts = self.app.get_accounts()
        if accounts:
            try:
                self.token()
                return {"signed_in": True, "account": accounts[0].get("username")}
            except MailError:
                pass
        return {"signed_in": False, "message": "Not signed in. Call auth_start."}


auth = Auth()


# ---------------------------------------------------------------- graph
def _base(mailbox: str | None) -> str:
    return f"/users/{quote(mailbox)}" if mailbox else "/me"


def _explain(resp: httpx.Response, mailbox: str | None, sending: bool) -> str:
    try:
        err = resp.json().get("error", {})
    except Exception:
        err = {}
    code, msg = err.get("code", ""), err.get("message", resp.text[:300])
    base = f"Graph {resp.status_code} {code}: {msg}"
    lc = f"{code} {msg}".lower()
    if "sendas" in lc or "send as" in lc or "send on behalf" in lc or "invalidfrom" in lc or (sending and resp.status_code == 403):
        return base + (
            " | Send was rejected for the From address. If sending from an alias: the tenant needs "
            "`Set-OrganizationConfig -SendFromAliasEnabled $true` (can take hours to apply) and the address "
            "must be an SMTP proxy address of your mailbox (see list_aliases). If sending from a shared "
            "mailbox: your user needs Send As (Add-RecipientPermission -AccessRights SendAs) on it.")
    if resp.status_code == 403 or code in ("ErrorAccessDenied", "AccessDenied"):
        return base + (
            f" | Access denied to mailbox '{mailbox or 'me'}'. For a shared mailbox your user needs Full Access "
            "(Add-MailboxPermission -AccessRights FullAccess) and, to send, Send As; the app also needs "
            "Mail.ReadWrite.Shared / Mail.Send.Shared with admin consent.")
    if resp.status_code == 404 and mailbox:
        return base + f" | Mailbox '{mailbox}' or item not found, or you lack Full Access to it."
    if resp.status_code == 401:
        return base + " | Token rejected; call auth_start to sign in again."
    return base


def graph(method: str, path: str, *, mailbox: str | None = None, sending: bool = False,
          params: dict | None = None, json: Any = None, headers: dict | None = None) -> Any:
    h = {"Authorization": f"Bearer {auth.token()}"}
    h.update(headers or {})
    url = path if path.startswith("http") else GRAPH + path
    with httpx.Client(timeout=60) as c:
        resp = c.request(method, url, params=params, json=json, headers=h)
    if resp.status_code >= 400:
        raise MailError(_explain(resp, mailbox, sending))
    if resp.status_code == 204 or not resp.content:
        return None
    return resp.json()


def _addr(r: dict | None) -> str | None:
    if not r:
        return None
    e = r.get("emailAddress", {})
    return f"{e.get('name')} <{e.get('address')}>" if e.get("name") else e.get("address")


def _recips(v: str | list[str] | None) -> list[dict]:
    if not v:
        return []
    items = re.split(r"[;,]", v) if isinstance(v, str) else v
    return [{"emailAddress": {"address": a.strip()}} for a in items if a and a.strip()]


def _summary(m: dict) -> dict:
    return {
        "id": m.get("id"), "subject": m.get("subject"), "from": _addr(m.get("from")),
        "to": [_addr(r) for r in m.get("toRecipients", [])],
        "cc": [_addr(r) for r in m.get("ccRecipients", [])],
        "received": m.get("receivedDateTime"), "preview": m.get("bodyPreview"), "isRead": m.get("isRead"),
    }


def _readback(m: dict) -> dict:
    return {
        "draft_id": m.get("id"), "from": _addr(m.get("from")),
        "to": [_addr(r) for r in m.get("toRecipients", [])],
        "cc": [_addr(r) for r in m.get("ccRecipients", [])],
        "bcc": [_addr(r) for r in m.get("bccRecipients", [])],
        "subject": m.get("subject"), "body_type": (m.get("body") or {}).get("contentType"),
        "body": (m.get("body") or {}).get("content"), "isDraft": m.get("isDraft"),
    }


TEXT_PREF = {"Prefer": 'outlook.body-content-type="text"'}
SUMMARY_FIELDS = "id,subject,from,toRecipients,ccRecipients,receivedDateTime,bodyPreview,isRead"
DRAFT_FIELDS = "id,subject,from,toRecipients,ccRecipients,bccRecipients,body,isDraft"


# ---------------------------------------------------------------- tools
mcp = MCPServer("grok-bot-m365", instructions=(
    "Microsoft 365 mail. Supports aliases (from_address) and shared mailboxes (mailbox). "
    "There is no direct send: create a draft, show it to the user, and call send_draft only after approval."))


@mcp.tool()
def auth_start() -> dict:
    """Start Microsoft device-code sign-in. Returns a URL and code for the user to enter; then call auth_status."""
    return auth.start()


@mcp.tool()
def auth_status() -> dict:
    """Check/complete sign-in started by auth_start. Returns signed_in true/false (pending while waiting)."""
    return auth.status()


@mcp.tool()
def list_aliases() -> dict:
    """List the signed-in user's primary address and all SMTP aliases (proxyAddresses)."""
    me = graph("GET", "/me", params={"$select": "mail,userPrincipalName,proxyAddresses,otherMails"})
    primary, aliases = me.get("mail"), []
    for p in me.get("proxyAddresses") or []:
        if p.lower().startswith("smtp:"):
            a = p[5:]
            if p.startswith("SMTP:"):
                primary = a
            aliases.append(a)
    configured = _env_list("M365_ALIASES")
    missing = [a for a in configured if a not in {x.lower() for x in aliases}]
    return {"configured_aliases": configured, "configured_missing_from_mailbox": missing,
            "configured_shared_mailboxes": _env_list("M365_SHARED_MAILBOXES"),
            "primary": primary, "userPrincipalName": me.get("userPrincipalName"),
            "smtp_addresses": aliases, "otherMails": me.get("otherMails") or []}


@mcp.tool()
def list_messages(mailbox: str | None = None, folder: str = "inbox", to_address: str | None = None,
                  unread_only: bool = False, query: str | None = None, top: int = 25) -> list[dict]:
    """List messages in a folder (well-known name like inbox/sentitems/drafts, or a folder id).

    to_address: only messages where that address is in To or Cc. Filtering is done CLIENT-SIDE
    (exact, case-insensitive match on toRecipients/ccRecipients), paging through up to 500 recent
    messages, because Graph can't $filter on recipient collections and $search would also match bodies.
    query: Graph $search (KQL, e.g. 'from:bob subject:invoice'); results are relevance-ordered.
    mailbox: shared mailbox address (default: your own mailbox).
    """
    top = max(1, min(int(top), 100))
    params: dict[str, Any] = {"$select": SUMMARY_FIELDS, "$top": 50 if to_address else top}
    if query:
        params["$search"] = '"' + query.replace('"', '\\"') + '"'
        if unread_only:
            params["$filter"] = "isRead eq false"
    else:
        params["$orderby"] = "receivedDateTime desc"
        if unread_only:
            params["$filter"] = "isRead eq false"
    url: str | None = f"{_base(mailbox)}/mailFolders/{quote(folder)}/messages"
    out: list[dict] = []
    scanned = 0
    target = to_address.strip().lower() if to_address else None
    while url and len(out) < top:
        data = graph("GET", url, mailbox=mailbox, params=params)
        params = None  # nextLink already carries params
        for m in data.get("value", []):
            scanned += 1
            if target:
                rs = m.get("toRecipients", []) + m.get("ccRecipients", [])
                if not any((r.get("emailAddress", {}).get("address") or "").lower() == target for r in rs):
                    continue
            out.append(_summary(m))
            if len(out) >= top:
                break
        url = data.get("@odata.nextLink") if target and scanned < 500 else None
    return out


@mcp.tool()
def get_message(id: str, mailbox: str | None = None) -> dict:
    """Get one message with its body as plain text."""
    m = graph("GET", f"{_base(mailbox)}/messages/{quote(id)}", mailbox=mailbox, headers=TEXT_PREF,
              params={"$select": SUMMARY_FIELDS + ",bccRecipients,replyTo,body,hasAttachments,conversationId"})
    d = _summary(m)
    d.update(body=(m.get("body") or {}).get("content"), hasAttachments=m.get("hasAttachments"),
             replyTo=[_addr(r) for r in m.get("replyTo", [])], conversationId=m.get("conversationId"))
    return d


def _env_list(name: str) -> list[str]:
    return [a.strip().lower() for a in os.environ.get(name, "").split(",") if a.strip()]


def _hints(from_address: str | None, mailbox: str | None) -> list[str]:
    """Advisory only: never blocks; flags addresses not in M365_ALIASES / M365_SHARED_MAILBOXES."""
    out = []
    aliases, shared = _env_list("M365_ALIASES"), _env_list("M365_SHARED_MAILBOXES")
    if from_address and aliases and from_address.lower() not in aliases and from_address.lower() not in shared:
        out.append(f"from_address {from_address} is not in M365_ALIASES ({', '.join(aliases)}); "
                   "send may be rejected unless it is a proxy address of the mailbox.")
    if mailbox and shared and mailbox.lower() not in shared:
        out.append(f"mailbox {mailbox} is not in M365_SHARED_MAILBOXES ({', '.join(shared)}).")
    return out


def _from_field(from_address: str | None, mailbox: str | None) -> dict | None:
    addr = from_address or mailbox
    return {"emailAddress": {"address": addr}} if addr else None


@mcp.tool()
def create_draft(to: str | list[str], subject: str, body: str, cc: str | list[str] | None = None,
                 bcc: str | list[str] | None = None, from_address: str | None = None,
                 mailbox: str | None = None, reply_to_message_id: str | None = None,
                 html: bool = False) -> dict:
    """Create (not send) a draft. from_address = send from one of your aliases. mailbox = create the
    draft in a shared mailbox and send as it (needs Full Access + Send As). If reply_to_message_id is
    given, builds a threaded reply to that message (recipients/subject overridden by the given values).
    Returns draft_id and a readback; call send_draft only after the user approves."""
    msg: dict[str, Any] = {
        "subject": subject, "body": {"contentType": "HTML" if html else "Text", "content": body},
        "toRecipients": _recips(to), "ccRecipients": _recips(cc), "bccRecipients": _recips(bcc),
    }
    f = _from_field(from_address, mailbox)
    if f:
        msg["from"] = f
    base = _base(mailbox)
    if reply_to_message_id:
        d = graph("POST", f"{base}/messages/{quote(reply_to_message_id)}/createReply", mailbox=mailbox, json={})
        d = graph("PATCH", f"{base}/messages/{quote(d['id'])}", mailbox=mailbox, json=msg)
    else:
        d = graph("POST", f"{base}/messages", mailbox=mailbox, json=msg)
    r = _readback(d)
    if h := _hints(from_address, mailbox):
        r["warnings"] = h
    return r


@mcp.tool()
def reply_draft(message_id: str, body: str, reply_all: bool = False, from_address: str | None = None,
                mailbox: str | None = None, html: bool = False) -> dict:
    """Create (not send) a reply/reply-all draft to a message; your text is put above the quoted thread.
    Returns draft_id and readback; send with send_draft after approval."""
    base = _base(mailbox)
    action = "createReplyAll" if reply_all else "createReply"
    d = graph("POST", f"{base}/messages/{quote(message_id)}/{action}", mailbox=mailbox, json={})
    did = d["id"]
    quoted = (d.get("body") or {}).get("content", "")
    if html or (d.get("body") or {}).get("contentType", "").lower() == "html":
        content = (body if html else "<p>" + body.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                   .replace("\n", "<br>") + "</p>")
        new_body = {"contentType": "HTML", "content": content + quoted}
    else:
        new_body = {"contentType": "Text", "content": body + "\n\n" + quoted}
    patch: dict[str, Any] = {"body": new_body}
    f = _from_field(from_address, mailbox)
    if f:
        patch["from"] = f
    d = graph("PATCH", f"{base}/messages/{quote(did)}", mailbox=mailbox, json=patch)
    r = _readback(d)
    if h := _hints(from_address, mailbox):
        r["warnings"] = h
    return r


@mcp.tool()
def send_draft(draft_id: str, mailbox: str | None = None) -> dict:
    """Send an existing draft. Only call after the user has approved the draft readback."""
    graph("POST", f"{_base(mailbox)}/messages/{quote(draft_id)}/send", mailbox=mailbox, sending=True)
    return {"sent": True, "draft_id": draft_id}


@mcp.tool()
def move_message(id: str, destination_folder: str, mailbox: str | None = None) -> dict:
    """Move a message to a folder (well-known name like archive/deleteditems, or folder id)."""
    m = graph("POST", f"{_base(mailbox)}/messages/{quote(id)}/move", mailbox=mailbox,
              json={"destinationId": destination_folder})
    return {"moved": True, "new_id": m.get("id")}


@mcp.tool()
def mark_read(id: str, is_read: bool = True, mailbox: str | None = None) -> dict:
    """Mark a message read (or unread with is_read=false)."""
    graph("PATCH", f"{_base(mailbox)}/messages/{quote(id)}", mailbox=mailbox, json={"isRead": is_read})
    return {"id": id, "isRead": is_read}


if __name__ == "__main__":
    mcp.run("stdio")
