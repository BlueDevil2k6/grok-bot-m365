"""Microsoft 365 mail MCP server (stdio). Delegated auth via MSAL device-code flow.

Sending is two-step only: create_draft / reply_draft, then send_draft, which must restate the
draft's recipients and subject. Tokens are never logged or returned; the cache is OS-encrypted.
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
from msal_extensions import FilePersistence, PersistedTokenCache, build_encrypted_persistence
from mcp.server.mcpserver import MCPServer
from mcp.server.mcpserver.exceptions import ToolError

GRAPH = "https://graph.microsoft.com/v1.0"
SCOPES = [
    "Mail.ReadWrite", "Mail.Send", "Mail.ReadWrite.Shared", "Mail.Send.Shared", "User.Read",
]  # offline_access is added automatically by MSAL (it's a reserved scope)
CACHE_PATH = Path(os.path.expanduser(os.environ.get("M365_TOKEN_CACHE", "~/.config/grok-bot-m365/token_cache.bin")))


class MailError(ToolError):
    """Shown verbatim to the MCP client as the tool error."""


# ---------------------------------------------------------------- auth
def _build_cache() -> PersistedTokenCache:
    """Token cache in the OS secret store (DPAPI / Keychain / libsecret). A plaintext 0600 file is
    used only if M365_ALLOW_PLAINTEXT_CACHE is set; the cache persists itself on every change."""
    CACHE_PATH.parent.mkdir(parents=True, exist_ok=True, mode=0o700)
    try:
        persistence = build_encrypted_persistence(str(CACHE_PATH))
    except Exception as e:
        if os.environ.get("M365_ALLOW_PLAINTEXT_CACHE", "").lower() not in ("1", "true", "yes"):
            raise MailError(
                f"No encrypted token storage is available ({type(e).__name__}: {e}). Install and unlock an OS "
                "keyring (on Linux: libsecret + PyGObject), or set M365_ALLOW_PLAINTEXT_CACHE=1 to keep "
                "tokens in a 0600 file.") from e
        persistence = FilePersistence(str(CACHE_PATH))
    return PersistedTokenCache(persistence)


class Auth:
    def __init__(self) -> None:
        self._app: msal.PublicClientApplication | None = None
        self._flow: dict | None = None
        self._flow_result: dict | None = None
        self._thread: threading.Thread | None = None

    @property
    def app(self) -> msal.PublicClientApplication:
        if self._app is None:
            cid = os.environ.get("M365_CLIENT_ID")
            if not cid:
                raise MailError("M365_CLIENT_ID env var is not set. See README.md.")
            tenant = os.environ.get("M365_TENANT_ID", "organizations")
            self._app = msal.PublicClientApplication(
                cid, authority=f"https://login.microsoftonline.com/{tenant}", token_cache=_build_cache()
            )
        return self._app

    def token(self) -> str:
        accounts = self.app.get_accounts()
        if accounts:
            res = self.app.acquire_token_silent(SCOPES, account=accounts[0])
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
def _env_list(name: str) -> list[str]:
    return [a.strip().lower() for a in os.environ.get(name, "").split(",") if a.strip()]


def _seg(v: str) -> str:
    """Encode one URL path segment. '/' is escaped so an id can't walk to another Graph path."""
    if not v or v.strip() in (".", ".."):
        raise MailError(f"Invalid id or name: {v!r}")
    return quote(v, safe="")


def _base(mailbox: str | None) -> str:
    """Path prefix for a mailbox. Shared mailboxes must be listed in M365_SHARED_MAILBOXES."""
    if not mailbox:
        return "/me"
    allowed = _env_list("M365_SHARED_MAILBOXES")
    if mailbox.strip().lower() not in allowed:
        raise MailError(f"Mailbox '{mailbox}' is not in M365_SHARED_MAILBOXES "
                        f"({', '.join(allowed) or 'empty'}); add it there to allow access.")
    return f"/users/{_seg(mailbox.strip())}"


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
    if path.startswith("http") and not path.startswith(GRAPH + "/"):
        raise MailError(f"Refusing to send the token to a non-Graph URL: {path[:100]}")
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
    "There is no direct send: create a draft, show it to the user, and call send_draft only after approval, "
    "passing the draft's exact recipients and subject."))


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
    url: str | None = f"{_base(mailbox)}/mailFolders/{_seg(folder)}/messages"
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
    m = graph("GET", f"{_base(mailbox)}/messages/{_seg(id)}", mailbox=mailbox, headers=TEXT_PREF,
              params={"$select": SUMMARY_FIELDS + ",bccRecipients,replyTo,body,hasAttachments,conversationId"})
    d = _summary(m)
    d.update(body=(m.get("body") or {}).get("content"), hasAttachments=m.get("hasAttachments"),
             replyTo=[_addr(r) for r in m.get("replyTo", [])], conversationId=m.get("conversationId"))
    return d


def _own_addresses() -> set[str]:
    me = graph("GET", "/me", params={"$select": "mail,proxyAddresses"})
    out = {p[5:].lower() for p in me.get("proxyAddresses") or [] if p.lower().startswith("smtp:")}
    if me.get("mail"):
        out.add(me["mail"].lower())
    return out


def _from_field(from_address: str | None, mailbox: str | None) -> dict | None:
    """From must be one of your own SMTP addresses or an allowed shared mailbox (Exchange still
    enforces Send As); anything else is refused before a draft is created."""
    if from_address:
        a = from_address.strip().lower()
        if a not in _env_list("M365_SHARED_MAILBOXES") and a not in _own_addresses():
            raise MailError(f"from_address {from_address} is neither one of your SMTP addresses (see "
                            "list_aliases) nor in M365_SHARED_MAILBOXES.")
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
    base = _base(mailbox)
    f = _from_field(from_address, mailbox)
    if f:
        msg["from"] = f
    if reply_to_message_id:
        d = graph("POST", f"{base}/messages/{_seg(reply_to_message_id)}/createReply", mailbox=mailbox, json={})
        d = graph("PATCH", f"{base}/messages/{_seg(d['id'])}", mailbox=mailbox, json=msg)
    else:
        d = graph("POST", f"{base}/messages", mailbox=mailbox, json=msg)
    return _readback(d)


@mcp.tool()
def reply_draft(message_id: str, body: str, reply_all: bool = False, from_address: str | None = None,
                mailbox: str | None = None, html: bool = False) -> dict:
    """Create (not send) a reply/reply-all draft to a message; your text is put above the quoted thread.
    Returns draft_id and readback; send with send_draft after approval."""
    base = _base(mailbox)
    f = _from_field(from_address, mailbox)
    action = "createReplyAll" if reply_all else "createReply"
    d = graph("POST", f"{base}/messages/{_seg(message_id)}/{action}", mailbox=mailbox, json={})
    did = d["id"]
    quoted = (d.get("body") or {}).get("content", "")
    if html or (d.get("body") or {}).get("contentType", "").lower() == "html":
        content = (body if html else "<p>" + body.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                   .replace("\n", "<br>") + "</p>")
        new_body = {"contentType": "HTML", "content": content + quoted}
    else:
        new_body = {"contentType": "Text", "content": body + "\n\n" + quoted}
    patch: dict[str, Any] = {"body": new_body}
    if f:
        patch["from"] = f
    d = graph("PATCH", f"{base}/messages/{_seg(did)}", mailbox=mailbox, json=patch)
    return _readback(d)


@mcp.tool()
def send_draft(draft_id: str, recipients: str | list[str], subject: str, mailbox: str | None = None) -> dict:
    """Send an existing draft. Only call after the user has approved the draft readback.

    recipients: every To/Cc/Bcc address on the draft; subject: the draft's subject. Both are shown
    in the approval prompt and are checked against the draft itself; on any mismatch nothing is sent.
    """
    base = _base(mailbox)
    path = f"{base}/messages/{_seg(draft_id)}"
    m = graph("GET", path, mailbox=mailbox, params={"$select": DRAFT_FIELDS})
    d = _readback(m)
    actual = {(r.get("emailAddress", {}).get("address") or "").lower()
              for k in ("toRecipients", "ccRecipients", "bccRecipients") for r in m.get(k, [])}
    expected = {r["emailAddress"]["address"].lower() for r in _recips(recipients)}
    if not d["isDraft"]:
        raise MailError(f"{draft_id} is not a draft; nothing was sent.")
    if actual != expected or (d["subject"] or "").strip() != subject.strip():
        raise MailError("Draft does not match the approved recipients/subject; nothing was sent. "
                        f"Draft has recipients {sorted(actual)} and subject {d['subject']!r}.")
    graph("POST", f"{path}/send", mailbox=mailbox, sending=True)
    return {"sent": True, **d}


@mcp.tool()
def move_message(id: str, destination_folder: str, mailbox: str | None = None) -> dict:
    """Move a message to a folder (well-known name like archive/deleteditems, or folder id)."""
    m = graph("POST", f"{_base(mailbox)}/messages/{_seg(id)}/move", mailbox=mailbox,
              json={"destinationId": destination_folder})
    return {"moved": True, "new_id": m.get("id")}


@mcp.tool()
def mark_read(id: str, is_read: bool = True, mailbox: str | None = None) -> dict:
    """Mark a message read (or unread with is_read=false)."""
    graph("PATCH", f"{_base(mailbox)}/messages/{_seg(id)}", mailbox=mailbox, json={"isRead": is_read})
    return {"id": id, "isRead": is_read}


if __name__ == "__main__":
    mcp.run("stdio")
