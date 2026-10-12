import json
import httpx
import pytest
import respx

import server
from server import GRAPH, MailError


@pytest.fixture(autouse=True)
def fake_token(monkeypatch):
    monkeypatch.setattr(server.auth, "token", lambda: "TEST")


def body(route, i=-1):
    return json.loads(route.calls[i].request.content)


@respx.mock
def test_list_aliases():
    respx.get(f"{GRAPH}/me").respond(json={"mail": "user@contoso.com", "userPrincipalName": "user@contoso.com",
        "proxyAddresses": ["SMTP:user@contoso.com", "smtp:t@contoso.com", "X500:/o=x", "smtp:sales@contoso.com"]})
    r = server.list_aliases()
    assert r["primary"] == "user@contoso.com"
    assert r["smtp_addresses"] == ["user@contoso.com", "t@contoso.com", "sales@contoso.com"]


def msg(i, to):
    return {"id": f"m{i}", "subject": f"s{i}", "from": {"emailAddress": {"address": "a@x.com"}},
            "toRecipients": [{"emailAddress": {"address": to}}], "ccRecipients": [], "isRead": False}


@respx.mock
def test_list_messages_to_alias_clientside_paging():
    nxt = f"{GRAPH}/me/mailFolders/inbox/messages?page=2"
    route = respx.get(url__startswith=f"{GRAPH}/me/mailFolders/inbox/messages").mock(side_effect=[
        httpx.Response(200, json={"value": [msg(1, "user@contoso.com"), msg(2, "SALES@contoso.com")],
                                   "@odata.nextLink": nxt}),
        httpx.Response(200, json={"value": [msg(3, "sales@contoso.com")]}),
    ])
    r = server.list_messages(to_address="sales@contoso.com", unread_only=True)
    assert [m["id"] for m in r] == ["m2", "m3"]
    q = route.calls[0].request.url.params
    assert q["$filter"] == "isRead eq false" and q["$orderby"] == "receivedDateTime desc"


@respx.mock
def test_list_messages_shared_search():
    route = respx.get(url__startswith=f"{GRAPH}/users/ops%40contoso.com/mailFolders/inbox/messages").respond(
        json={"value": [msg(1, "ops@contoso.com")]})
    r = server.list_messages(mailbox="ops@contoso.com", query="invoice", top=5)
    assert r[0]["to"] == ["ops@contoso.com"]
    assert route.calls[0].request.url.params["$search"] == '"invoice"'


@respx.mock
def test_get_message_text():
    route = respx.get(f"{GRAPH}/me/messages/m1").respond(json={**msg(1, "t@x"), "body": {"content": "hi"}})
    assert server.get_message("m1")["body"] == "hi"
    assert "text" in route.calls[0].request.headers["Prefer"]


@respx.mock
def test_create_draft_alias():
    route = respx.post(f"{GRAPH}/me/messages").mock(side_effect=lambda req: httpx.Response(
        201, json={"id": "d1", "isDraft": True, **json.loads(req.content)}))
    r = server.create_draft(to="a@x.com; b@x.com", subject="Hi", body="Yo", from_address="sales@contoso.com")
    sent = body(route)
    assert sent["from"]["emailAddress"]["address"] == "sales@contoso.com"
    assert len(sent["toRecipients"]) == 2 and sent["body"]["contentType"] == "Text"
    assert r["draft_id"] == "d1" and r["from"] == "sales@contoso.com" and r["body"] == "Yo"


@respx.mock
def test_create_draft_shared_mailbox_defaults_from():
    route = respx.post(f"{GRAPH}/users/ops%40contoso.com/messages").mock(side_effect=lambda req: httpx.Response(
        201, json={"id": "d2", **json.loads(req.content)}))
    r = server.create_draft(to=["a@x.com"], subject="S", body="B", mailbox="ops@contoso.com")
    assert body(route)["from"]["emailAddress"]["address"] == "ops@contoso.com"
    assert r["from"] == "ops@contoso.com"


@respx.mock
def test_reply_draft_all():
    respx.post(f"{GRAPH}/me/messages/m1/createReplyAll").respond(
        json={"id": "r1", "body": {"contentType": "html", "content": "<div>old</div>"}})
    patch = respx.patch(f"{GRAPH}/me/messages/r1").mock(side_effect=lambda req: httpx.Response(
        200, json={"id": "r1", **json.loads(req.content)}))
    r = server.reply_draft("m1", "Thanks <b>\nok", reply_all=True, from_address="t@contoso.com")
    p = body(patch)
    assert p["body"]["content"] == "<p>Thanks &lt;b&gt;<br>ok</p><div>old</div>"
    assert p["from"]["emailAddress"]["address"] == "t@contoso.com"
    assert r["draft_id"] == "r1"


@respx.mock
def test_send_draft_ok_and_alias_rejected():
    respx.post(f"{GRAPH}/me/messages/d1/send").respond(202)
    assert server.send_draft("d1") == {"sent": True, "draft_id": "d1"}
    respx.post(f"{GRAPH}/me/messages/d2/send").respond(403, json={"error": {
        "code": "ErrorSendAsDenied", "message": "The user account which was used to submit this request does not have the right to send mail on behalf of the specified sending account."}})
    with pytest.raises(MailError, match="SendFromAliasEnabled"):
        server.send_draft("d2")


@respx.mock
def test_shared_mailbox_403():
    respx.get(url__startswith=f"{GRAPH}/users/ops%40contoso.com/").respond(
        403, json={"error": {"code": "ErrorAccessDenied", "message": "Access is denied."}})
    with pytest.raises(MailError, match="Full Access"):
        server.list_messages(mailbox="ops@contoso.com")


@respx.mock
def test_move_and_mark_read():
    respx.post(f"{GRAPH}/me/messages/m1/move").respond(201, json={"id": "m1b"})
    assert server.move_message("m1", "archive")["new_id"] == "m1b"
    route = respx.patch(f"{GRAPH}/me/messages/m1").respond(200, json={"id": "m1"})
    server.mark_read("m1")
    assert body(route) == {"isRead": True}


def test_not_signed_in(monkeypatch):
    monkeypatch.undo()
    class FakeApp:
        def get_accounts(self): return []
    a = server.Auth()
    a._app = FakeApp()
    with pytest.raises(MailError, match="(?i)call auth_start"):
        a.token()
    assert a.status()["signed_in"] is False


def test_no_send_mail_tool():
    assert not hasattr(server, "send_mail")


@respx.mock
def test_alias_hint_env(monkeypatch):
    monkeypatch.setenv("M365_ALIASES", "user@contoso.com,user@example.com")
    respx.post(f"{GRAPH}/me/messages").mock(side_effect=lambda req: httpx.Response(
        201, json={"id": "d", **json.loads(req.content)}))
    assert "warnings" not in server.create_draft(to="a@x.com", subject="s", body="b", from_address="User@Example.com")
    r = server.create_draft(to="a@x.com", subject="s", body="b", from_address="bob@else.com")
    assert "not in M365_ALIASES" in r["warnings"][0] and r["from"] == "bob@else.com"
