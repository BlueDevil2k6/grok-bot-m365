import json
import httpx
import pytest
import respx

import server
from server import GRAPH, MailError


@pytest.fixture(autouse=True)
def fake_token(monkeypatch):
    monkeypatch.setattr(server.auth, "token", lambda: "TEST")
    monkeypatch.setenv("M365_SHARED_MAILBOXES", "ops@contoso.com")


def mock_me():
    return respx.get(f"{GRAPH}/me").respond(json={"mail": "user@contoso.com", "userPrincipalName": "user@contoso.com",
        "proxyAddresses": ["SMTP:user@contoso.com", "smtp:t@contoso.com", "X500:/o=x", "smtp:sales@contoso.com"]})


def body(route, i=-1):
    return json.loads(route.calls[i].request.content)


@respx.mock
def test_list_aliases():
    mock_me()
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
    mock_me()
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
    mock_me()
    respx.post(f"{GRAPH}/me/messages/m1/createReplyAll").respond(
        json={"id": "r1", "body": {"contentType": "html", "content": "<div>old</div>"}})
    patch = respx.patch(f"{GRAPH}/me/messages/r1").mock(side_effect=lambda req: httpx.Response(
        200, json={"id": "r1", **json.loads(req.content)}))
    r = server.reply_draft("m1", "Thanks <b>\nok", reply_all=True, from_address="t@contoso.com")
    p = body(patch)
    assert p["body"]["content"] == "<p>Thanks &lt;b&gt;<br>ok</p><div>old</div>"
    assert p["from"]["emailAddress"]["address"] == "t@contoso.com"
    assert r["draft_id"] == "r1"


def draft(i, subject="Hi", to=("a@x.com",), bcc=()):
    return {"id": i, "subject": subject, "isDraft": True, "body": {"contentType": "Text", "content": "b"},
            "toRecipients": [{"emailAddress": {"address": a}} for a in to],
            "bccRecipients": [{"emailAddress": {"address": a}} for a in bcc]}


@respx.mock
def test_send_draft_ok_and_alias_rejected():
    respx.get(f"{GRAPH}/me/messages/d1").respond(json=draft("d1", to=("A@x.com", "b@x.com")))
    send = respx.post(f"{GRAPH}/me/messages/d1/send").respond(202)
    r = server.send_draft("d1", recipients="b@x.com; a@X.com", subject=" Hi ")
    assert r["sent"] is True and r["draft_id"] == "d1" and r["to"] == ["A@x.com", "b@x.com"] and send.called
    respx.get(f"{GRAPH}/me/messages/d2").respond(json=draft("d2"))
    respx.post(f"{GRAPH}/me/messages/d2/send").respond(403, json={"error": {
        "code": "ErrorSendAsDenied", "message": "The user account which was used to submit this request does not have the right to send mail on behalf of the specified sending account."}})
    with pytest.raises(MailError, match="SendFromAliasEnabled"):
        server.send_draft("d2", recipients=["a@x.com"], subject="Hi")


@respx.mock
@pytest.mark.parametrize("recipients,subject,d", [
    ("a@x.com", "Hi", draft("d", bcc=("evil@attacker.com",))),  # hidden extra recipient
    ("a@x.com", "Hi", draft("d", to=("evil@attacker.com",))),    # swapped recipient
    ("a@x.com", "Hi", draft("d", subject="Other")),              # different draft
    ("a@x.com", "Hi", {**draft("d"), "isDraft": False}),         # not a draft
])
def test_send_draft_refuses_mismatch(recipients, subject, d):
    respx.get(f"{GRAPH}/me/messages/d").respond(json=d)
    send = respx.post(f"{GRAPH}/me/messages/d/send").respond(202)
    with pytest.raises(MailError, match="nothing was sent"):
        server.send_draft("d", recipients=recipients, subject=subject)
    assert not send.called


@respx.mock
@pytest.mark.parametrize("call", [
    lambda: server.get_message("x/../../../users/ceo@contoso.com/messages/abc"),
    lambda: server.mark_read("../../users/ceo@contoso.com"),
    lambda: server.list_messages(folder="../../users/ceo@contoso.com/mailFolders/inbox"),
])
def test_ids_cannot_traverse_paths(call):
    route = respx.route().respond(404, json={"error": {"code": "ErrorItemNotFound", "message": "nf"}})
    with pytest.raises(MailError):
        call()
    url = str(route.calls[0].request.url)
    assert url.startswith(f"{GRAPH}/me/") and "/users/" not in url and "%2F" in url


@pytest.mark.parametrize("bad", ["..", "."])
def test_dot_ids_rejected(bad):
    with pytest.raises(MailError, match="Invalid id"):
        server.get_message(bad)


def test_mailbox_not_in_allowlist_refused():
    with pytest.raises(MailError, match="not in M365_SHARED_MAILBOXES"):
        server.list_messages(mailbox="ceo@contoso.com")
    with pytest.raises(MailError, match="not in M365_SHARED_MAILBOXES"):
        server.send_draft("d", recipients="a@x.com", subject="s", mailbox="ceo@contoso.com")


def test_graph_refuses_non_graph_url():
    with pytest.raises(MailError, match="non-Graph URL"):
        server.graph("GET", "https://evil.example/steal")


def test_plaintext_cache_requires_opt_in(monkeypatch):
    def no_keyring(_):
        raise RuntimeError("no keyring")
    monkeypatch.setattr(server, "build_encrypted_persistence", no_keyring)
    monkeypatch.delenv("M365_ALLOW_PLAINTEXT_CACHE", raising=False)
    with pytest.raises(MailError, match="M365_ALLOW_PLAINTEXT_CACHE"):
        server._build_cache()


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
def test_from_address_must_be_own_or_allowed():
    mock_me()
    route = respx.post(f"{GRAPH}/me/messages").mock(side_effect=lambda req: httpx.Response(
        201, json={"id": "d", **json.loads(req.content)}))
    assert server.create_draft(to="a@x.com", subject="s", body="b", from_address="Sales@Contoso.com")["draft_id"] == "d"
    assert server.create_draft(to="a@x.com", subject="s", body="b", from_address="ops@contoso.com")["draft_id"] == "d"
    with pytest.raises(MailError, match="neither one of your SMTP addresses"):
        server.create_draft(to="a@x.com", subject="s", body="b", from_address="ceo@contoso.com")
    with pytest.raises(MailError, match="neither one of your SMTP addresses"):
        server.reply_draft("m1", "hi", from_address="ceo@contoso.com")
    assert route.call_count == 2
