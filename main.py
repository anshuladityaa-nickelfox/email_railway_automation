import os
import re
import asyncio
import httpx
import msal
import uvicorn
import threading
import time as time_module
from datetime import datetime, timedelta, timezone, time
from zoneinfo import ZoneInfo
from starlette.applications import Starlette
from starlette.routing import Route
from starlette.requests import Request
from starlette.responses import JSONResponse, PlainTextResponse
from dotenv import load_dotenv

load_dotenv()

GRAPH_ENDPOINT = "https://graph.microsoft.com/v1.0"
SCOPES = ["Mail.Read", "Mail.ReadWrite", "Mail.Send"]
SUBJECT_PREFIX = "new project lead from"
IST = ZoneInfo("Asia/Kolkata")

REPLY_TEMPLATE_WORKING_HOURS = """Hi,

Thank you for reaching out and for sharing your project brief. This looks like a great fit for us, and we would love to discuss it further.

If you are available right now, we can jump on a quick 10-minute call to align on a few points and understand your requirements better. We are flexible, we can do a proper meeting link, a quick phone call, or even a WhatsApp audio call, whichever is more convenient for you.

If now is not a good time, and you can be available in the next couple of hours, you can also book a suitable slot here: http://nickelfox.com/booking

Looking forward to speaking with you."""

REPLY_TEMPLATE_OFF_HOURS = """Hi,

Thank you for reaching out and for sharing your project brief. This looks like a great fit for us, and we would love to discuss it further.

We are not available at this exact moment, but we will be free in the next couple of hours and would really like to have a call then. We are flexible and can do a meeting link, a quick phone call, or even a WhatsApp audio call.

If you would like to lock in a specific time for later today, you can also book a suitable slot here: http://nickelfox.com/booking

Looking forward to speaking with you soon."""


def get_token() -> str:
    cache = msal.SerializableTokenCache()
    cache.deserialize(os.environ["MSAL_TOKEN_CACHE"])
    app = msal.PublicClientApplication(
        client_id=os.environ["CLIENT_ID"],
        authority=f"https://login.microsoftonline.com/{os.environ['TENANT_ID']}",
        token_cache=cache,
    )
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            return result["access_token"]
    raise ValueError("Token expired. Re-run setup_auth.py and update MSAL_TOKEN_CACHE.")


def is_working_hours() -> bool:
    now = datetime.now(IST).time()
    # Working hours: 10:00 AM to 3:00 AM (next day)
    # Off hours: 3:00 AM to 10:00 AM
    return not (time(3, 0) <= now < time(10, 0))


def strip_html(text: str) -> str:
    text = re.sub(r'<[^>]+>', ' ', text)
    text = re.sub(r'&nbsp;', ' ', text)
    text = re.sub(r'&amp;', '&', text)
    text = re.sub(r'&lt;', '<', text)
    text = re.sub(r'&gt;', '>', text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip()


async def fetch_email(message_id: str) -> dict:
    token = get_token()
    async with httpx.AsyncClient() as client:
        resp = await client.get(
            f"{GRAPH_ENDPOINT}/me/messages/{message_id}",
            headers={"Authorization": f"Bearer {token}"},
            params={"$select": "id,subject,body,from,conversationId"},
            timeout=30,
        )
        if not resp.is_success:
            raise ValueError(f"Failed to fetch email: {resp.status_code} {resp.text}")
        return resp.json()


async def already_replied(conversation_id: str) -> bool:
    token = get_token()
    async with httpx.AsyncClient() as client:
        resp = await client.get(
            f"{GRAPH_ENDPOINT}/me/mailFolders/sentitems/messages",
            headers={"Authorization": f"Bearer {token}"},
            params={
                "$filter": f"conversationId eq '{conversation_id}'",
                "$select": "id",
                "$top": "1",
            },
            timeout=30,
        )
        if resp.is_success:
            return len(resp.json().get("value", [])) > 0
    return False


async def send_reply(message_id: str, body: str):
    token = get_token()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    async with httpx.AsyncClient() as client:
        create_resp = await client.post(
            f"{GRAPH_ENDPOINT}/me/messages/{message_id}/createReply",
            headers=headers,
            json={},
            timeout=30,
        )
        if not create_resp.is_success:
            raise ValueError(f"createReply failed: {create_resp.status_code} {create_resp.text}")

        draft_id = create_resp.json()["id"]

        update_resp = await client.patch(
            f"{GRAPH_ENDPOINT}/me/messages/{draft_id}",
            headers=headers,
            json={"body": {"contentType": "Text", "content": body}},
            timeout=30,
        )
        if not update_resp.is_success:
            raise ValueError(f"PATCH body failed: {update_resp.status_code} {update_resp.text}")

        send_resp = await client.post(
            f"{GRAPH_ENDPOINT}/me/messages/{draft_id}/send",
            headers=headers,
            timeout=30,
        )
        if not send_resp.is_success:
            raise ValueError(f"Send failed: {send_resp.status_code} {send_resp.text}")


def is_lead_email(subject: str, body: str) -> bool:
    return subject.lower().startswith(SUBJECT_PREFIX)


async def process_email(message_id: str):
    try:
        email = await fetch_email(message_id)

        subject = email.get("subject", "")
        raw_body = email.get("body", {}).get("content", "")
        body_content = strip_html(raw_body)
        sender = email.get("from", {}).get("emailAddress", {})
        sender_name = sender.get("name", "")
        sender_email = sender.get("address", "")
        conversation_id = email.get("conversationId", "")
        email_id = email.get("id", message_id)

        if not is_lead_email(subject, body_content):
            print(f"Skipped (subject mismatch): {subject}")
            return

        if await already_replied(conversation_id):
            print(f"Already replied, skipping: {subject}")
            return

        working = is_working_hours()
        template = REPLY_TEMPLATE_WORKING_HOURS if working else REPLY_TEMPLATE_OFF_HOURS
        await send_reply(email_id, template)
        hours_label = "working hours" if working else "off hours"
        print(f"Reply sent ({hours_label}) -- To: {sender_name} <{sender_email}> | Subject: {subject}")

    except Exception as e:
        print(f"Error processing email {message_id}: {e}")


async def webhook(request: Request):
    validation_token = request.query_params.get("validationToken")
    if validation_token:
        return PlainTextResponse(content=validation_token, media_type="text/plain")

    try:
        data = await request.json()
        for notification in data.get("value", []):
            if notification.get("clientState") != "nickelfox-email-automation":
                continue
            message_id = notification.get("resourceData", {}).get("id")
            if message_id:
                asyncio.create_task(process_email(message_id))
    except Exception as e:
        print(f"Webhook parse error: {e}")

    return JSONResponse({"status": "ok"})


async def health(request: Request):
    return JSONResponse({"status": "ok"})


app = Starlette(routes=[
    Route("/webhook", webhook, methods=["GET", "POST"]),
    Route("/health", health, methods=["GET"]),
])

def _renewal_loop():
    while True:
        time_module.sleep(47 * 3600)
        sub_id = os.environ.get("SUBSCRIPTION_ID")
        if not sub_id:
            print("TOKEN WARNING: SUBSCRIPTION_ID not set — cannot renew subscription.")
            continue
        try:
            token = get_token()
            new_expiry = (datetime.now(timezone.utc) + timedelta(days=3)).strftime(
                "%Y-%m-%dT%H:%M:%S.0000000Z"
            )
            with httpx.Client() as client:
                resp = client.patch(
                    f"{GRAPH_ENDPOINT}/subscriptions/{sub_id}",
                    headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
                    json={"expirationDateTime": new_expiry},
                    timeout=30,
                )
            if resp.is_success:
                print(f"Subscription renewed until {new_expiry}")
            else:
                print(f"Renewal failed: {resp.status_code} {resp.text}")
        except ValueError as e:
            print(f"TOKEN EXPIRED: {e} -- Re-run setup_auth.py and update MSAL_TOKEN_CACHE in Railway env vars.")
        except Exception as e:
            print(f"Renewal error: {e}")


threading.Thread(target=_renewal_loop, daemon=True).start()


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)
