import os
import httpx
import msal
from datetime import datetime, timedelta, timezone
from dotenv import load_dotenv

load_dotenv()

GRAPH_ENDPOINT = "https://graph.microsoft.com/v1.0"
SCOPES = ["Mail.Read", "Mail.ReadWrite", "Mail.Send"]
RAILWAY_URL = "https://web-production-7cf443.up.railway.app"


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
    raise ValueError("Token expired. Re-run setup_auth.py first.")


def create_subscription():
    token = get_token()
    expiry = (datetime.now(timezone.utc) + timedelta(days=3)).strftime(
        "%Y-%m-%dT%H:%M:%S.0000000Z"
    )

    with httpx.Client() as client:
        response = client.post(
            f"{GRAPH_ENDPOINT}/subscriptions",
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json",
            },
            json={
                "changeType": "created",
                "notificationUrl": f"{RAILWAY_URL}/webhook",
                "resource": "me/mailFolders/inbox/messages",
                "expirationDateTime": expiry,
                "clientState": "nickelfox-email-automation",
            },
            timeout=30,
        )

        if not response.is_success:
            print(f"Failed: {response.status_code}")
            print(response.text)
            return

        sub = response.json()
        print(f"Subscription created successfully!")
        print(f"Subscription ID : {sub['id']}")
        print(f"Expires         : {sub['expirationDateTime']}")
        print(f"\nAdd this to Railway environment variables:")
        print(f"SUBSCRIPTION_ID={sub['id']}")


if __name__ == "__main__":
    create_subscription()
