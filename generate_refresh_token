"""
generate_refresh_token.py
─────────────────────────
Run this ONCE interactively on any machine that has a browser.
It will print a refresh token that you then save as a GitHub Secret
named OD_REFRESH_TOKEN.

Usage:
    pip install msal
    python generate_refresh_token.py

Required environment variables (or edit the constants below):
    OD_TENANT_ID
    OD_CLIENT_ID
"""

import os
from msal import PublicClientApplication

# ── Edit these or set them as environment variables ──────────────────────────
TENANT_ID = os.environ.get("OD_TENANT_ID", "YOUR_OD_TENANT_ID_HERE")
CLIENT_ID = os.environ.get("OD_CLIENT_ID", "YOUR_OD_CLIENT_ID_HERE")
# ─────────────────────────────────────────────────────────────────────────────

SCOPES = [
    "https://graph.microsoft.com/Files.ReadWrite",
    "https://graph.microsoft.com/offline_access",
]

app = PublicClientApplication(
    CLIENT_ID,
    authority=f"https://login.microsoftonline.com/{TENANT_ID}",
)

# Try device-code flow (works headlessly too — just visit the URL on any device)
flow = app.initiate_device_flow(scopes=SCOPES)
if "user_code" not in flow:
    raise RuntimeError(f"Device flow initiation failed: {flow}")

print("\n" + "=" * 60)
print(flow["message"])          # "Go to https://microsoft.com/devicelogin and enter code XXXXX"
print("=" * 60 + "\n")

result = app.acquire_token_by_device_flow(flow)   # blocks until user completes login

if "refresh_token" not in result:
    raise RuntimeError(
        f"Token acquisition failed: {result.get('error_description', result)}"
    )

print("\n✅  Success!\n")
print("Add the following secret to GitHub → Settings → Secrets → Actions:\n")
print(f"  Name : OD_REFRESH_TOKEN")
print(f"  Value: {result['refresh_token']}\n")
print("The refresh token is valid for 90 days.")
print("Re-run this script before it expires to get a new one.\n")
