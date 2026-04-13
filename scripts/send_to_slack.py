import os
import sys
from slack_sdk import WebClient
from slack_sdk.errors import SlackApiError

token = os.environ.get("SLACK_BOT_TOKEN")
user_id = os.environ.get("SLACK_USER_ID")

if not token or not user_id:
    print("Error: SLACK_BOT_TOKEN or SLACK_USER_ID not set")
    sys.exit(1)

client = WebClient(token=token)

try:
    response = client.files_upload_v2(
        channel=user_id,
        file="daily_horoscope_hindi.xlsx",
        filename="daily_horoscope_hindi.xlsx",
        initial_comment="इस महीने का दैनिक राशिफल तैयार है। कृपया संलग्न फ़ाइल देखें। 🪐"
    )
    print("✅ File sent to Slack successfully")
except SlackApiError as e:
    print(f"❌ Slack error: {e.response['error']}")
    sys.exit(1)
