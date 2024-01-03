import time
from slack_sdk import WebClient
from slack_sdk.errors import SlackApiError

from utils.aws_utils import get_aws_secret
SLACK_BOT_TOKEN = get_aws_secret("SLACK").get("TOKEN")

client = WebClient(token=SLACK_BOT_TOKEN)

def post_message(message, channel = 'U01J1EK7L1H'):
    try:
        response = client.chat_postMessage(channel = channel, text = message)
    except SlackApiError as e:
    # You will get a SlackApiError if "ok" is False
        assert e.response["ok"] is False

        error_str = e.response["error"]
        if(error_str in ["ratelimited", "rate_limited"]):
            time.sleep(1)
            response = client.chat_postMessage(channel = channel, text = message)
        else:
            print(f"Got an error: {e.response['error']}")