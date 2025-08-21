from flask import Flask, render_template, request, redirect, session, url_for
import os
import json
import base64
from email.mime.text import MIMEText
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
import os
os.environ["OAUTHLIB_INSECURE_TRANSPORT"] = "1"

app = Flask(__name__)
app.secret_key = "your_super_secret_key"  # Change this!

# Google API settings
SCOPES = ["https://www.googleapis.com/auth/gmail.send"]
CLIENT_SECRETS_FILE = "credentials.json"
TOKEN_FILE = "token.json"


# ---------- Helpers ----------
def save_credentials(creds):
    with open(TOKEN_FILE, "w") as token_file:
        token_file.write(creds.to_json())


def load_credentials():
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, "r") as token_file:
            data = json.load(token_file)
            creds = Credentials.from_authorized_user_info(data, SCOPES)
            return creds
    return None


# ---------- Routes ----------
@app.route("/")
def index():
    creds = load_credentials()
    if not creds or not creds.valid:
        return redirect(url_for("authorize"))
    return render_template("index.html")


@app.route("/authorize")
def authorize():
    flow = Flow.from_client_secrets_file(
        CLIENT_SECRETS_FILE,
        scopes=SCOPES,
        redirect_uri=url_for("oauth2callback", _external=True),
    )
    auth_url, state = flow.authorization_url(
        access_type="offline", include_granted_scopes="true"
    )
    session["state"] = state  # store state safely
    return redirect(auth_url)


@app.route("/oauth2callback")
def oauth2callback():
    # Recreate the flow
    flow = Flow.from_client_secrets_file(
        CLIENT_SECRETS_FILE,
        scopes=SCOPES,
        redirect_uri=url_for("oauth2callback", _external=True),
    )

    # Fetch token from Google callback
    flow.fetch_token(authorization_response=request.url)

    creds = flow.credentials
    save_credentials(creds)  # persist token
    session["credentials"] = creds_to_dict(creds)

    return redirect(url_for("index"))


@app.route("/send", methods=["POST"])
def send_email():
    creds = load_credentials()
    if not creds or not creds.valid:
        return redirect(url_for("authorize"))

    service = build("gmail", "v1", credentials=creds)

    to = request.form["to"]
    subject = request.form["subject"]
    message_text = request.form["message"]

    message = MIMEText(message_text)
    message["to"] = to
    message["subject"] = subject

    raw = base64.urlsafe_b64encode(message.as_bytes()).decode()

    send_message = (
        service.users().messages().send(userId="me", body={"raw": raw}).execute()
    )

    return f" Message sent! ID: {send_message['id']}"


# ---------- Utils ----------
def creds_to_dict(creds):
    return {
        "token": creds.token,
        "refresh_token": creds.refresh_token,
        "token_uri": creds.token_uri,
        "client_id": creds.client_id,
        "client_secret": creds.client_secret,
        "scopes": creds.scopes,
    }


# ---------- Run ----------
if __name__ == "__main__":
    app.run(debug=True)
