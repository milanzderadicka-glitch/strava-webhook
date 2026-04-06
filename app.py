import os
import requests
from flask import Flask, redirect, request

app = Flask(__name__)

@app.route("/")
def home():
    return """
    <h1>Strv Excel Projekt</h1>
    <p>Aplikace běží a vidí Strava údaje.</p>
    <p><a href="/login">Přihlásit se ke Stravě</a></p>
    """

@app.route("/login")
def login():
    client_id = os.getenv("STRAVA_CLIENT_ID")
    redirect_uri = request.host_url.rstrip("/") + "/exchange_token"

    auth_url = (
        "https://www.strava.com/oauth/authorize"
        f"?client_id={client_id}"
        "&response_type=code"
        f"&redirect_uri={redirect_uri}"
        "&approval_prompt=auto"
        "&scope=read,activity:read_all"
    )

    return redirect(auth_url)

@app.route("/exchange_token")
def exchange_token():
    code = request.args.get("code")

    if not code:
        return "Strava nevratila autorizacni kod."

    client_id = os.getenv("STRAVA_CLIENT_ID")
    client_secret = os.getenv("STRAVA_CLIENT_SECRET")

    response = requests.post(
        "https://www.strava.com/oauth/token",
        data={
            "client_id": client_id,
            "client_secret": client_secret,
            "code": code,
            "grant_type": "authorization_code",
        },
    )

    data = response.json()

    access_token = data.get("access_token")
    refresh_token = data.get("refresh_token")
    athlete = data.get("athlete", {})

    if access_token and refresh_token:
        return (
            "Pripojeni ke Strave probehlo uspesne.<br>"
            f"Athlete ID: {athlete.get('id')}<br>"
            "Access token i refresh token byly ziskany.<br><br>"
            f"Refresh token: {refresh_token}"
        )
    else:
        return f"Token se nepodarilo ziskat. Odpoved Stravy: {data}"

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
