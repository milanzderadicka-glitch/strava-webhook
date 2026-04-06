import os
import requests
from flask import Flask

app = Flask(__name__)

def get_access_token():
    client_id = os.getenv("STRAVA_CLIENT_ID")
    client_secret = os.getenv("STRAVA_CLIENT_SECRET")
    refresh_token = os.getenv("STRAVA_REFRESH_TOKEN")

    response = requests.post(
        "https://www.strava.com/oauth/token",
        data={
            "client_id": client_id,
            "client_secret": client_secret,
            "refresh_token": refresh_token,
            "grant_type": "refresh_token",
        },
    )

    return response.json()

@app.route("/")
def home():
    token_data = get_access_token()

    access_token = token_data.get("access_token")
    athlete = token_data.get("athlete", {})

    if access_token:
        firstname = athlete.get("firstname", "")
        lastname = athlete.get("lastname", "")
        athlete_id = athlete.get("id", "")

        return (
            "<h1>Strv Excel Projekt</h1>"
            "<p>Automaticke pripojeni ke Strave funguje.</p>"
            f"<p>Sportovec: {firstname} {lastname}</p>"
            f"<p>Athlete ID: {athlete_id}</p>"
        )
    else:
        return f"Pripojeni se nepodarilo. Odpoved Stravy: {token_data}"

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
