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

def get_athlete_data(access_token):
    response = requests.get(
        "https://www.strava.com/api/v3/athlete",
        headers={"Authorization": f"Bearer {access_token}"},
    )
    return response.json()

@app.route("/")
def home():
    token_data = get_access_token()
    access_token = token_data.get("access_token")

    if not access_token:
        return f"Pripojeni se nepodarilo. Odpoved Stravy: {token_data}"

    athlete_data = get_athlete_data(access_token)

    firstname = athlete_data.get("firstname", "")
    lastname = athlete_data.get("lastname", "")
    athlete_id = athlete_data.get("id", "")

    return (
        "<h1>Strv Excel Projekt</h1>"
        "<p>Automaticke pripojeni ke Strave funguje.</p>"
        f"<p>Sportovec: {firstname} {lastname}</p>"
        f"<p>Athlete ID: {athlete_id}</p>"
    )

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
