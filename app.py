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

def get_recent_activities(access_token):
    response = requests.get(
        "https://www.strava.com/api/v3/athlete/activities?per_page=5",
        headers={"Authorization": f"Bearer {access_token}"},
    )
    return response.json()

def get_activity_detail(access_token, activity_id):
    response = requests.get(
        f"https://www.strava.com/api/v3/activities/{activity_id}",
        headers={"Authorization": f"Bearer {access_token}"},
    )
    return response.json()

@app.route("/")
def home():
    token_data = get_access_token()
    access_token = token_data.get("access_token")

    if not access_token:
        return f"Pripojeni se nepodarilo. Odpoved Stravy: {token_data}"

    activities = get_recent_activities(access_token)

    if not activities:
        return "Nebyla nalezena zadna aktivita."

    latest_activity = activities[0]
    activity_id = latest_activity.get("id")

    detail = get_activity_detail(access_token, activity_id)

    name = detail.get("name", "Bez nazvu")
    sport_type = detail.get("sport_type", "Neznamy typ")
    start_date = detail.get("start_date_local", "Neznamy cas")
    distance_km = round(detail.get("distance", 0) / 1000, 2)
    moving_time = detail.get("moving_time", 0)
    average_heartrate = detail.get("average_heartrate", "neni")
    max_heartrate = detail.get("max_heartrate", "neni")
    calories = detail.get("calories", "neni")
    elevation = detail.get("total_elevation_gain", 0)

    return (
        "<h1>Strv Excel Projekt</h1>"
        "<h2>Detail posledni aktivity</h2>"
        f"<p>Nazev: {name}</p>"
        f"<p>Typ aktivity: {sport_type}</p>"
        f"<p>Datum a cas: {start_date}</p>"
        f"<p>Vzdalenost: {distance_km} km</p>"
        f"<p>Moving time (sekundy): {moving_time}</p>"
        f"<p>Prumerna TF: {average_heartrate}</p>"
        f"<p>Maximalni TF: {max_heartrate}</p>"
        f"<p>Kalorie: {calories}</p>"
        f"<p>Stoupani: {elevation} m</p>"
        f"<p>Strava ID aktivity: {activity_id}</p>"
    )

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
