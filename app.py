import os
import requests
from flask import Flask, redirect, request

app = Flask(__name__)

def get_microsoft_token():
    tenant_id = os.getenv("MS_TENANT_ID")
    client_id = os.getenv("MS_CLIENT_ID")
    client_secret = os.getenv("MS_CLIENT_SECRET")

    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"

    response = requests.post(
        token_url,
        data={
            "client_id": client_id,
            "client_secret": client_secret,
            "scope": "https://graph.microsoft.com/.default",
            "grant_type": "client_credentials",
        },
    )

    return response.json()

def get_microsoft_auth_url():
    client_id = os.getenv("MS_CLIENT_ID")
    redirect_uri = "https://strava-webhook-l8mx.onrender.com/ms-callback"

    return (
        "https://login.microsoftonline.com/common/oauth2/v2.0/authorize"
        f"?client_id={client_id}"
        "&response_type=code"
        f"&redirect_uri={redirect_uri}"
        "&response_mode=query"
        "&scope=offline_access User.Read Files.ReadWrite"
    )
def exchange_microsoft_code(code):
    tenant_id = os.getenv("MS_TENANT_ID")
    client_id = os.getenv("MS_CLIENT_ID")
    client_secret = os.getenv("MS_CLIENT_SECRET")
    redirect_uri = "https://strava-webhook-l8mx.onrender.com/ms-callback"

    token_url = "https://login.microsoftonline.com/common/oauth2/v2.0/token"

    response = requests.post(
        token_url,
        data={
            "client_id": client_id,
            "client_secret": client_secret,
            "code": code,
            "redirect_uri": redirect_uri,
            "grant_type": "authorization_code",
        },
    )

    return response.json()

def refresh_microsoft_token():
    client_id = os.getenv("MS_CLIENT_ID")
    client_secret = os.getenv("MS_CLIENT_SECRET")
    refresh_token = os.getenv("MS_REFRESH_TOKEN")
    redirect_uri = "https://strava-webhook-l8mx.onrender.com/ms-callback"

    token_url = "https://login.microsoftonline.com/common/oauth2/v2.0/token"

    response = requests.post(
        token_url,
        data={
            "client_id": client_id,
            "client_secret": client_secret,
            "refresh_token": refresh_token,
            "redirect_uri": redirect_uri,
            "grant_type": "refresh_token",
            "scope": "offline_access User.Read Files.ReadWrite",
        },
    )

    return response.json()

def get_drive_info(access_token):
    response = requests.get(
        "https://graph.microsoft.com/v1.0/me/drive",
        headers={"Authorization": f"Bearer {access_token}"},
    )
    return response.json()

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

def get_activity_zones(access_token, activity_id):
    response = requests.get(
        f"https://www.strava.com/api/v3/activities/{activity_id}/zones",
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
    zones = get_activity_zones(access_token, activity_id)

    name = detail.get("name", "Bez nazvu")
    sport_type = detail.get("sport_type", "Neznamy typ")
    start_date = detail.get("start_date_local", "Neznamy cas")
    distance_km = round(detail.get("distance", 0) / 1000, 2)
    moving_time = detail.get("moving_time", 0)
    average_heartrate = detail.get("average_heartrate", "neni")
    max_heartrate = detail.get("max_heartrate", "neni")
    calories = detail.get("calories", "neni")
    elevation = detail.get("total_elevation_gain", 0)

    zone_html = "<h3>Tepove zony</h3><ul>"

    for zone_group in zones:
        if zone_group.get("type") == "heartrate":
            distribution = zone_group.get("distribution_buckets", [])
            for i, bucket in enumerate(distribution, start=1):
                seconds = bucket.get("time", 0)
                zone_html += f"<li>Strava zona {i}: {seconds} s</li>"

    zone_html += "</ul>"

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
        f"{zone_html}"
    )
@app.route("/test-ms")
def test_ms():
    token_data = get_microsoft_token()
    access_token = token_data.get("access_token")

    if access_token:
        return (
            "<h1>Strv Excel Projekt</h1>"
            "<p>Microsoft prihlaseni funguje.</p>"
            "<p>Access token byl uspesne ziskan.</p>"
        )
    else:
        return f"Microsoft prihlaseni selhalo. Odpoved: {token_data}"

@app.route("/test-drive")
def test_drive():
    token_data = get_microsoft_token()
    access_token = token_data.get("access_token")

    if not access_token:
        return f"Microsoft prihlaseni selhalo. Odpoved: {token_data}"

    drive_data = get_drive_info(access_token)

    drive_id = drive_data.get("id")
    drive_type = drive_data.get("driveType")
    owner = drive_data.get("owner", {})
    user = owner.get("user", {})
    display_name = user.get("displayName", "")

    if drive_id:
        return (
            "<h1>Strv Excel Projekt</h1>"
            "<p>Pristup k OneDrivu funguje.</p>"
            f"<p>Drive ID: {drive_id}</p>"
            f"<p>Drive type: {drive_type}</p>"
            f"<p>Vlastnik: {display_name}</p>"
        )
    else:
        return f"Pristup k OneDrivu selhal. Odpoved: {drive_data}"

@app.route("/login-ms")
def login_ms():
    return redirect(get_microsoft_auth_url())


@app.route("/ms-callback")
def ms_callback():
    code = request.args.get("code")

    if not code:
        return "Microsoft nevratil autorizacni kod."

    token_data = exchange_microsoft_code(code)

    access_token = token_data.get("access_token")
    refresh_token = token_data.get("refresh_token")

    if access_token and refresh_token:
        return (
        "Microsoft prihlaseni probehlo uspesne.<br>"
        "Access token i refresh token byly ziskany."
    )
    else:
        return f"Ziskani Microsoft tokenu selhalo. Odpoved: {token_data}"
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
