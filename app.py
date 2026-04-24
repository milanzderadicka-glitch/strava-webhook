import os
import base64
import requests
from datetime import datetime
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

def get_shared_file_info(access_token):
    excel_url = os.getenv("EXCEL_SHARE_URL")

    encoded = base64.urlsafe_b64encode(excel_url.encode("utf-8")).decode("utf-8")
    encoded = encoded.rstrip("=").replace("/", "_").replace("+", "-")
    share_id = f"u!{encoded}"

    response = requests.get(
        f"https://graph.microsoft.com/v1.0/shares/{share_id}/driveItem",
        headers={"Authorization": f"Bearer {access_token}"},
    )

    return response.json()

def get_file_info_by_id(access_token):
    file_id = os.getenv("EXCEL_FILE_ID")

    response = requests.get(
        f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}",
        headers={"Authorization": f"Bearer {access_token}"},
    )

    return response.json()

def get_workbook_worksheets(access_token):
    file_id = os.getenv("EXCEL_FILE_ID")

    response = requests.get(
        f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/workbook/worksheets",
        headers={"Authorization": f"Bearer {access_token}"},
    )

    return response.json()

def get_parametry_headers(access_token):
    file_id = os.getenv("EXCEL_FILE_ID")

    response = requests.get(
        f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/workbook/worksheets('Parametry_tréninku')/range(address='A1:X1')",
        headers={"Authorization": f"Bearer {access_token}"},
    )

    return response.json()

def get_parametry_recent_rows(access_token):
    file_id = os.getenv("EXCEL_FILE_ID")

    response = requests.get(
        f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/workbook/worksheets('Parametry_tréninku')/range(address='A1730:X1760')",
        headers={"Authorization": f"Bearer {access_token}"},
    )

    return response.json()

def get_parametry_used_range(access_token):
    file_id = os.getenv("EXCEL_FILE_ID")

    response = requests.get(
        f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/workbook/worksheets('Parametry_tréninku')/usedRange(valuesOnly=true)",
        headers={"Authorization": f"Bearer {access_token}"},
    )

    return response.json()

def get_last_parametry_row(access_token):
    file_id = os.getenv("EXCEL_FILE_ID")

    response = requests.get(
        f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/workbook/worksheets('Parametry_tréninku')/range(address='A4474:X4474')",
        headers={"Authorization": f"Bearer {access_token}"},
    )

    return response.json()

def get_parametry_poradove_column(access_token):
    file_id = os.getenv("EXCEL_FILE_ID")

    response = requests.get(
        f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/workbook/worksheets('Parametry_tréninku')/range(address='A2:A5001')",
        headers={"Authorization": f"Bearer {access_token}"},
    )

    return response.json()

def get_parametry_strava_id_column(access_token):
    file_id = os.getenv("EXCEL_FILE_ID")

    response = requests.get(
        f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/workbook/worksheets('Parametry_tréninku')/range(address='X2:X5001')",
        headers={"Authorization": f"Bearer {access_token}"},
    )

    return response.json()

def find_last_filled_poradove_row(values, start_row=2):
    for i in range(len(values) - 1, -1, -1):
        cell = values[i][0] if values[i] else ""

        if cell not in ("", None):
            excel_row = start_row + i
            poradove_cislo = cell
            return excel_row, poradove_cislo

    return None, None

def find_last_filled_strava_id_row(values, start_row=2):
    for i in range(len(values) - 1, -1, -1):
        cell = values[i][0] if values[i] else ""

        if cell not in ("", None):
            excel_row = start_row + i
            strava_id = cell
            return excel_row, strava_id

    return None, None

def get_existing_strava_ids(values):
    ids = set()

    for row in values:
        if row and len(row) > 0:
            cell = row[0]
            if cell not in ("", None):
                ids.add(str(cell).strip())

    return ids

def format_strava_date(date_str):
    if not date_str:
        return ""
    dt = datetime.fromisoformat(date_str.replace("Z", "+00:00"))
    return dt.strftime("%d.%m.%Y")

def format_strava_time(date_str):
    if not date_str:
        return ""
    dt = datetime.fromisoformat(date_str.replace("Z", "+00:00"))
    return f"{dt.hour}:{dt.minute:02d}"

def format_hhmmss(seconds):
    if seconds in (None, ""):
        return ""
    seconds = int(round(float(seconds)))
    h = seconds // 3600
    m = (seconds % 3600) // 60
    s = seconds % 60
    return f"{h}:{m:02d}:{s:02d}"

def format_hmm(seconds):
    if seconds in (None, ""):
        return ""
    minutes = int(round(float(seconds) / 60))
    h = minutes // 60
    m = minutes % 60
    return f"{h}:{m:02d}"

def map_activity_from_strava(sport_type):
    mapping = {
        "Run": "běh",
        "Walk": "chůze",
        "Swim": "plavání",
        "Tennis": "tenis",
        "Squash": "squash",
        "Snowboard": "snowboard",
        "AlpineSki": "lyže - sjezd",
        "NordicSki": "běžky",
        "InlineSkate": "in-line",
        "Kayaking": "kajak",
        "Canoeing": "kanoe",
        "Triathlon": "triatlon",
        "WeightTraining": "posilovna",
        "Crossfit": "kruháč",
        "VirtualRide": "rotoped",
    }
    return mapping.get(sport_type, "")

def write_test_row(access_token):
    file_id = os.getenv("EXCEL_FILE_ID")

    # 1) zjistit dalsi radek a dalsi poradove cislo podle sloupce A
    col_data = get_parametry_poradove_column(access_token)
    col_values = col_data.get("values", [])
    excel_row, last_poradove_cislo = find_last_filled_poradove_row(col_values, start_row=2)

    if not excel_row or last_poradove_cislo in (None, ""):
        return {"error": "Nepodarilo se najit posledni vyplneny radek podle sloupce A."}

    next_excel_row = excel_row + 1
    next_poradove_cislo = int(last_poradove_cislo) + 1

    # 2) nacist posledni aktivitu ze Stravy
    strava_token_data = get_access_token()
    strava_access_token = strava_token_data.get("access_token")

    if not strava_access_token:
        return {"error": f"Nepodarilo se ziskat Strava access token. Odpoved: {strava_token_data}"}

    activities = get_recent_activities(strava_access_token)
    if not activities:
        return {"error": "Strava nevratila zadne aktivity."}

    latest_activity = activities[0]
    activity_id = latest_activity.get("id")

    detail = get_activity_detail(strava_access_token, activity_id)
    zones = get_activity_zones(strava_access_token, activity_id)

    # 3) zakladni mapovani
    start_date_local = detail.get("start_date_local", "")
    sport_type = detail.get("sport_type", "")
    mapped_activity = map_activity_from_strava(sport_type)

    datum = format_strava_date(start_date_local)
    cas = format_strava_time(start_date_local)
    delka = format_hhmmss(detail.get("moving_time", 0))

    tf_prumer = detail.get("average_heartrate")
    tf_max = detail.get("max_heartrate")
    kalorie = detail.get("calories")
    km_celkem = detail.get("distance")
    stoupani = detail.get("total_elevation_gain")

    tf_prumer = int(round(tf_prumer)) if tf_prumer not in (None, "") else ""
    tf_max = int(round(tf_max)) if tf_max not in (None, "") else ""
    kalorie = int(round(kalorie)) if kalorie not in (None, "") else ""
    km_celkem = round(float(km_celkem) / 1000, 2) if km_celkem not in (None, "") else ""
    stoupani = int(round(stoupani)) if stoupani not in (None, "") else ""

    # 4) zony
    hr_buckets = [0, 0, 0, 0, 0]
    for zone_group in zones:
        if zone_group.get("type") == "heartrate":
            distribution = zone_group.get("distribution_buckets", [])
            for i, bucket in enumerate(distribution[:5]):
                hr_buckets[i] = bucket.get("time", 0)

    pod_zonou_1 = format_hmm(hr_buckets[0])
    zona_1 = format_hmm(hr_buckets[1])
    zona_2 = format_hmm(hr_buckets[2])
    zona_3 = format_hmm(hr_buckets[3])
    nad_zonou_3 = format_hmm(hr_buckets[4])

    # 5) pripravit radek A:X
    values = [[
        next_poradove_cislo,     # A Pořadové číslo
        datum,                   # B Datum
        cas,                     # C Čas
        "",                      # D Kód aktivity
        mapped_activity,         # E Aktivita
        "",                      # F ID
        "",                      # G Trasa
        delka,                   # H Délka tréninku
        tf_prumer,               # I TF průměr
        tf_max,                  # J TF maximum
        kalorie,                 # K Spotřebovaná energie
        "",                      # L Tréninkový efekt
        zona_1,                  # M Zóna 1
        zona_2,                  # N Zóna 2
        zona_3,                  # O Zóna 3
        nad_zonou_3,             # P Nad zónou 3
        pod_zonou_1,             # Q Pod zónou 1
        "",                      # R Subjektivní hodnocení
        km_celkem,               # S Km celkem
        stoupani,                # T Stoupání
        "",                      # U Poznámka
        "",                      # V AI poznámka
        "",                      # W Vypočtený tréninkový efekt
        str(activity_id),        # X Strava ID
    ]]

    response = requests.patch(
        f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/workbook/worksheets('Parametry_tréninku')/range(address='A{next_excel_row}:X{next_excel_row}')",
        headers={
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        },
        json={"values": values},
    )

    if response.text:
        try:
            return response.json()
        except Exception:
            return {"status": response.status_code, "text": response.text}

    return {"status": "ok", "row": next_excel_row, "poradove_cislo": next_poradove_cislo, "strava_id": activity_id}
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

@app.route("/test-ms-refresh")
def test_ms_refresh():
    token_data = refresh_microsoft_token()

    access_token = token_data.get("access_token")
    new_refresh_token = token_data.get("refresh_token")

    if access_token and new_refresh_token:
        return (
            "<h1>Strv Excel Projekt</h1>"
            "<p>Microsoft refresh token funguje.</p>"
            "<p>Novy access token i novy refresh token byly ziskany.</p>"
        )
    else:
        return f"Obnoveni Microsoft tokenu selhalo. Odpoved: {token_data}"

@app.route("/test-drive-refresh")
def test_drive_refresh():
    token_data = refresh_microsoft_token()

    access_token = token_data.get("access_token")

    if not access_token:
        return f"Obnoveni Microsoft tokenu selhalo. Odpoved: {token_data}"

    drive_data = get_drive_info(access_token)

    drive_id = drive_data.get("id")
    drive_type = drive_data.get("driveType")
    owner = drive_data.get("owner", {})
    user = owner.get("user", {})
    display_name = user.get("displayName", "")

    if drive_id:
        return (
            "<h1>Strv Excel Projekt</h1>"
            "<p>Pristup k OneDrivu pres refresh token funguje.</p>"
            f"<p>Drive ID: {drive_id}</p>"
            f"<p>Drive type: {drive_type}</p>"
            f"<p>Vlastnik: {display_name}</p>"
        )
    else:
        return f"Pristup k OneDrivu pres refresh token selhal. Odpoved: {drive_data}"

@app.route("/test-excel-link")
def test_excel_link():
    excel_url = os.getenv("EXCEL_SHARE_URL")

    if excel_url:
        return (
            "<h1>Strv Excel Projekt</h1>"
            "<p>EXCEL_SHARE_URL je ulozeny.</p>"
            f"<p>{excel_url}</p>"
        )
    else:
        return "EXCEL_SHARE_URL neni ulozeny."
@app.route("/test-shared-file")
def test_shared_file():
    token_data = refresh_microsoft_token()

    access_token = token_data.get("access_token")

    if not access_token:
        return f"Obnoveni Microsoft tokenu selhalo. Odpoved: {token_data}"

    file_data = get_shared_file_info(access_token)

    file_name = file_data.get("name")
    file_id = file_data.get("id")
    web_url = file_data.get("webUrl")

    if file_id:
        return (
            "<h1>Strv Excel Projekt</h1>"
            "<p>Soubor byl nalezen pres sdileny odkaz.</p>"
            f"<p>Nazev: {file_name}</p>"
            f"<p>ID souboru: {file_id}</p>"
            f"<p>Web URL: {web_url}</p>"
        )
    else:
        return f"Nepodarilo se nacist metadata souboru. Odpoved: {file_data}"

@app.route("/test-file-id")
def test_file_id():
    token_data = refresh_microsoft_token()

    access_token = token_data.get("access_token")

    if not access_token:
        return f"Obnoveni Microsoft tokenu selhalo. Odpoved: {token_data}"

    file_data = get_file_info_by_id(access_token)

    file_name = file_data.get("name")
    file_id = file_data.get("id")
    web_url = file_data.get("webUrl")

    if file_id:
        return (
            "<h1>Strv Excel Projekt</h1>"
            "<p>Soubor byl nalezen pres EXCEL_FILE_ID.</p>"
            f"<p>Nazev: {file_name}</p>"
            f"<p>ID souboru: {file_id}</p>"
            f"<p>Web URL: {web_url}</p>"
        )
    else:
        return f"Nepodarilo se nacist metadata souboru pres ID. Odpoved: {file_data}"

@app.route("/test-worksheets")
def test_worksheets():
    token_data = refresh_microsoft_token()

    access_token = token_data.get("access_token")

    if not access_token:
        return f"Obnoveni Microsoft tokenu selhalo. Odpoved: {token_data}"

    sheets_data = get_workbook_worksheets(access_token)
    sheets = sheets_data.get("value", [])

    if sheets:
        html = "<h1>Strv Excel Projekt</h1><p>Seznam listu workbooku:</p><ul>"
        for sheet in sheets:
            html += f"<li>{sheet.get('name')}</li>"
        html += "</ul>"
        return html
    else:
        return f"Nepodarilo se nacist listy workbooku. Odpoved: {sheets_data}"

@app.route("/test-headers")
def test_headers():
    token_data = refresh_microsoft_token()

    access_token = token_data.get("access_token")

    if not access_token:
        return f"Obnoveni Microsoft tokenu selhalo. Odpoved: {token_data}"

    headers_data = get_parametry_headers(access_token)
    values = headers_data.get("values", [])

    if values and len(values) > 0:
        html = "<h1>Strv Excel Projekt</h1><p>Hlavičky listu Parametry_tréninku:</p><ul>"
        for header in values[0]:
            html += f"<li>{header}</li>"
        html += "</ul>"
        return html
    else:
        return f"Nepodarilo se nacist hlavicky. Odpoved: {headers_data}"

@app.route("/test-recent-rows")
def test_recent_rows():
    token_data = refresh_microsoft_token()

    access_token = token_data.get("access_token")

    if not access_token:
        return f"Obnoveni Microsoft tokenu selhalo. Odpoved: {token_data}"

    rows_data = get_parametry_recent_rows(access_token)
    values = rows_data.get("values", [])

    if values and len(values) > 0:
        html = "<h1>Strv Excel Projekt</h1><p>Posledni nacitene radky:</p><ul>"
        for row in values[-5:]:
            html += f"<li>{row}</li>"
        html += "</ul>"
        return html
    else:
        return f"Nepodarilo se nacist posledni radky. Odpoved: {rows_data}"

@app.route("/test-used-range")
def test_used_range():
    token_data = refresh_microsoft_token()

    access_token = token_data.get("access_token")

    if not access_token:
        return f"Obnoveni Microsoft tokenu selhalo. Odpoved: {token_data}"

    used_data = get_parametry_used_range(access_token)

    address = used_data.get("address")
    row_count = used_data.get("rowCount")
    column_count = used_data.get("columnCount")

    if address:
        return (
            "<h1>Strv Excel Projekt</h1>"
            "<p>Used range listu Parametry_tréninku:</p>"
            f"<p>Adresa: {address}</p>"
            f"<p>Pocet radku: {row_count}</p>"
            f"<p>Pocet sloupcu: {column_count}</p>"
        )
    else:
        return f"Nepodarilo se nacist used range. Odpoved: {used_data}"

@app.route("/test-last-row")
def test_last_row():
    token_data = refresh_microsoft_token()

    access_token = token_data.get("access_token")

    if not access_token:
        return f"Obnoveni Microsoft tokenu selhalo. Odpoved: {token_data}"

    row_data = get_last_parametry_row(access_token)
    values = row_data.get("values", [])

    if values and len(values) > 0:
        return (
            "<h1>Strv Excel Projekt</h1>"
            "<p>Posledni skutecny radek listu Parametry_tréninku:</p>"
            f"<p>{values[0]}</p>"
        )
    else:
        return f"Nepodarilo se nacist posledni radek. Odpoved: {row_data}"

@app.route("/test-next-row")
def test_next_row():
    last_excel_row = 4474
    last_poradove_cislo = 4501

    next_excel_row = last_excel_row + 1
    next_poradove_cislo = last_poradove_cislo + 1

    return (
        "<h1>Strv Excel Projekt</h1>"
        "<p>Vypocet dalsiho radku a poradi:</p>"
        f"<p>Dalsi radek v Excelu: {next_excel_row}</p>"
        f"<p>Dalsi poradove cislo: {next_poradove_cislo}</p>"
    )

@app.route("/test-poradove-column")
def test_poradove_column():
    token_data = refresh_microsoft_token()

    access_token = token_data.get("access_token")

    if not access_token:
        return f"Obnoveni Microsoft tokenu selhalo. Odpoved: {token_data}"

    col_data = get_parametry_poradove_column(access_token)
    values = col_data.get("values", [])

    if values and len(values) > 0:
        html = "<h1>Strv Excel Projekt</h1><p>Konec sloupce Pořadové číslo:</p><ul>"
        for row in values[-10:]:
            html += f"<li>{row}</li>"
        html += "</ul>"
        return html
    else:
        return f"Nepodarilo se nacist sloupec A. Odpoved: {col_data}"

@app.route("/test-find-last-poradove")
def test_find_last_poradove():
    token_data = refresh_microsoft_token()

    access_token = token_data.get("access_token")

    if not access_token:
        return f"Obnoveni Microsoft tokenu selhalo. Odpoved: {token_data}"

    col_data = get_parametry_poradove_column(access_token)
    values = col_data.get("values", [])

    excel_row, poradove_cislo = find_last_filled_poradove_row(values, start_row=2)

    if excel_row and poradove_cislo:
        return (
            "<h1>Strv Excel Projekt</h1>"
            "<p>Posledni skutecne vyplneny radek podle sloupce Pořadové číslo:</p>"
            f"<p>Radek v Excelu: {excel_row}</p>"
            f"<p>Posledni poradove cislo: {poradove_cislo}</p>"
        )
    else:
        return "Nepodarilo se najit posledni vyplneny radek podle sloupce A."

@app.route("/test-next-from-poradove")
def test_next_from_poradove():
    token_data = refresh_microsoft_token()

    access_token = token_data.get("access_token")

    if not access_token:
        return f"Obnoveni Microsoft tokenu selhalo. Odpoved: {token_data}"

    col_data = get_parametry_poradove_column(access_token)
    values = col_data.get("values", [])

    excel_row, poradove_cislo = find_last_filled_poradove_row(values, start_row=2)

    if excel_row and poradove_cislo:
        next_excel_row = excel_row + 1
        next_poradove_cislo = int(poradove_cislo) + 1

        return (
            "<h1>Strv Excel Projekt</h1>"
            "<p>Dalsi hodnoty vypoctene z Pořadového čísla:</p>"
            f"<p>Dalsi radek v Excelu: {next_excel_row}</p>"
            f"<p>Dalsi poradove cislo: {next_poradove_cislo}</p>"
        )
    else:
        return "Nepodarilo se spocitat dalsi radek a dalsi poradove cislo."

@app.route("/test-write-row")
def test_write_row():
    token_data = refresh_microsoft_token()

    access_token = token_data.get("access_token")

    if not access_token:
        return f"Obnoveni Microsoft tokenu selhalo. Odpoved: {token_data}"

    result = write_test_row(access_token)

    return (
        "<h1>Strv Excel Projekt</h1>"
        "<p>Test zapisove route probehl.</p>"
        f"<p>Odpoved: {result}</p>"
    )

@app.route("/test-strava-id-column")
def test_strava_id_column():
    token_data = refresh_microsoft_token()

    access_token = token_data.get("access_token")

    if not access_token:
        return f"Obnoveni Microsoft tokenu selhalo. Odpoved: {token_data}"

    col_data = get_parametry_strava_id_column(access_token)
    values = col_data.get("values", [])

    if values and len(values) > 0:
        html = "<h1>Strv Excel Projekt</h1><p>Konec sloupce Strava ID:</p><ul>"
        for row in values[-10:]:
            html += f"<li>{row}</li>"
        html += "</ul>"
        return html
    else:
        return f"Nepodarilo se nacist sloupec Strava ID. Odpoved: {col_data}"

@app.route("/test-find-last-strava-id")
def test_find_last_strava_id():
    token_data = refresh_microsoft_token()

    access_token = token_data.get("access_token")

    if not access_token:
        return f"Obnoveni Microsoft tokenu selhalo. Odpoved: {token_data}"

    col_data = get_parametry_strava_id_column(access_token)
    values = col_data.get("values", [])

    excel_row, strava_id = find_last_filled_strava_id_row(values, start_row=2)

    if excel_row and strava_id:
        return (
            "<h1>Strv Excel Projekt</h1>"
            "<p>Posledni vyplnene Strava ID:</p>"
            f"<p>Radek v Excelu: {excel_row}</p>"
            f"<p>Strava ID: {strava_id}</p>"
        )
    else:
        return "Nepodarilo se najit posledni vyplnene Strava ID."

@app.route("/test-existing-strava-ids")
def test_existing_strava_ids():
    token_data = refresh_microsoft_token()

    access_token = token_data.get("access_token")

    if not access_token:
        return f"Obnoveni Microsoft tokenu selhalo. Odpoved: {token_data}"

    col_data = get_parametry_strava_id_column(access_token)
    values = col_data.get("values", [])

    ids = get_existing_strava_ids(values)
    sample_ids = list(ids)[:10]

    html = "<h1>Strv Excel Projekt</h1>"
    html += f"<p>Pocet existujicich Strava ID: {len(ids)}</p>"
    html += "<p>Ukazka ID:</p><ul>"
    for item in sample_ids:
        html += f"<li>{item}</li>"
    html += "</ul>"

    return html

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
