import os
from flask import Flask

app = Flask(__name__)

@app.route("/")
def home():
    client_id = os.getenv("STRAVA_CLIENT_ID")
    client_secret = os.getenv("STRAVA_CLIENT_SECRET")

    if client_id and client_secret:
        return "Aplikace běží a vidi Strava udaje."
    else:
        return "Aplikace běží, ale nevidi Strava udaje."

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
