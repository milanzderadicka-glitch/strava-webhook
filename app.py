import os
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

    if code:
        return f"Strava vratila autorizacni kod: {code}"
    else:
        return "Strava nevratila autorizacni kod."

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
