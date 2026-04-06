import os
from flask import Flask

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
    return "Login route funguje."

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
