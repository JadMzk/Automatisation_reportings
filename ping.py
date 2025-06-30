import requests
import time

while True:
    requests.get("https://autoreportings.streamlit.app/")
    time.sleep(1800)  # Ping toutes les 30 minutes
