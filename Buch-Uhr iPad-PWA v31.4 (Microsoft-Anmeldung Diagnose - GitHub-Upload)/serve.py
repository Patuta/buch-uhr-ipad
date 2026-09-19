# coding: utf-8
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
from pathlib import Path
import os

ROOT = Path(__file__).resolve().parent
os.chdir(ROOT)

print("Buch-Uhr iPad-PWA läuft lokal unter http://localhost:8080")
print("Für die Microsoft-Appregistrierung muss diese Adresse als SPA-Umleitungs-URI eingetragen sein.")
ThreadingHTTPServer(("127.0.0.1", 8080), SimpleHTTPRequestHandler).serve_forever()
