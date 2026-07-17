#!/usr/bin/env python3
"""Servidor local simples para executar a PWA Twenty."""

from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
from pathlib import Path
import os
import webbrowser


ROOT = Path(__file__).resolve().parent
HOST = "127.0.0.1"
PORT = 8080


if __name__ == "__main__":
    os.chdir(ROOT)
    url = f"http://{HOST}:{PORT}"
    print(f"Twenty disponível em {url}")
    print("Mantém esta janela aberta. Usa Ctrl+C para parar.")
    webbrowser.open(url)
    ThreadingHTTPServer((HOST, PORT), SimpleHTTPRequestHandler).serve_forever()
