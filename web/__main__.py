"""Pornire server: python -m web"""

import logging

from waitress import serve

from . import config
from .app import create_app

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

if __name__ == "__main__":
    # Un singur proces: firul de geocodare trăiește în interiorul lui
    serve(create_app(), host="0.0.0.0", port=config.PORT, threads=8)
