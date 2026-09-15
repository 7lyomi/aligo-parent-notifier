"""Application factory: configure independent app instances and wire their modules."""

import os
from pathlib import Path

from dotenv import load_dotenv
from flask import Flask

from .service import NotifierService
from .settings import Settings
from .storage import RuntimePaths
from .web import web

PROJECT_DIR = Path(__file__).resolve().parent.parent


def create_app(settings: Settings | None = None, runtime_dir: Path | None = None, client=None):
    if settings is None:
        load_dotenv(PROJECT_DIR / ".env")
        settings = Settings.from_env()
    folder = "demo" if settings.is_demo else "live"
    runtime_dir = (
        runtime_dir
        if runtime_dir is not None
        else Path(os.getenv("APP_DATA_DIR", str(PROJECT_DIR / "runtime" / folder)))
    )
    paths = RuntimePaths(Path(runtime_dir).resolve())
    service = NotifierService(
        settings, paths, PROJECT_DIR / "samples" / "sample_students.xlsx", client
    )
    app = Flask(
        __name__,
        template_folder=str(PROJECT_DIR / "templates"),
        static_folder=str(PROJECT_DIR / "static"),
        static_url_path="/static",
    )
    app.secret_key = os.getenv("FLASK_SECRET_KEY") or paths.session_secret()
    app.config.update(
        MAX_CONTENT_LENGTH=8 * 1024 * 1024,
        MAX_FORM_MEMORY_SIZE=2 * 1024 * 1024,
        SESSION_COOKIE_HTTPONLY=True,
        SESSION_COOKIE_SAMESITE="Strict",
    )
    app.extensions["notifier"] = service
    app.register_blueprint(web)
    return app
