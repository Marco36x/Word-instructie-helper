"""Text Rewriter.

Een kleine lokale FastAPI-tool waarin je tekst kunt plakken en die door een
gratis OpenRouter-model herschreven wordt naar ondubbelzinnige, duidelijke,
technische tekst.

Start:
    pip install -r requirements.txt
    # zet OPENROUTER_API_KEY in een .env naast dit bestand (zie .env.example)
    python main.py            # opent http://127.0.0.1:8770

Optioneel een ander model kiezen via de `OPENROUTER_MODEL` env-variabele
(standaard ``meta-llama/llama-3.3-70b-instruct:free``).
"""

from __future__ import annotations

import argparse
import logging
import os
import sys
import webbrowser
from pathlib import Path
from typing import Any

from fastapi import FastAPI, HTTPException
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field

from openrouter_client import (
    DEFAULT_MODEL,
    FREE_MODEL_CHOICES,
    OpenRouterError,
    load_dotenv,
    rewrite_text,
)

logging.basicConfig(
    level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s"
)
logger = logging.getLogger("text_rewriter")

BASE_DIR = Path(__file__).parent.resolve()
STATIC_DIR = BASE_DIR / "static"

load_dotenv(BASE_DIR)


class RewriteRequest(BaseModel):
    text: str = Field(..., min_length=1)
    model: str | None = None
    instructions: str | None = None


class RewriteResponse(BaseModel):
    rewritten: str
    model: str
    fallback_used: bool = False
    requested_model: str | None = None


app = FastAPI(title="Text Rewriter")
app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")


@app.get("/", response_class=HTMLResponse)
def index() -> HTMLResponse:
    index_path = STATIC_DIR / "index.html"
    return HTMLResponse(index_path.read_text(encoding="utf-8"))


@app.get("/api/models")
def list_models() -> dict[str, Any]:
    selected = os.environ.get("OPENROUTER_MODEL", DEFAULT_MODEL)
    return {"models": FREE_MODEL_CHOICES, "default": selected}


@app.post("/api/rewrite", response_model=RewriteResponse)
def rewrite(req: RewriteRequest) -> RewriteResponse:
    api_key = os.environ.get("OPENROUTER_API_KEY", "").strip()
    if not api_key:
        raise HTTPException(
            status_code=400,
            detail=(
                "OPENROUTER_API_KEY ontbreekt. Maak een gratis key aan op "
                "https://openrouter.ai/keys en zet 'm in text_rewriter/.env"
            ),
        )

    requested_model = req.model or os.environ.get("OPENROUTER_MODEL") or DEFAULT_MODEL

    try:
        result = rewrite_text(
            req.text,
            api_key=api_key,
            model=requested_model,
            instructions=req.instructions,
        )
    except OpenRouterError as exc:
        raise HTTPException(
            status_code=exc.status_code or 502, detail=str(exc)
        ) from exc

    return RewriteResponse(
        rewritten=result.rewritten,
        model=result.model,
        fallback_used=result.fallback_used,
        requested_model=result.requested_model,
    )


@app.get("/api/health")
def health() -> dict[str, Any]:
    return {
        "status": "ok",
        "has_api_key": bool(os.environ.get("OPENROUTER_API_KEY", "").strip()),
        "default_model": os.environ.get("OPENROUTER_MODEL", DEFAULT_MODEL),
    }


def _parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Text Rewriter (OpenRouter)")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8770)
    parser.add_argument(
        "--no-browser",
        action="store_true",
        help="Open de browser niet automatisch na het starten.",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> None:
    args = _parse_args(argv if argv is not None else sys.argv[1:])
    url = f"http://{args.host}:{args.port}"
    logger.info("Text Rewriter beschikbaar op %s", url)
    if not args.no_browser:
        try:
            webbrowser.open(url)
        except Exception:  # pragma: no cover - browser open is best-effort
            logger.debug("Kon browser niet openen", exc_info=True)

    import uvicorn

    uvicorn.run(app, host=args.host, port=args.port, log_level="info")


if __name__ == "__main__":
    main()
