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

import httpx
from fastapi import FastAPI, HTTPException
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field

logging.basicConfig(
    level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s"
)
logger = logging.getLogger("text_rewriter")

BASE_DIR = Path(__file__).parent.resolve()
STATIC_DIR = BASE_DIR / "static"

DEFAULT_MODEL = "meta-llama/llama-3.3-70b-instruct:free"
OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions"

# Een korte lijst gratis OpenRouter-modellen die geschikt zijn voor tekst.
# De gebruiker kan via de UI of OPENROUTER_MODEL elk gewenst model kiezen;
# deze lijst dient alleen als dropdown-suggestie.
FREE_MODEL_CHOICES: list[dict[str, str]] = [
    {
        "id": "meta-llama/llama-3.3-70b-instruct:free",
        "label": "Llama 3.3 70B Instruct (gratis)",
    },
    {
        "id": "deepseek/deepseek-chat-v3.1:free",
        "label": "DeepSeek Chat v3.1 (gratis)",
    },
    {
        "id": "google/gemini-2.0-flash-exp:free",
        "label": "Gemini 2.0 Flash Experimental (gratis)",
    },
    {
        "id": "qwen/qwen-2.5-72b-instruct:free",
        "label": "Qwen 2.5 72B Instruct (gratis)",
    },
    {
        "id": "openrouter/free",
        "label": "OpenRouter auto (gratis router)",
    },
]


def _load_dotenv() -> None:
    """Minimal .env loader: only sets vars not already in the environment."""

    env_path = BASE_DIR / ".env"
    if not env_path.is_file():
        return
    for raw in env_path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value


_load_dotenv()


SYSTEM_PROMPT_NL = (
    "Je bent een Nederlandstalige redacteur die ruwe notities omzet in "
    "duidelijke, ondubbelzinnige, technische tekst. "
    "Regels:\n"
    "1. Behoud altijd de oorspronkelijke betekenis en feiten. Verzin niets bij.\n"
    "2. Schrijf in helder, zakelijk Nederlands. Korte zinnen, actieve vorm.\n"
    "3. Verwijder dubbelzinnigheid: vervang vage woorden door precieze termen.\n"
    "4. Gebruik consistente terminologie en juiste vaktermen.\n"
    "5. Behoud bestaande opsommingen of nummering als de structuur dat vraagt.\n"
    "6. Geef alleen de herschreven tekst terug, zonder uitleg, zonder "
    "voor- of nawoord, zonder markdown-codeblokken."
)


class RewriteRequest(BaseModel):
    text: str = Field(..., min_length=1)
    model: str | None = None
    instructions: str | None = None


class RewriteResponse(BaseModel):
    rewritten: str
    model: str


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

    model = (req.model or os.environ.get("OPENROUTER_MODEL") or DEFAULT_MODEL).strip()

    user_content = req.text.strip()
    if req.instructions and req.instructions.strip():
        user_content = (
            f"Extra aanwijzingen voor het herschrijven:\n{req.instructions.strip()}\n\n"
            f"Tekst om te herschrijven:\n{user_content}"
        )
    else:
        user_content = f"Tekst om te herschrijven:\n{user_content}"

    payload: dict[str, Any] = {
        "model": model,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT_NL},
            {"role": "user", "content": user_content},
        ],
        "temperature": 0.2,
    }

    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
        # OpenRouter raadt deze headers aan voor leaderboards; ze zijn optioneel.
        "HTTP-Referer": "https://github.com/Marco36x/Word-instructie-helper",
        "X-Title": "Word Instructie Helper - Text Rewriter",
    }

    logger.info("Rewrite request: model=%s, chars=%d", model, len(req.text))

    try:
        with httpx.Client(timeout=120.0) as client:
            response = client.post(OPENROUTER_URL, headers=headers, json=payload)
    except httpx.HTTPError as exc:
        logger.exception("OpenRouter request failed")
        raise HTTPException(
            status_code=502, detail=f"Kon OpenRouter niet bereiken: {exc}"
        ) from exc

    if response.status_code != 200:
        detail = response.text
        try:
            data = response.json()
            detail = data.get("error", {}).get("message") or detail
        except ValueError:
            pass
        logger.error("OpenRouter error %s: %s", response.status_code, detail)
        raise HTTPException(
            status_code=response.status_code,
            detail=f"OpenRouter-fout: {detail}",
        )

    data = response.json()
    try:
        rewritten = data["choices"][0]["message"]["content"].strip()
    except (KeyError, IndexError, TypeError) as exc:
        logger.error("Onverwacht OpenRouter-antwoord: %s", data)
        raise HTTPException(
            status_code=502,
            detail="OpenRouter gaf een onverwacht antwoord terug.",
        ) from exc

    return RewriteResponse(rewritten=rewritten, model=model)


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
