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

DEFAULT_MODEL = "openai/gpt-oss-120b:free"
OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions"

# Lijst gratis OpenRouter-modellen die geschikt zijn voor tekst.
# Volgorde = fallback-volgorde: als het gekozen model 429/upstream-rate-limit
# geeft, probeert de server het volgende model in deze lijst.
# Modellen die hier staan zijn op het moment van schrijven beschikbaar op
# https://openrouter.ai/models?free=true. Pas aan als OpenRouter ze offline
# haalt; de gebruiker kan met OPENROUTER_MODEL altijd eigen modellen forceren.
FREE_MODEL_CHOICES: list[dict[str, str]] = [
    {
        "id": "openai/gpt-oss-120b:free",
        "label": "OpenAI gpt-oss 120B (gratis)",
    },
    {
        "id": "openai/gpt-oss-20b:free",
        "label": "OpenAI gpt-oss 20B (gratis)",
    },
    {
        "id": "meta-llama/llama-3.3-70b-instruct:free",
        "label": "Meta Llama 3.3 70B Instruct (gratis)",
    },
    {
        "id": "qwen/qwen3-next-80b-a3b-instruct:free",
        "label": "Qwen3 Next 80B Instruct (gratis)",
    },
    {
        "id": "z-ai/glm-4.5-air:free",
        "label": "Z.ai GLM 4.5 Air (gratis)",
    },
    {
        "id": "nousresearch/hermes-3-llama-3.1-405b:free",
        "label": "Nous Hermes 3 Llama 3.1 405B (gratis)",
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

    requested_model = (
        req.model or os.environ.get("OPENROUTER_MODEL") or DEFAULT_MODEL
    ).strip()

    user_content = req.text.strip()
    if req.instructions and req.instructions.strip():
        user_content = (
            f"Extra aanwijzingen voor het herschrijven:\n{req.instructions.strip()}\n\n"
            f"Tekst om te herschrijven:\n{user_content}"
        )
    else:
        user_content = f"Tekst om te herschrijven:\n{user_content}"

    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
        # OpenRouter raadt deze headers aan voor leaderboards; ze zijn optioneel.
        "HTTP-Referer": "https://github.com/Marco36x/Word-instructie-helper",
        "X-Title": "Word Instructie Helper - Text Rewriter",
    }

    # Bouw de fallback-keten: eerst het gevraagde model, daarna de overige
    # gratis modellen uit FREE_MODEL_CHOICES (zonder duplicaten). Veel gratis
    # OpenRouter-modellen geven HTTP 429 'Provider returned error' wanneer de
    # upstream rate-limit even op is; door automatisch door te schakelen blijft
    # de tool bruikbaar zolang er ten minste een gratis model beschikbaar is.
    fallback_ids = [m["id"] for m in FREE_MODEL_CHOICES if m["id"] != requested_model]
    model_chain: list[str] = [requested_model, *fallback_ids]

    last_error: tuple[int, str] | None = None
    for index, model in enumerate(model_chain):
        payload: dict[str, Any] = {
            "model": model,
            "messages": [
                {"role": "system", "content": SYSTEM_PROMPT_NL},
                {"role": "user", "content": user_content},
            ],
            "temperature": 0.2,
        }

        logger.info(
            "Rewrite request (attempt %d/%d): model=%s, chars=%d",
            index + 1,
            len(model_chain),
            model,
            len(req.text),
        )

        try:
            with httpx.Client(timeout=120.0) as client:
                response = client.post(OPENROUTER_URL, headers=headers, json=payload)
        except httpx.HTTPError as exc:
            logger.exception("OpenRouter request failed for %s", model)
            last_error = (502, f"Kon OpenRouter niet bereiken: {exc}")
            continue

        if response.status_code == 200:
            data = response.json()
            try:
                rewritten = data["choices"][0]["message"]["content"].strip()
            except (KeyError, IndexError, TypeError):
                logger.error("Onverwacht OpenRouter-antwoord van %s: %s", model, data)
                last_error = (
                    502,
                    "OpenRouter gaf een onverwacht antwoord terug.",
                )
                continue
            if not rewritten:
                last_error = (502, "OpenRouter gaf een leeg antwoord terug.")
                continue
            return RewriteResponse(
                rewritten=rewritten,
                model=model,
                fallback_used=(model != requested_model),
                requested_model=requested_model,
            )

        detail = response.text
        upstream_code: int | None = None
        try:
            err_body = response.json()
            err_obj = err_body.get("error") if isinstance(err_body, dict) else None
            if isinstance(err_obj, dict):
                detail = err_obj.get("message") or detail
                upstream_code = err_obj.get("code")
        except ValueError:
            pass
        logger.warning(
            "OpenRouter error %s (upstream=%s) for %s: %s",
            response.status_code,
            upstream_code,
            model,
            detail,
        )
        last_error = (response.status_code, f"OpenRouter-fout: {detail}")

        # Alleen doorvallen op rate-limit (429) of model-niet-gevonden (404).
        # Andere fouten (401 auth, 400 bad request) hoeven we niet te retryen.
        if response.status_code not in (404, 429) and upstream_code not in (404, 429):
            break

    status, detail = last_error or (502, "Onbekende OpenRouter-fout.")
    raise HTTPException(status_code=status, detail=detail)


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
