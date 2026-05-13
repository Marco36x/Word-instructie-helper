"""Gedeelde OpenRouter-logica voor de Text Rewriter.

Wordt zowel door de FastAPI-server (``main.py``) als door de Tkinter-desktop
(``desktop.py``) gebruikt zodat beide UIs exact dezelfde modellen, prompt en
fallback-keten delen.
"""

from __future__ import annotations

import logging
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable

import httpx

logger = logging.getLogger("text_rewriter.openrouter")

DEFAULT_MODEL = "openai/gpt-oss-120b:free"
OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions"

# Lijst gratis OpenRouter-modellen die geschikt zijn voor tekst.
# Volgorde = fallback-volgorde: als het gekozen model 429/upstream-rate-limit
# geeft, probeert de client het volgende model in deze lijst.
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


class OpenRouterError(Exception):
    """Fout bij het aanroepen van OpenRouter."""

    def __init__(self, message: str, *, status_code: int | None = None) -> None:
        super().__init__(message)
        self.status_code = status_code


@dataclass
class RewriteResult:
    rewritten: str
    model: str
    fallback_used: bool
    requested_model: str


def load_dotenv(base_dir: Path) -> None:
    """Minimal .env loader: only sets vars not already in the environment."""

    env_path = base_dir / ".env"
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


def save_api_key_to_dotenv(base_dir: Path, api_key: str) -> Path:
    """Schrijf/overschrijf ``OPENROUTER_API_KEY`` in ``base_dir/.env``.

    Behoudt overige regels in het bestand. Maakt het bestand aan als het
    nog niet bestaat. Returnt het pad naar het .env-bestand.
    """

    env_path = base_dir / ".env"
    new_line = f"OPENROUTER_API_KEY={api_key.strip()}"
    if env_path.is_file():
        lines = env_path.read_text(encoding="utf-8").splitlines()
        replaced = False
        for i, raw in enumerate(lines):
            stripped = raw.lstrip()
            if stripped.startswith("OPENROUTER_API_KEY="):
                lines[i] = new_line
                replaced = True
                break
        if not replaced:
            lines.append(new_line)
        content = "\n".join(lines)
        if not content.endswith("\n"):
            content += "\n"
    else:
        content = new_line + "\n"
    env_path.write_text(content, encoding="utf-8")
    return env_path


def rewrite_text(
    text: str,
    *,
    api_key: str,
    model: str | None = None,
    instructions: str | None = None,
    timeout: float = 120.0,
    progress: Callable[[str, int, int], None] | None = None,
) -> RewriteResult:
    """Roep OpenRouter aan met automatische fallback over gratis modellen.

    :param text: De te herschrijven tekst.
    :param api_key: OpenRouter API key (``sk-or-v1-...``).
    :param model: Gewenst model-ID. ``None`` = ``DEFAULT_MODEL``.
    :param instructions: Optionele extra aanwijzingen voor de redacteur.
    :param timeout: HTTP timeout in seconden.
    :param progress: Optionele callback ``(model, attempt, total)`` voor
        UI-feedback; wordt aangeroepen vlak voor elke poging.
    """

    api_key = (api_key or "").strip()
    if not api_key:
        raise OpenRouterError(
            "OPENROUTER_API_KEY ontbreekt. Maak een gratis key aan op "
            "https://openrouter.ai/keys.",
            status_code=400,
        )

    text = (text or "").strip()
    if not text:
        raise OpenRouterError("Geen tekst om te herschrijven.", status_code=400)

    requested_model = (model or DEFAULT_MODEL).strip()

    user_content = text
    if instructions and instructions.strip():
        user_content = (
            f"Extra aanwijzingen voor het herschrijven:\n{instructions.strip()}\n\n"
            f"Tekst om te herschrijven:\n{text}"
        )
    else:
        user_content = f"Tekst om te herschrijven:\n{text}"

    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
        # OpenRouter raadt deze headers aan voor leaderboards; ze zijn optioneel.
        "HTTP-Referer": "https://github.com/Marco36x/Word-instructie-helper",
        "X-Title": "Word Instructie Helper - Text Rewriter",
    }

    fallback_ids = [m["id"] for m in FREE_MODEL_CHOICES if m["id"] != requested_model]
    model_chain: list[str] = [requested_model, *fallback_ids]

    last_error: tuple[int, str] | None = None
    for index, current_model in enumerate(model_chain):
        if progress is not None:
            try:
                progress(current_model, index + 1, len(model_chain))
            except Exception:  # noqa: BLE001 - callbacks mogen nooit de call breken
                logger.exception("progress-callback faalde")

        payload: dict[str, Any] = {
            "model": current_model,
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
            current_model,
            len(text),
        )

        try:
            with httpx.Client(timeout=timeout) as client:
                response = client.post(OPENROUTER_URL, headers=headers, json=payload)
        except httpx.HTTPError as exc:
            logger.exception("OpenRouter request failed for %s", current_model)
            last_error = (502, f"Kon OpenRouter niet bereiken: {exc}")
            continue

        if response.status_code == 200:
            data = response.json()
            try:
                rewritten = data["choices"][0]["message"]["content"].strip()
            except (KeyError, IndexError, TypeError):
                logger.error(
                    "Onverwacht OpenRouter-antwoord van %s: %s", current_model, data
                )
                last_error = (502, "OpenRouter gaf een onverwacht antwoord terug.")
                continue
            if not rewritten:
                last_error = (502, "OpenRouter gaf een leeg antwoord terug.")
                continue
            return RewriteResult(
                rewritten=rewritten,
                model=current_model,
                fallback_used=(current_model != requested_model),
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
            current_model,
            detail,
        )
        last_error = (response.status_code, f"OpenRouter-fout: {detail}")

        # Alleen doorvallen op rate-limit (429) of model-niet-gevonden (404).
        if response.status_code not in (404, 429) and upstream_code not in (404, 429):
            break

    status, detail = last_error or (502, "Onbekende OpenRouter-fout.")
    raise OpenRouterError(detail, status_code=status)
