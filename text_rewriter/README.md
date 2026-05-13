# Text Rewriter

Kleine lokale Python-tool om een paar regels tekst te herschrijven naar
**ondubbelzinnige, duidelijke, technische tekst** met behulp van een
**gratis model via [OpenRouter](https://openrouter.ai/)**.

Staat los van de Word Instructie Helper hoofd-app, maar woont in dezelfde repo.

## Wat doet het
- Webformulier waarin je een stuk tekst plakt.
- Verstuurt de tekst naar een gratis OpenRouter-model met een vaste
  Nederlandstalige systeemprompt die om ondubbelzinnige, technische tekst
  vraagt.
- Toont het resultaat ernaast, met een knop om naar het klembord te kopieren.
- Je kunt extra aanwijzingen meegeven (bv. "gebruik de wij-vorm" of
  "maak er een stapsgewijze instructie van").
- Modelkeuze via dropdown (lijst gratis modellen) of via `OPENROUTER_MODEL`
  in `.env`.

## Een (gratis) API key aanmaken
1. Ga naar <https://openrouter.ai/> en maak een account aan (gratis).
2. Open <https://openrouter.ai/keys> en klik **Create Key**.
3. Kopieer de key (begint met `sk-or-v1-...`).

> OpenRouter biedt gratis modellen aan zoals
> `meta-llama/llama-3.3-70b-instruct:free`, `deepseek/deepseek-chat-v3.1:free`,
> `google/gemini-2.0-flash-exp:free` en de router `openrouter/free`.
> Voor deze gratis modellen worden geen credits afgeschreven, maar er gelden
> wel rate limits.

## Installeren en starten

Vanuit de repo-root:

```bash
cd text_rewriter
python3 -m venv .venv
. .venv/bin/activate            # Windows: .venv\Scripts\activate
pip install -r requirements.txt

cp .env.example .env
# open .env en plak je OPENROUTER_API_KEY erin

python main.py                  # opent http://127.0.0.1:8770
```

Opties:
- `--host 0.0.0.0` om vanaf een ander apparaat te benaderen.
- `--port 9000` om een andere poort te kiezen.
- `--no-browser` om de browser niet automatisch te openen.

## Configuratie via `.env`

| Variabele            | Verplicht | Beschrijving                                    |
|----------------------|-----------|-------------------------------------------------|
| `OPENROUTER_API_KEY` | Ja        | Je gratis OpenRouter API key.                   |
| `OPENROUTER_MODEL`   | Nee       | Default-model. Anders: `meta-llama/llama-3.3-70b-instruct:free`. |

## API

De tool draait op FastAPI; handig als je het ook vanuit een script wilt
aanroepen.

`POST /api/rewrite`

```json
{
  "text": "ruwe tekst hier",
  "model": "meta-llama/llama-3.3-70b-instruct:free",
  "instructions": "optionele extra aanwijzingen"
}
```

Respons:

```json
{
  "rewritten": "herschreven tekst...",
  "model": "meta-llama/llama-3.3-70b-instruct:free"
}
```

`GET /api/health` geeft aan of er een API key geconfigureerd is.

## Veelvoorkomende problemen

- **`OPENROUTER_API_KEY ontbreekt`**: maak `text_rewriter/.env` aan met je
  key (zie `.env.example`).
- **HTTP 429 / rate-limited**: de gratis modellen hebben per dag een limiet.
  Wacht even of probeer een ander gratis model in de dropdown.
- **HTTP 401**: ongeldige of verlopen API key. Maak een nieuwe key aan op
  <https://openrouter.ai/keys>.
