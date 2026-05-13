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
> `openai/gpt-oss-120b:free`, `openai/gpt-oss-20b:free`,
> `meta-llama/llama-3.3-70b-instruct:free`, `qwen/qwen3-next-80b-a3b-instruct:free`
> en `z-ai/glm-4.5-air:free`. Voor deze gratis modellen worden geen credits
> afgeschreven, maar er gelden wel rate limits. De tool valt automatisch terug
> op een ander gratis model uit de dropdown als het gekozen model op dat
> moment rate-limited is.

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

| Variabele            | Verplicht | Beschrijving                                            |
|----------------------|-----------|---------------------------------------------------------|
| `OPENROUTER_API_KEY` | Ja        | Je gratis OpenRouter API key.                           |
| `OPENROUTER_MODEL`   | Nee       | Default-model. Anders: `openai/gpt-oss-120b:free`.      |

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
  "model": "openai/gpt-oss-120b:free",
  "fallback_used": false,
  "requested_model": "openai/gpt-oss-120b:free"
}
```

Als het gevraagde model rate-limited is (HTTP 429), valt de server
automatisch terug op een ander gratis model uit de lijst. In dat geval is
`fallback_used: true` en bevat `requested_model` het oorspronkelijk
gevraagde model.

`GET /api/health` geeft aan of er een API key geconfigureerd is.

## Veelvoorkomende problemen

- **`OPENROUTER_API_KEY ontbreekt`**: maak `text_rewriter/.env` aan met je
  key (zie `.env.example`).
- **HTTP 429 / rate-limited**: de gratis modellen hebben per dag een limiet.
  De server probeert automatisch andere gratis modellen uit de dropdown
  voordat hij opgeeft. Als alles rate-limited is, wacht een paar minuten.
- **HTTP 401**: ongeldige of verlopen API key. Maak een nieuwe key aan op
  <https://openrouter.ai/keys>.
