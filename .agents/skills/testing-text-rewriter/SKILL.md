---
name: testing-text-rewriter
description: Test the text_rewriter tool end-to-end (UI + real OpenRouter call). Use when verifying any change in text_rewriter/ or when an OpenRouter model id might have churned.
---

# Testing text_rewriter

The `text_rewriter/` folder is a standalone FastAPI tool that rewrites messy Dutch into clear technical Dutch via OpenRouter's free models. It is NOT wired into the main Word-instructie-helper app; test it on its own.

## Devin Secrets Needed
- `OPENROUTER_API_KEY` (org scope) — free key from https://openrouter.ai/keys. Starts with `sk-or-v1-`.

## Quick start (server)
```bash
cd /home/ubuntu/repos/Word-instructie-helper
. .venv/bin/activate  # repo has a shared venv at root
pip install -r text_rewriter/requirements.txt  # only first time
OPENROUTER_API_KEY="$OPENROUTER_API_KEY" python text_rewriter/main.py \
  --host 127.0.0.1 --port 8770 --no-browser
```
Server logs to stdout; UI at http://127.0.0.1:8770. Run in background and tail logs to spot 401/429/404 from OpenRouter.

## Pre-test sanity check (ALWAYS do this first)
OpenRouter's free-model roster churns and free models are often rate-limited. Before any UI test, verify the default model and at least one fallback respond:
```bash
curl -s https://openrouter.ai/api/v1/chat/completions \
  -H "Authorization: Bearer $OPENROUTER_API_KEY" \
  -H "Content-Type: application/json" \
  -d '{"model":"openai/gpt-oss-120b:free","messages":[{"role":"user","content":"ping"}],"max_tokens":4}' | jq '.error // .choices[0].message.content'
```
- `404 No endpoints found` → model id is dead; replace it in `FREE_MODEL_CHOICES` in `text_rewriter/main.py`.
- `429` with `Provider returned error` → rate-limited right now; either wait, switch default, or rely on the built-in fallback chain (the tool retries the next id in `FREE_MODEL_CHOICES` on 404/429 and stops immediately on 401).
- `401` → key invalid / not set.

To discover currently-listed free models: `curl -s https://openrouter.ai/api/v1/models | jq '.data[] | select(.id|endswith(":free")) | .id'`. Pick a few, smoke-test each with the curl above before adding to the dropdown.

## Core test plan (3 tests)
The canonical plan lives at `text_rewriter/TEST_PLAN.md`. Highlights:
1. **UI laadt + dropdown** — 6 free models, default `OpenAI gpt-oss 120B (gratis)`, `Kopieer resultaat` disabled.
2. **Echte herschrijving** — paste the rommelige zin from the plan, click `Herschrijf tekst`, assert:
   - Status flips `Bezig met herschrijven...` → `Klaar.`
   - Output ≥ 80 chars; contains NONE of `ehm`, `joh`, `ofzo`, `nou ja`, `geloof ik`; contains ≥ 1 of `knop`, `klik`, `bestand`, `openen`.
   - Meta line shows the requested model id, OR `(fallback van …)` if the chain kicked in.
3. **Kopieer resultaat** — click the button, clear input, `Ctrl+V`; pasted content must equal the output textarea (not the original input).

## Gotchas / future-proofing
- The default model in `FREE_MODEL_CHOICES[0]` may be 429 on any given day. The fallback chain handles this silently; the UI then shows `Klaar (gekozen model was rate-limited; fallback gebruikt).` and `(fallback van …)` in the meta. Treat that as a pass for Test 2.
- If multiple models in a row are 429/404, surface that and re-run the pre-test curl to refresh the model list — do not modify tests to make them pass.
- The repo has no CI configured, so PR checks will always be empty. Don't wait on CI.
- Always record the UI flow when testing this tool (it's a visual app) and annotate test_start/assertion via `annotate_recording`.
- Don't paste API keys into the chat — use the `secrets` tool with `should_save=true, save_scope="org"` so future sessions inherit `OPENROUTER_API_KEY`.
