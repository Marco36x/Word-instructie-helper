---
name: testing-word-instructie-helper
description: End-to-end testen van de Word-instructie-helper. Gebruik dit bij UI- of clipboard-gerelateerde wijzigingen aan de FastAPI-app of de previews/clipboard-endpoints.
---

# Testing Word-instructie-helper

Lokale FastAPI-app die `.docx`-previews toont en met een klik de inhoud van een Word-document op het klembord zet (om in Word te plakken met `Ctrl+V`). Er is ook een alternatieve `INCLUDETEXT`-knop.

## Hoe te starten
```bash
python3 -m venv .venv && . .venv/bin/activate
pip install -r requirements.txt
python main.py --host 127.0.0.1 --port 8765 --no-browser
```
Daarna opent de app op http://127.0.0.1:8765.

De blueprint installeert al `xclip` + `libreoffice-core` + `libreoffice-writer` en zet de venv klaar.

## Twee klembord-paden
- **Windows (echte gebruik)**: `main.py:_copy_docx_to_clipboard` gebruikt `win32com.client.Dispatch("Word.Application")`, `Document.Content.Copy()` en `pythoncom.OleFlushClipboard()` om volledige Word-opmaak op het klembord te zetten. Niet testbaar zonder Windows + Microsoft Word.
- **Linux/macOS (dev-fallback)**: extractie van `word/document.xml` via stdlib (`zipfile` + `xml.etree`) en `xclip`/`wl-copy`/`xsel`/`pbcopy` voor het klembord. Alleen platte tekst.

Linux heeft `DISPLAY=:0` nodig voor `xclip` (in deze omgeving is dat al geconfigureerd).

## Een minimaal test-.docx maken zonder externe deps
Een `.docx` is gewoon een ZIP met drie XML-files (`[Content_Types].xml`, `_rels/.rels`, `word/document.xml`). Snippet om een werkend test-bestand te bouwen vanuit Python staat hieronder; LibreOffice kan dit ook converteren naar PDF voor previews.

```python
import zipfile, pathlib
WORD = pathlib.Path("word_files"); WORD.mkdir(exist_ok=True)
target = WORD / "voorbeeld.docx"
CT = '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
RELS = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
DOC = '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Hallo wereld</w:t></w:r></w:p></w:body></w:document>'
with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", CT)
    z.writestr("_rels/.rels", RELS)
    z.writestr("word/document.xml", DOC)
```

## Belangrijke endpoints
- `GET  /api/health` → `{status:"ok", platform:...}`
- `GET  /api/files[?regenerate=true]` → metadata + preview-URLs per `.docx`
- `POST /api/regenerate/{stem}` → previews opnieuw genereren
- `POST /api/copy-content/{stem}` → klembord vullen met document-inhoud. Response bevat `mode` (`"word_com"` op Windows, `"plain_text"` als fallback).

## UI-test recept (Linux-fallback)
Adversarial-checks vereisen dat je het klembord met een sentinel pre-vult per stap, anders kun je niet aantonen dat de actie het klembord daadwerkelijk overschrijft.

```bash
# Test 1: pagina-klik moet document-inhoud kopieren (niet INCLUDETEXT).
printf %s "__SENTINEL__" | xclip -selection clipboard -i
# klik op de pagina-thumbnail in de UI
xclip -selection clipboard -o   # verwacht: tekst uit het document

# Test 2: knop "Kopieer inhoud (plak in Word)".
printf %s "__SENTINEL2__" | xclip -selection clipboard -i
# klik op de blauwe primaire knop
xclip -selection clipboard -o   # verwacht: tekst uit het document

# Test 3: secundaire INCLUDETEXT-knop (regressie).
printf %s "__SENTINEL3__" | xclip -selection clipboard -i
# klik op de INCLUDETEXT-knop
xclip -selection clipboard -o   # verwacht: INCLUDETEXT "<absoluut pad>.docx"
```

## Wat te verwachten in de UI
- Selectie van een pagina-thumbnail krijgt een groene rand via CSS-class `selected` (`--success` border).
- De primaire knop wordt kort groen + tekst "Gekopieerd!" via class `.copied` (1.5s).
- Een toast onderaan toont onderscheid tussen rich-mode (`Inhoud op klembord. ...`) en plain-text-mode (`Alleen tekst gekopieerd...`).

## Wat je niet kunt testen in deze omgeving
- De Windows Word COM-route en het behoud van Word-opmaak. Documenteer dit altijd als `untested` in het rapport.
- Of een al geopend Word-venster open blijft (`doc.Close` zonder `word.Quit`). Niet verifieerbaar zonder Windows.

## Devin Secrets Needed
- Geen. De app is volledig lokaal en heeft geen externe credentials nodig.
