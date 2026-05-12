# Word-instructie-helper

Lokale FastAPI-app om compleet Word-documenten naar je huidige document over
te nemen.

## Wat doet het
- Bladert door alle `.docx`-bestanden in `word_files/` en toont per bestand een
  preview per pagina.
- Met een klik op een pagina (of op de knop **Kopieer inhoud**) wordt het
  hele document op je klembord gezet. Schakel daarna over naar je
  doel-document in Word en plak met `Ctrl+V`.
- Alternatief: de knop **INCLUDETEXT** kopieert een Word-veld-instructie
  (`INCLUDETEXT "pad\\naar\\bestand.docx"`) die je in Word tussen `{ }`
  plakt na `Ctrl+F9`.

## Platforms
- **Windows** (aanbevolen): gebruikt Microsoft Word via COM (`pywin32`) zodat
  alle opmaak, tabellen en afbeeldingen behouden blijven bij het plakken.
- **Linux/macOS**: ontwikkel-fallback die alleen platte tekst op het klembord
  zet (vereist `xclip`/`wl-copy`/`pbcopy`).

## Starten
```
pip install -r requirements.txt
python main.py
```
