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

## Starten (vanuit broncode)
```
pip install -r requirements.txt
python main.py
```

## Standalone .exe (Windows)

Voor distributie zonder Python-install kun je een enkele `Word-instructie-helper.exe`
bouwen met PyInstaller. Microsoft Word moet wel op de doel-PC geinstalleerd
zijn voor de Word COM-route (rijke opmaak); zonder Word valt de app terug
op platte tekst.

### Optie A: kant-en-klare build via GitHub Actions (aanbevolen)
1. Kopieer eenmalig het voorbeeld in [`docs/build-windows-exe.yml.example`](docs/build-windows-exe.yml.example)
   naar `.github/workflows/build-windows-exe.yml` en commit dat. (Dit is een
   apart bestand omdat workflows niet automatisch door een externe app
   toegevoegd kunnen worden — je moet 'm zelf committen.)
2. Open de Actions-tab op GitHub en kies de workflow **Build Windows executable**.
3. Klik **Run workflow** (of push naar `main` / push een tag `vX.Y.Z`).
4. Na ~5-10 min staat `Word-instructie-helper-windows` als artifact onder de run.
   Bij een `v*`-tag wordt de .exe automatisch aan de GitHub Release gehangen.

### Optie B: lokaal bouwen op Windows
```
pip install -r requirements.txt pyinstaller
pyinstaller --clean --noconfirm Word-instructie-helper.spec
```
Het resultaat staat in `dist/Word-instructie-helper.exe` (~60 MB).

### Gebruik op de doel-PC
1. Plaats `Word-instructie-helper.exe` in een eigen map.
2. Maak naast de .exe een submap `word_files/` (of laat de .exe deze automatisch
   aanmaken bij eerste start) en plaats daar je `.docx`-bestanden.
3. Dubbelklik op de .exe — de FastAPI-server start en je standaardbrowser
   opent automatisch op `http://127.0.0.1:8765`.
4. Klik op een document/pagina, schakel naar je doel-document in Word en plak
   met `Ctrl+V`.

De .exe maakt naast zichzelf een `previews/`-map aan voor gecachte
pagina-previews; deze kan veilig weggegooid worden.
