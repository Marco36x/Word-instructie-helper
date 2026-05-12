# Word-instructie-helper

Lokale Windows-desktop-app om compleet Word-documenten naar je huidige
document over te nemen. De UI draait in een eigen Windows-venster (via
pywebview / WebView2), met op de achtergrond een FastAPI-server op
localhost.

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

Standaard opent de app als native venster (vereist pywebview, op Windows met
WebView2 — al voorgeinstalleerd op Win10/11):
```
pip install -r requirements.txt
python main.py
```

Legacy-modi voor debugging of als de WebView2-runtime ontbreekt:
```
python main.py --web          # start HTTP-server en open de browser
python main.py --no-browser   # alleen HTTP-server (headless)
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
3. Dubbelklik op de .exe. De app opent een eigen Windows-venster met de UI
   (geen browser, geen consolevenster). Onder water draait een lokale
   FastAPI-server op een vrije poort.
4. Klik op een document/pagina, schakel naar je doel-document in Word en plak
   met `Ctrl+V`.

De .exe maakt naast zichzelf een `previews/`-map aan voor gecachte
pagina-previews; deze kan veilig weggegooid worden.

#### Vereisten op de doel-PC
- **WebView2-runtime**: standaard aanwezig op Windows 10 (sinds april 2024)
  en Windows 11. Op oudere installaties: download de Evergreen-installer
  van Microsoft (“WebView2 Runtime”) en installeer eenmalig.
- **Microsoft Word**: nodig voor de rich-format kopieer-route (behoud van
  tabellen/opmaak/afbeeldingen). Zonder Word valt de app terug op platte
  tekst.

#### Opdrachtregel-opties van de .exe
- (geen) of `--desktop`: open in een eigen venster (standaard).
- `--web`: start de HTTP-server en open de browser (legacy).
- `--no-browser`: alleen de HTTP-server zonder venster of browser (voor
  scripting/CI).
