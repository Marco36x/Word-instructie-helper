# Word-instructie-helper

Standalone Python desktop-app die previews toont van alle Word-bestanden in
een map. Met een klik kopieer je het Word `INCLUDETEXT`-commando met het
volledige pad naar je klembord, zodat je het in een Word-veld (`Ctrl+F9`)
kunt plakken om de inhoud van het bronbestand in te voegen.

## Hoe werkt het

- Plaats `.docx`-bestanden in de map `word_files/` naast `main.py`.
- Start de app: `python main.py`.
- Bij iedere start (en bij **Vernieuwen**) worden previews gegenereerd voor
  nieuwe bestanden en preview-mappen van verwijderde bestanden opgeruimd.
- Klik op een preview om het volgende op je klembord te zetten:
  ```
  INCLUDETEXT "C:\\volledig\\pad\\naar\\bestand.docx"
  ```
- In Word: druk `Ctrl+F9` om een veld `{ }` te plaatsen, plak het commando
  ertussen en druk `F9` om de inhoud te laden.

## UI

- **Kolommen** boven in de werkbalk past het aantal kolommen aan
  (minimaal 4, maximaal 12). De keuze wordt opgeslagen in `config.json`.
- **Vernieuwen** scant de map opnieuw en genereert ontbrekende previews.
- **Alles opnieuw genereren** vervangt alle previews door verse versies.
- **Open word_files-map** opent de map in Verkenner / Finder / bestandsbeheer.

## Installatie (Windows)

Vereist Python 3.10+ en geinstalleerde Microsoft Word.

```cmd
git clone https://github.com/Marco36x/Word-instructie-helper.git
cd Word-instructie-helper
python -m venv .venv
.venv\Scripts\activate
python -m pip install -r requirements.txt
python main.py
```

Of zonder venv (gebruikt de globale Python):

```cmd
python -m pip install -r requirements.txt
python main.py
```

## Linux/macOS (fallback voor ontwikkeling)

LibreOffice is nodig voor de `.docx` -> PDF-conversie:

```bash
sudo apt install libreoffice  # Debian/Ubuntu
# of: brew install --cask libreoffice  (macOS)

python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python main.py
```

## Mappenstructuur

```
Word-instructie-helper/
├── main.py            # de desktop-app (Tkinter)
├── requirements.txt
├── config.json        # door de app aangemaakt (kolomvoorkeur)
├── word_files/        # leg hier je .docx-bestanden neer
└── previews/          # door de app aangemaakt: per bestand een submap
                       # met PDF + PNG-previews per pagina
```
