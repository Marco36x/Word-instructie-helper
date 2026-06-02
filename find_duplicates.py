"""
Zoek dubbele Word-bestanden in genummerde submappen.

Gebruik:
    python find_duplicates.py <pad-naar-hoofdmap>

Het script:
 1. Doorzoekt alle submappen waarvan de naam met een cijfer begint.
 2. Zoekt per submap naar .docx / .doc bestanden.
 3. Herkent de datum aan het eind van de bestandsnaam en bepaalt
    zo de "basisnaam" (alles vóór de datum).
 4. Groepeert bestanden met dezelfde basisnaam (= duplicaten).
 5. Genereert een HTML-rapport (dubbele_bestanden.html) in de
    opgegeven hoofdmap.  Het bestand met de nieuwste datum krijgt
    een groene kleur; de overige zijn rood.
"""

import os
import re
import sys
from datetime import datetime
from pathlib import Path

# ---------------------------------------------------------------------------
# Datum-herkenning aan het eind van de bestandsnaam (zonder extensie)
# ---------------------------------------------------------------------------

# Patronen die we herkennen, meest specifiek eerst.
_DATE_PATTERNS: list[tuple[str, str]] = [
    # 2024-01-15  of  2024_01_15  of  2024.01.15
    (r"[\s_\-]*(\d{4})[\-_\.](\d{2})[\-_\.](\d{2})$", "%Y-%m-%d"),
    # 15-01-2024  of  15_01_2024  of  15.01.2024
    (r"[\s_\-]*(\d{2})[\-_\.](\d{2})[\-_\.](\d{4})$", "%d-%m-%Y"),
    # 20240115 (8 cijfers aaneengesloten)
    (r"[\s_\-]*(\d{8})$", "%Y%m%d"),
]


def _extract_date(stem: str) -> tuple[str, datetime | None]:
    """Geeft (basisnaam, datum) terug.  Als er geen datum herkend wordt:
    (volledige stem, None)."""
    for pattern, fmt in _DATE_PATTERNS:
        m = re.search(pattern, stem)
        if m:
            raw = m.group(0)
            digits = m.groups()
            try:
                if fmt == "%Y%m%d":
                    dt = datetime.strptime(digits[0], fmt)
                elif fmt == "%d-%m-%Y":
                    dt = datetime.strptime(
                        f"{digits[0]}-{digits[1]}-{digits[2]}", fmt
                    )
                else:
                    dt = datetime.strptime(
                        f"{digits[0]}-{digits[1]}-{digits[2]}", fmt
                    )
                base = stem[: m.start()].rstrip(" _-")
                return base, dt
            except ValueError:
                continue
    return stem, None


# ---------------------------------------------------------------------------
# Hoofdmap scannen
# ---------------------------------------------------------------------------


def _numbered_dirs(root: Path) -> list[Path]:
    """Alle directe submappen waarvan de naam met een cijfer begint."""
    dirs = []
    for entry in sorted(root.iterdir()):
        if entry.is_dir() and entry.name and entry.name[0].isdigit():
            dirs.append(entry)
    return dirs


def _word_files(directory: Path) -> list[Path]:
    """Alle .docx en .doc bestanden in *directory* (niet recursief).
    Sla tijdelijke Word-bestanden (~$…) over."""
    exts = {".docx", ".doc"}
    return sorted(
        p
        for p in directory.iterdir()
        if p.is_file()
        and p.suffix.lower() in exts
        and not p.name.startswith("~$")
    )


# ---------------------------------------------------------------------------
# Duplicaten detecteren
# ---------------------------------------------------------------------------


def find_duplicates(root: Path) -> dict[Path, dict[str, list[tuple[str, datetime | None]]]]:
    """Retourneert per genummerde submap een dict van basisnaam →
    lijst van (volledige bestandsnaam, datum).  Alleen basisnamen
    met méér dan één bestand (= duplicaten) worden opgenomen."""
    result: dict[Path, dict[str, list[tuple[str, datetime | None]]]] = {}

    for d in _numbered_dirs(root):
        groups: dict[str, list[tuple[str, datetime | None]]] = {}
        for f in _word_files(d):
            stem = f.stem  # bestandsnaam zonder extensie
            base, dt = _extract_date(stem)
            base_lower = base.strip().lower()
            groups.setdefault(base_lower, []).append((f.name, dt))

        # Alleen duplicaten bewaren (>1 bestand met dezelfde basisnaam)
        dupes = {k: v for k, v in groups.items() if len(v) > 1}
        if dupes:
            result[d] = dupes

    return result


# ---------------------------------------------------------------------------
# HTML-rapport genereren
# ---------------------------------------------------------------------------

_HTML_TEMPLATE = """\
<!DOCTYPE html>
<html lang="nl">
<head>
<meta charset="utf-8">
<title>Dubbele Word-bestanden</title>
<style>
  body {{ font-family: Segoe UI, Arial, sans-serif; margin: 2rem; background: #f9f9f9; }}
  h1 {{ color: #333; }}
  h2 {{ color: #555; margin-top: 2rem; border-bottom: 1px solid #ccc; padding-bottom: .3rem; }}
  h3 {{ color: #666; margin-top: 1.2rem; }}
  ul {{ list-style: none; padding-left: 0; }}
  li {{ padding: 4px 8px; margin: 2px 0; border-radius: 4px; }}
  .oudere {{ color: #c0392b; }}
  .nieuwste {{ color: #27ae60; font-weight: bold; }}
  .geen-datum {{ color: #7f8c8d; font-style: italic; }}
  .summary {{ color: #888; font-size: 0.9em; margin-bottom: 2rem; }}
</style>
</head>
<body>
<h1>Dubbele Word-bestanden</h1>
<p class="summary">Gescande map: <code>{root}</code><br>Gegenereerd op: {timestamp}</p>
{content}
</body>
</html>
"""


def _generate_html(
    root: Path,
    duplicates: dict[Path, dict[str, list[tuple[str, datetime | None]]]],
) -> str:
    if not duplicates:
        content = "<p>Geen dubbele bestanden gevonden.</p>"
        return _HTML_TEMPLATE.format(
            root=str(root),
            timestamp=datetime.now().strftime("%Y-%m-%d %H:%M"),
            content=content,
        )

    parts: list[str] = []
    for directory, groups in duplicates.items():
        parts.append(f"<h2>📁 {directory.name}</h2>")
        for _base, files in sorted(groups.items()):
            # Bepaal welk bestand de nieuwste datum heeft
            dated = [(name, dt) for name, dt in files if dt is not None]
            latest_name: str | None = None
            if dated:
                latest_name = max(dated, key=lambda x: x[1])[0]  # type: ignore[arg-type]

            parts.append("<ul>")
            # Sorteer op datum (oudste eerst), bestanden zonder datum onderaan
            sorted_files = sorted(
                files, key=lambda x: (x[1] is None, x[1] or datetime.min)
            )
            for name, dt in sorted_files:
                if dt is None:
                    css = "geen-datum"
                elif name == latest_name:
                    css = "nieuwste"
                else:
                    css = "oudere"
                parts.append(f'  <li class="{css}">{name}</li>')
            parts.append("</ul>")

    content = "\n".join(parts)
    return _HTML_TEMPLATE.format(
        root=str(root),
        timestamp=datetime.now().strftime("%Y-%m-%d %H:%M"),
        content=content,
    )


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def main() -> None:
    if len(sys.argv) < 2:
        print("Gebruik: python find_duplicates.py <pad-naar-hoofdmap>")
        sys.exit(1)

    root = Path(sys.argv[1]).resolve()
    if not root.is_dir():
        print(f"FOUT: '{root}' is geen geldige map.")
        sys.exit(1)

    print(f"Scannen van: {root}")
    duplicates = find_duplicates(root)

    if not duplicates:
        print("Geen dubbele bestanden gevonden.")
        return

    total = sum(len(f) for g in duplicates.values() for f in g.values())
    print(f"{total} dubbele bestanden gevonden in {len(duplicates)} map(pen).")

    output = root / "dubbele_bestanden.html"
    html = _generate_html(root, duplicates)
    output.write_text(html, encoding="utf-8")
    print(f"Rapport opgeslagen: {output}")


if __name__ == "__main__":
    main()
