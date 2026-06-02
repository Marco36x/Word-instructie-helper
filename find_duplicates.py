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
import tkinter as tk
import webbrowser
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, font as tkfont, messagebox, ttk

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
# Tkinter GUI
# ---------------------------------------------------------------------------

_COLOR_NIEUWSTE = "#27ae60"
_COLOR_OUDERE = "#c0392b"
_COLOR_GEEN_DATUM = "#7f8c8d"


class DuplicateFinderApp:
    """Tkinter-venster voor het opsporen van dubbele Word-bestanden."""

    APP_TITLE = "Dubbele Word-bestanden zoeken"

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(self.APP_TITLE)
        self.root.geometry("900x600")
        self.root.minsize(700, 450)

        self._build_styles()
        self._build_widgets()

    # -- Styling ----------------------------------------------------------

    def _build_styles(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        base = tkfont.nametofont("TkDefaultFont")
        base.configure(size=10)
        self.root.option_add("*Font", base)
        style.configure(
            "Header.TLabel", font=(base.actual("family"), 14, "bold")
        )
        style.configure("Sub.TLabel", foreground="#555555")
        style.configure("Primary.TButton", padding=(14, 6))

    # -- Widgets ----------------------------------------------------------

    def _build_widgets(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(2, weight=1)

        # Header
        header = ttk.Frame(self.root, padding=(16, 12, 16, 4))
        header.grid(row=0, column=0, sticky="ew")
        ttk.Label(
            header, text=self.APP_TITLE, style="Header.TLabel"
        ).pack(anchor="w")
        ttk.Label(
            header,
            text="Kies een hoofdmap — het programma zoekt in genummerde "
            "submappen naar Word-bestanden met dezelfde naam.",
            style="Sub.TLabel",
        ).pack(anchor="w")

        # Map-kiezer
        picker = ttk.Frame(self.root, padding=(16, 8, 16, 4))
        picker.grid(row=1, column=0, sticky="ew")
        picker.columnconfigure(1, weight=1)

        ttk.Label(picker, text="Hoofdmap:").grid(
            row=0, column=0, sticky="w", padx=(0, 8)
        )
        self.path_var = tk.StringVar()
        self.path_entry = ttk.Entry(picker, textvariable=self.path_var)
        self.path_entry.grid(row=0, column=1, sticky="ew", padx=(0, 8))

        ttk.Button(
            picker, text="Bladeren…", command=self._on_browse
        ).grid(row=0, column=2, padx=(0, 8))

        self.scan_btn = ttk.Button(
            picker,
            text="Zoek dubbele bestanden",
            style="Primary.TButton",
            command=self._on_scan,
        )
        self.scan_btn.grid(row=0, column=3)

        # Resultaten
        result_frame = ttk.LabelFrame(
            self.root, text="Resultaten", padding=(12, 8)
        )
        result_frame.grid(
            row=2, column=0, sticky="nsew", padx=16, pady=(8, 8)
        )
        result_frame.columnconfigure(0, weight=1)
        result_frame.rowconfigure(0, weight=1)

        self.result_text = tk.Text(
            result_frame,
            wrap="word",
            state="disabled",
            font=("Segoe UI", 10),
            padx=8,
            pady=8,
        )
        scrollbar = ttk.Scrollbar(
            result_frame, orient="vertical", command=self.result_text.yview
        )
        self.result_text.configure(yscrollcommand=scrollbar.set)
        self.result_text.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        # Tags voor kleuren
        self.result_text.tag_configure(
            "nieuwste", foreground=_COLOR_NIEUWSTE, font=("Segoe UI", 10, "bold")
        )
        self.result_text.tag_configure(
            "oudere", foreground=_COLOR_OUDERE
        )
        self.result_text.tag_configure(
            "geen_datum", foreground=_COLOR_GEEN_DATUM, font=("Segoe UI", 10, "italic")
        )
        self.result_text.tag_configure(
            "mapnaam", font=("Segoe UI", 11, "bold"), spacing1=10
        )

        # Statusbalk
        status_frame = ttk.Frame(self.root, padding=(16, 4, 16, 8))
        status_frame.grid(row=3, column=0, sticky="ew")
        status_frame.columnconfigure(0, weight=1)

        self.status_var = tk.StringVar(value="Kies een map om te beginnen.")
        ttk.Label(status_frame, textvariable=self.status_var).grid(
            row=0, column=0, sticky="w"
        )
        self.html_btn = ttk.Button(
            status_frame,
            text="Open HTML-rapport",
            command=self._on_open_html,
        )
        self.html_btn.grid(row=0, column=1, sticky="e")
        self.html_btn.grid_remove()  # verborgen tot er een rapport is

        self._html_path: Path | None = None

    # -- Callbacks --------------------------------------------------------

    def _on_browse(self) -> None:
        path = filedialog.askdirectory(title="Kies de hoofdmap")
        if path:
            self.path_var.set(path)

    def _on_scan(self) -> None:
        raw = self.path_var.get().strip()
        if not raw:
            messagebox.showwarning(
                "Geen map gekozen", "Kies eerst een hoofdmap."
            )
            return
        root = Path(raw).resolve()
        if not root.is_dir():
            messagebox.showerror(
                "Ongeldige map", f"'{root}' is geen geldige map."
            )
            return

        duplicates = find_duplicates(root)
        self._show_results(root, duplicates)

        # Genereer HTML-rapport
        if duplicates:
            output = root / "dubbele_bestanden.html"
            html = _generate_html(root, duplicates)
            output.write_text(html, encoding="utf-8")
            self._html_path = output
            self.html_btn.grid()
            total = sum(
                len(f) for g in duplicates.values() for f in g.values()
            )
            self.status_var.set(
                f"{total} dubbele bestanden gevonden — "
                f"rapport opgeslagen: {output}"
            )
        else:
            self._html_path = None
            self.html_btn.grid_remove()
            self.status_var.set("Geen dubbele bestanden gevonden.")

    def _show_results(
        self,
        root: Path,
        duplicates: dict[Path, dict[str, list[tuple[str, datetime | None]]]],
    ) -> None:
        self.result_text.configure(state="normal")
        self.result_text.delete("1.0", "end")

        if not duplicates:
            self.result_text.insert("end", "Geen dubbele bestanden gevonden.\n")
            self.result_text.configure(state="disabled")
            return

        for directory, groups in duplicates.items():
            self.result_text.insert(
                "end", f"📁 {directory.name}\n", "mapnaam"
            )
            for _base, files in sorted(groups.items()):
                dated = [(n, d) for n, d in files if d is not None]
                latest_name: str | None = None
                if dated:
                    latest_name = max(dated, key=lambda x: x[1])[0]  # type: ignore[arg-type]

                sorted_files = sorted(
                    files,
                    key=lambda x: (x[1] is None, x[1] or datetime.min),
                )
                for name, dt in sorted_files:
                    if dt is None:
                        tag = "geen_datum"
                    elif name == latest_name:
                        tag = "nieuwste"
                    else:
                        tag = "oudere"
                    self.result_text.insert("end", f"   {name}\n", tag)

            self.result_text.insert("end", "\n")

        self.result_text.configure(state="disabled")

    def _on_open_html(self) -> None:
        if self._html_path and self._html_path.exists():
            webbrowser.open(self._html_path.as_uri())


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def _run_cli(path: str) -> None:
    root = Path(path).resolve()
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


# ---------------------------------------------------------------------------
# Entrypoint
# ---------------------------------------------------------------------------


def main() -> None:
    if len(sys.argv) >= 2:
        # CLI-modus: pad als argument
        _run_cli(sys.argv[1])
    else:
        # GUI-modus: Tkinter-venster
        root = tk.Tk()
        DuplicateFinderApp(root)
        root.mainloop()


if __name__ == "__main__":
    main()
