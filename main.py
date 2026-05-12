"""Word Preview App - standalone desktop versie (Tkinter).

Een desktop-app die van alle ``.docx``-bestanden in ``word_files/`` previews
toont (PNG van iedere pagina). Klik op een preview om het Word
``INCLUDETEXT``-commando met het absolute pad naar het klembord te kopieren.
In Word: druk Ctrl+F9 voor een veld ``{ }``, plak het commando ertussen, en
druk F9 om de inhoud te laden.

Werkt op Windows met geinstalleerde Microsoft Word (via ``docx2pdf``). Op
Linux/macOS wordt LibreOffice headless gebruikt als fallback (handig voor
ontwikkeling/test).
"""

from __future__ import annotations

import json
import logging
import os
import platform
import shutil
import subprocess
import sys
import threading
from pathlib import Path
from queue import Empty, Queue
from typing import Iterable

import tkinter as tk
from tkinter import messagebox, simpledialog, ttk

from PIL import Image, ImageTk

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
logger = logging.getLogger("word_preview")

BASE_DIR = Path(__file__).parent.resolve()
WORD_DIR = BASE_DIR / "word_files"
PREVIEW_DIR = BASE_DIR / "previews"
CONFIG_PATH = BASE_DIR / "config.json"

WORD_DIR.mkdir(exist_ok=True)
PREVIEW_DIR.mkdir(exist_ok=True)

PREVIEW_DPI = 110
MIN_COLUMNS = 4
MAX_COLUMNS = 12
DEFAULT_COLUMNS = 4
THUMB_MIN_WIDTH = 150

# ---------------------------------------------------------------------------
# Visueel palet — geinspireerd op moderne document-browsers (Office 365,
# Adobe Bridge, macOS Finder Gallery view). Eén centrale plek zodat het
# uiterlijk consistent blijft.
# ---------------------------------------------------------------------------

COLOR_BG = "#f5f6f8"
COLOR_TOOLBAR_BG = "#ffffff"
COLOR_TOOLBAR_BORDER = "#e2e6ec"
COLOR_CARD_BG = "#ffffff"
COLOR_CARD_BORDER = "#e2e6ec"
COLOR_CARD_BORDER_HOVER = "#94a3b8"
COLOR_CARD_BORDER_SELECTED = "#1d4ed8"
COLOR_CARD_BG_SELECTED = "#eef4ff"
COLOR_TEXT_PRIMARY = "#0f172a"
COLOR_TEXT_SECONDARY = "#64748b"
COLOR_TEXT_MUTED = "#94a3b8"
COLOR_ACCENT = "#1d4ed8"
COLOR_ACCENT_DARK = "#1e40af"
COLOR_DIVIDER = "#e2e6ec"
COLOR_THUMB_FRAME = "#f1f5f9"
COLOR_TOAST_BG = "#0f172a"
COLOR_TOAST_FG = "#f8fafc"
COLOR_BADGE_BG = "#1d4ed8"
COLOR_BADGE_FG = "#ffffff"

FONT_BASE = "Segoe UI"


# ---------------------------------------------------------------------------
# Conversie-pipeline (.docx -> .pdf -> .png)
# ---------------------------------------------------------------------------


def _is_windows() -> bool:
    return platform.system() == "Windows"


def _convert_docx_to_pdf(docx_path: Path, pdf_path: Path) -> None:
    """Convert ``.docx`` to ``.pdf``.

    On Windows we use ``docx2pdf`` (drives Word via COM); on other platforms
    we shell out to LibreOffice in headless mode.
    """
    pdf_path.parent.mkdir(parents=True, exist_ok=True)

    if _is_windows():
        try:
            from docx2pdf import convert  # type: ignore[import-not-found]
        except ImportError as exc:  # pragma: no cover - depends on platform
            raise RuntimeError(
                "docx2pdf is niet geinstalleerd. Voer 'pip install docx2pdf' uit."
            ) from exc
        convert(str(docx_path), str(pdf_path))
        if not pdf_path.exists():
            raise RuntimeError(f"PDF werd niet gegenereerd voor {docx_path}")
        return

    soffice = shutil.which("libreoffice") or shutil.which("soffice")
    if not soffice:
        raise RuntimeError(
            "Geen LibreOffice/soffice gevonden. Op Windows wordt Word gebruikt; "
            "op andere platforms is LibreOffice nodig voor de fallback."
        )
    result = subprocess.run(
        [
            soffice,
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            str(pdf_path.parent),
            str(docx_path),
        ],
        check=False,
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        raise RuntimeError(
            f"LibreOffice-conversie mislukt voor {docx_path.name}: {result.stderr.strip()}"
        )
    generated = pdf_path.parent / f"{docx_path.stem}.pdf"
    if generated != pdf_path:
        if pdf_path.exists():
            pdf_path.unlink()
        generated.rename(pdf_path)


def _render_pdf_to_pngs(pdf_path: Path, output_dir: Path, base: str) -> list[Path]:
    """Render alleen de **eerste pagina** van ``pdf_path`` naar een PNG.

    De applicatie toont per ``.docx`` één preview (de voorkant); oudere
    pagina-bestanden van vorige versies worden bij elke generatie netjes
    opgeruimd zodat er geen ongebruikte PNG's blijven slingeren.
    """
    import fitz  # PyMuPDF

    output_dir.mkdir(parents=True, exist_ok=True)
    for stale in output_dir.glob(f"{base}_page_*.png"):
        stale.unlink()

    paths: list[Path] = []
    with fitz.open(str(pdf_path)) as doc:
        if len(doc) == 0:
            return paths
        zoom = PREVIEW_DPI / 72
        matrix = fitz.Matrix(zoom, zoom)
        page = doc[0]
        pix = page.get_pixmap(matrix=matrix, alpha=False)
        png_path = output_dir / f"{base}_page_001.png"
        pix.save(str(png_path))
        paths.append(png_path)
    return paths


def _generate_previews(docx_path: Path) -> list[Path]:
    base = docx_path.stem
    out_dir = PREVIEW_DIR / base
    out_dir.mkdir(parents=True, exist_ok=True)
    pdf_path = out_dir / f"{base}.pdf"
    logger.info("Genereer previews voor %s", docx_path.name)
    _convert_docx_to_pdf(docx_path, pdf_path)
    return _render_pdf_to_pngs(pdf_path, out_dir, base)


def _existing_previews(docx_path: Path) -> list[Path]:
    """Geef alleen de preview van pagina 1 terug (eventuele oudere extra
    pagina-bestanden worden bij de volgende generatie opgeschoond)."""
    base = docx_path.stem
    out_dir = PREVIEW_DIR / base
    if not out_dir.exists():
        return []
    first = out_dir / f"{base}_page_001.png"
    if first.exists():
        return [first]
    # Fallback: oude generaties gebruikten een andere nummering — pak de
    # alfabetisch eerste preview die er nog ligt.
    existing = sorted(out_dir.glob(f"{base}_page_*.png"))
    return existing[:1]


def _list_word_files() -> list[Path]:
    return sorted(
        p
        for p in WORD_DIR.glob("*.docx")
        if not p.name.startswith("~$") and p.is_file()
    )


def _cleanup_orphan_previews(active_stems: set[str]) -> list[str]:
    """Verwijder preview-mappen waarvoor geen ``.docx`` meer bestaat."""
    removed: list[str] = []
    if not PREVIEW_DIR.exists():
        return removed
    for entry in PREVIEW_DIR.iterdir():
        if not entry.is_dir() or entry.name in active_stems:
            continue
        try:
            shutil.rmtree(entry)
            removed.append(entry.name)
            logger.info("Preview-map verwijderd voor verdwenen bestand: %s", entry.name)
        except OSError as exc:
            logger.warning("Kon preview-map %s niet verwijderen: %s", entry, exc)
    return removed


def _build_include_command(docx_path: Path) -> str:
    """Bouw het ``INCLUDETEXT``-commando dat tussen Word-veldhaakjes hoort.

    Word vereist verdubbelde backslashes in een pad, of forward slashes; we
    verdubbelen omdat dat de standaardnotatie is op Windows.
    """
    absolute = str(docx_path.resolve())
    escaped = absolute.replace("\\", "\\\\")
    return f'INCLUDETEXT "{escaped}"'


# ---------------------------------------------------------------------------
# Configuratie (kolomvoorkeur)
# ---------------------------------------------------------------------------


def _load_config() -> dict[str, object]:
    if not CONFIG_PATH.exists():
        return {}
    try:
        return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        logger.warning("Kon config niet lezen: %s", exc)
        return {}


def _save_config(data: dict[str, object]) -> None:
    try:
        CONFIG_PATH.write_text(json.dumps(data, indent=2), encoding="utf-8")
    except OSError as exc:
        logger.warning("Kon config niet opslaan: %s", exc)


def _load_columns() -> int:
    raw = _load_config().get("columns", DEFAULT_COLUMNS)
    try:
        value = int(raw)  # type: ignore[arg-type]
    except (TypeError, ValueError):
        value = DEFAULT_COLUMNS
    return max(MIN_COLUMNS, min(MAX_COLUMNS, value))


def _save_columns(value: int) -> None:
    config = _load_config()
    config["columns"] = int(value)
    _save_config(config)


def _load_names() -> dict[str, str]:
    raw = _load_config().get("names", {}) or {}
    if not isinstance(raw, dict):
        return {}
    return {str(k): str(v) for k, v in raw.items() if isinstance(v, str) and v}


def _save_name(stem: str, custom_name: str | None) -> None:
    config = _load_config()
    names = config.get("names", {}) or {}
    if not isinstance(names, dict):
        names = {}
    if custom_name:
        names[stem] = custom_name
    else:
        names.pop(stem, None)
    config["names"] = names
    _save_config(config)


# ---------------------------------------------------------------------------
# Helper om bestand/folder te openen op het systeem
# ---------------------------------------------------------------------------


def _open_path(path: Path) -> None:
    if sys.platform.startswith("win"):
        os.startfile(str(path))  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------


class WordPreviewApp(tk.Tk):
    """Tkinter desktop-app."""

    def __init__(self) -> None:
        super().__init__()
        self.title("Word Preview")
        self.geometry("1280x800")
        self.minsize(900, 560)

        self.columns_var = tk.IntVar(value=_load_columns())
        self.status_var = tk.StringVar(value="Klaar.")
        self._photo_refs: list[ImageTk.PhotoImage] = []
        self._busy = False
        self._files: list[Path] = []
        self._task_queue: Queue[tuple[str, object]] = Queue()
        self._render_after_id: str | None = None
        self._last_canvas_width = 0
        self._selectable_cells: list[tk.Frame] = []
        self._selected_cell: tk.Frame | None = None
        self._selected_badge: tk.Label | None = None
        self._names: dict[str, str] = _load_names()
        self._toast: tk.Toplevel | None = None
        self._toast_label: tk.Label | None = None
        self._toast_after_id: str | None = None

        self._configure_style()
        self._build_ui()

        self.after(50, self._poll_task_queue)
        self.after(100, lambda: self.refresh(generate_missing=True))

    # ------------------------------------------------------------------
    # Style / layout
    # ------------------------------------------------------------------

    def _configure_style(self) -> None:
        # Op Windows zorgen voor scherpe rendering bij display scaling > 100%.
        if _is_windows():
            try:
                from ctypes import windll  # type: ignore[attr-defined]
                windll.shcore.SetProcessDpiAwareness(1)
            except Exception:
                pass

        self.configure(bg=COLOR_BG)

        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        # Generieke ttk-defaults zodat we niet voor elk widget een aparte
        # stijl hoeven te definieren.
        style.configure(".", font=(FONT_BASE, 10))
        style.configure("TFrame", background=COLOR_BG)
        style.configure("TLabel", background=COLOR_BG,
                        foreground=COLOR_TEXT_PRIMARY)
        style.configure("TSeparator", background=COLOR_DIVIDER)

        # Toolbar
        style.configure("Toolbar.TFrame", background=COLOR_TOOLBAR_BG)
        style.configure(
            "AppTitle.TLabel", background=COLOR_TOOLBAR_BG,
            foreground=COLOR_TEXT_PRIMARY, font=(FONT_BASE, 14, "bold"),
            padding=(4, 2),
        )
        style.configure(
            "Toolbar.TLabel", background=COLOR_TOOLBAR_BG,
            foreground=COLOR_TEXT_SECONDARY, font=(FONT_BASE, 10),
            padding=(4, 2),
        )

        # Primaire actie-knop (blauw)
        style.configure(
            "Primary.TButton", background=COLOR_ACCENT, foreground="#ffffff",
            borderwidth=0, focusthickness=0, padding=(14, 7),
            font=(FONT_BASE, 10),
        )
        style.map(
            "Primary.TButton",
            background=[
                ("active", COLOR_ACCENT_DARK),
                ("disabled", "#94a3b8"),
            ],
            foreground=[("disabled", "#e2e8f0")],
        )

        # Secundaire knop (grijstinten)
        style.configure(
            "Secondary.TButton", background="#f1f5f9",
            foreground=COLOR_TEXT_PRIMARY, borderwidth=0, focusthickness=0,
            padding=(14, 7), font=(FONT_BASE, 10),
        )
        style.map(
            "Secondary.TButton",
            background=[
                ("active", "#e2e8f0"),
                ("disabled", "#f8fafc"),
            ],
            foreground=[("disabled", "#94a3b8")],
        )

        # Compacte knop voor in een card
        style.configure(
            "CardAction.TButton", background="#f1f5f9",
            foreground=COLOR_TEXT_PRIMARY, borderwidth=0, focusthickness=0,
            padding=(10, 4), font=(FONT_BASE, 9),
        )
        style.map(
            "CardAction.TButton",
            background=[("active", "#e2e8f0")],
            foreground=[("disabled", "#94a3b8")],
        )

        # Spinbox
        style.configure(
            "Modern.TSpinbox", fieldbackground="#ffffff",
            background="#ffffff", foreground=COLOR_TEXT_PRIMARY,
            bordercolor=COLOR_DIVIDER, arrowsize=12, padding=(6, 4),
            relief="flat",
        )

        # Status-bar
        style.configure(
            "StatusBar.TFrame", background=COLOR_TOOLBAR_BG,
        )
        style.configure(
            "Status.TLabel", background=COLOR_TOOLBAR_BG,
            foreground=COLOR_TEXT_SECONDARY, padding=(12, 6),
            font=(FONT_BASE, 9),
        )
        style.configure(
            "StatusMeta.TLabel", background=COLOR_TOOLBAR_BG,
            foreground=COLOR_TEXT_MUTED, padding=(12, 6),
            font=(FONT_BASE, 9),
        )

    def _build_ui(self) -> None:
        # ---------------- Toolbar ----------------
        toolbar_wrap = tk.Frame(self, bg=COLOR_TOOLBAR_BG)
        toolbar_wrap.pack(side=tk.TOP, fill=tk.X)

        toolbar = ttk.Frame(toolbar_wrap, style="Toolbar.TFrame",
                            padding=(20, 14))
        toolbar.pack(fill=tk.X)

        # Logo / titel
        title_box = ttk.Frame(toolbar, style="Toolbar.TFrame")
        title_box.pack(side=tk.LEFT, padx=(0, 24))
        ttk.Label(title_box, text="Word Preview",
                  style="AppTitle.TLabel").pack(anchor="w")
        ttk.Label(
            title_box,
            text="INCLUDETEXT-helper voor Word-velden",
            style="Toolbar.TLabel",
        ).pack(anchor="w")

        # Primaire acties
        self.refresh_btn = ttk.Button(
            toolbar, text="Vernieuwen", style="Primary.TButton",
            command=lambda: self.refresh(generate_missing=True),
        )
        self.refresh_btn.pack(side=tk.LEFT, padx=(0, 8))

        self.regen_btn = ttk.Button(
            toolbar, text="Genereer alles opnieuw",
            style="Secondary.TButton", command=self.regenerate_all,
        )
        self.regen_btn.pack(side=tk.LEFT, padx=(0, 16))

        ttk.Separator(toolbar, orient="vertical").pack(
            side=tk.LEFT, fill=tk.Y, padx=(0, 16),
        )

        # Rechtergedeelte: kolomkeuze + map openen
        right_box = ttk.Frame(toolbar, style="Toolbar.TFrame")
        right_box.pack(side=tk.RIGHT)

        self.open_dir_btn = ttk.Button(
            right_box, text="Open word_files-map",
            style="Secondary.TButton", command=self._open_word_dir,
        )
        self.open_dir_btn.pack(side=tk.RIGHT, padx=(16, 0))

        ttk.Label(right_box, text="Kolommen", style="Toolbar.TLabel").pack(
            side=tk.LEFT, padx=(0, 6),
        )
        self.col_spin = ttk.Spinbox(
            right_box, from_=MIN_COLUMNS, to=MAX_COLUMNS, width=4,
            textvariable=self.columns_var, command=self._on_columns_change,
            style="Modern.TSpinbox", justify="center",
        )
        self.col_spin.pack(side=tk.LEFT)
        self.col_spin.bind("<FocusOut>", lambda _e: self._on_columns_change())
        self.col_spin.bind("<Return>", lambda _e: self._on_columns_change())

        # Subtiele scheidingslijn onder de toolbar
        tk.Frame(self, bg=COLOR_TOOLBAR_BORDER, height=1).pack(
            side=tk.TOP, fill=tk.X,
        )

        # ---------------- Body: scrollable canvas ----------------
        body = tk.Frame(self, bg=COLOR_BG)
        body.pack(fill=tk.BOTH, expand=True)

        self.canvas = tk.Canvas(body, highlightthickness=0, background=COLOR_BG,
                                bd=0)
        v_scroll = ttk.Scrollbar(body, orient="vertical",
                                 command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=v_scroll.set)

        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.inner = tk.Frame(self.canvas, bg=COLOR_BG, padx=20, pady=20)
        self._inner_window = self.canvas.create_window(
            (0, 0), window=self.inner, anchor="nw",
        )

        self.inner.bind(
            "<Configure>",
            lambda _e: self.canvas.configure(scrollregion=self.canvas.bbox("all")),
        )
        self.canvas.bind("<Configure>", self._on_canvas_resize)

        # Mouse wheel scrolling (Windows + Linux)
        self.canvas.bind("<Enter>", lambda _e: self._bind_wheel())
        self.canvas.bind("<Leave>", lambda _e: self._unbind_wheel())

        # ---------------- Footer / status bar ----------------
        tk.Frame(self, bg=COLOR_TOOLBAR_BORDER, height=1).pack(
            side=tk.BOTTOM, fill=tk.X,
        )
        status_bar = ttk.Frame(self, style="StatusBar.TFrame")
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        self.meta_var = tk.StringVar(value="")
        ttk.Label(
            status_bar, textvariable=self.status_var, style="Status.TLabel",
            anchor="w",
        ).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Label(
            status_bar, textvariable=self.meta_var, style="StatusMeta.TLabel",
            anchor="e",
        ).pack(side=tk.RIGHT)

    # ------------------------------------------------------------------
    # Scrolling helpers
    # ------------------------------------------------------------------

    def _bind_wheel(self) -> None:
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel, add="+")
        self.canvas.bind_all("<Button-4>", lambda _e: self.canvas.yview_scroll(-1, "units"), add="+")
        self.canvas.bind_all("<Button-5>", lambda _e: self.canvas.yview_scroll(1, "units"), add="+")

    def _unbind_wheel(self) -> None:
        self.canvas.unbind_all("<MouseWheel>")
        self.canvas.unbind_all("<Button-4>")
        self.canvas.unbind_all("<Button-5>")

    def _on_mousewheel(self, event: tk.Event) -> None:  # type: ignore[type-arg]
        delta = -1 if getattr(event, "delta", 0) > 0 else 1
        self.canvas.yview_scroll(delta, "units")

    def _on_canvas_resize(self, event: tk.Event) -> None:  # type: ignore[type-arg]
        self.canvas.itemconfigure(self._inner_window, width=event.width)
        if abs(event.width - self._last_canvas_width) > 40:
            self._last_canvas_width = event.width
            self._schedule_relayout()

    def _schedule_relayout(self, delay_ms: int = 150) -> None:
        if self._busy:
            return
        if self._render_after_id is not None:
            try:
                self.after_cancel(self._render_after_id)
            except (ValueError, tk.TclError):
                pass
        self._render_after_id = self.after(delay_ms, self._render_all)

    # ------------------------------------------------------------------
    # Toolbar callbacks
    # ------------------------------------------------------------------

    def _on_columns_change(self) -> None:
        try:
            value = int(self.columns_var.get())
        except (tk.TclError, ValueError):
            value = DEFAULT_COLUMNS
        value = max(MIN_COLUMNS, min(MAX_COLUMNS, value))
        if value != self.columns_var.get():
            self.columns_var.set(value)
        _save_columns(value)
        self._schedule_relayout(delay_ms=20)

    def _open_word_dir(self) -> None:
        try:
            _open_path(WORD_DIR)
        except OSError as exc:
            messagebox.showerror("Map openen mislukt", str(exc))

    def _set_busy(self, busy: bool) -> None:
        self._busy = busy
        state = "disabled" if busy else "!disabled"
        for btn in (self.refresh_btn, self.regen_btn, self.open_dir_btn):
            btn.state([state])
        self.col_spin.state([state])

    # ------------------------------------------------------------------
    # Refresh / generate
    # ------------------------------------------------------------------

    def refresh(self, *, generate_missing: bool = True) -> None:
        if self._busy:
            return
        word_files = _list_word_files()
        removed = _cleanup_orphan_previews({p.stem for p in word_files})
        if removed:
            logger.info("Opgeruimd: %s", ", ".join(removed))

        to_generate = (
            [p for p in word_files if not _existing_previews(p)]
            if generate_missing
            else []
        )

        self._files = word_files
        self._update_meta()
        if not to_generate:
            self._render_all()
            removed_msg = f" {len(removed)} verlopen opgeruimd." if removed else ""
            self.status_var.set(
                f"{len(word_files)} bestand(en) geladen.{removed_msg}"
            )
            return

        self._render_all()  # toon huidige stand alvast
        self._run_generate_in_background(to_generate, label="Genereren")

    def _update_meta(self) -> None:
        count = len(self._files)
        if count == 0:
            self.meta_var.set("")
        elif count == 1:
            self.meta_var.set("1 bestand")
        else:
            self.meta_var.set(f"{count} bestanden")

    def regenerate_all(self) -> None:
        if self._busy:
            return
        word_files = _list_word_files()
        _cleanup_orphan_previews({p.stem for p in word_files})
        self._files = word_files
        if not word_files:
            self.status_var.set("Geen bestanden om te genereren.")
            self._render_all()
            return
        self._run_generate_in_background(word_files, label="Opnieuw genereren")

    def _regenerate_one(self, docx: Path) -> None:
        if self._busy or not docx.exists():
            return
        self._run_generate_in_background([docx], label=f"Opnieuw: {docx.name}")

    def _run_generate_in_background(
        self, files: list[Path], *, label: str
    ) -> None:
        self._set_busy(True)
        total = len(files)
        self.status_var.set(f"{label}... (0/{total})")

        def worker() -> None:
            for index, docx in enumerate(files, start=1):
                self._task_queue.put(
                    ("progress", (index, total, label, docx.name))
                )
                try:
                    _generate_previews(docx)
                    self._task_queue.put(("ok", docx.name))
                except Exception as exc:  # noqa: BLE001
                    logger.exception("Preview mislukt voor %s", docx.name)
                    self._task_queue.put(("error", (docx.name, str(exc))))
            self._task_queue.put(("done", None))

        threading.Thread(target=worker, daemon=True).start()

    def _poll_task_queue(self) -> None:
        try:
            while True:
                kind, payload = self._task_queue.get_nowait()
                if kind == "progress":
                    index, total, label, name = payload  # type: ignore[misc]
                    self.status_var.set(f"{label} ({index}/{total}): {name}")
                elif kind == "ok":
                    pass
                elif kind == "error":
                    name, msg = payload  # type: ignore[misc]
                    self.status_var.set(f"Fout: {name} — {msg}")
                elif kind == "done":
                    self._set_busy(False)
                    self._files = _list_word_files()
                    _cleanup_orphan_previews({p.stem for p in self._files})
                    self._render_all()
                    self._update_meta()
                    self.status_var.set(
                        f"{len(self._files)} bestand(en) klaar."
                    )
        except Empty:
            pass
        finally:
            self.after(80, self._poll_task_queue)

    # ------------------------------------------------------------------
    # Rendering
    # ------------------------------------------------------------------

    def _render_all(self) -> None:
        self._render_after_id = None
        for widget in self.inner.winfo_children():
            widget.destroy()
        self._photo_refs.clear()
        self._selectable_cells.clear()
        self._selected_cell = None
        self._selected_badge = None

        if not self._files:
            self._render_empty_state()
            return

        cols = max(MIN_COLUMNS, int(self.columns_var.get() or DEFAULT_COLUMNS))
        # Bepaal de bruikbare breedte op basis van de inner-frame als die al
        # zichtbaar is, anders het canvas, anders het hoofdvenster. Sommige
        # Windows-thema's geven pas een betrouwbare breedte na de eerste
        # mainloop-tick, dus updaten we eerst de geometrie.
        self.update_idletasks()
        candidates = [
            self.inner.winfo_width(),
            self.canvas.winfo_width(),
            self.winfo_width(),
        ]
        canvas_width = max((w for w in candidates if w > 1), default=1200)
        # Reserveer ruimte voor binnen-padding van .inner + iets voor de
        # scrollbar. Daarna delen we door het aantal kolommen.
        usable = max(canvas_width - 40, cols * (THUMB_MIN_WIDTH + 32))
        cell_outer = max(THUMB_MIN_WIDTH + 32, usable // cols)
        thumb_width = cell_outer - 32  # interne padding van card
        thumb_height = int(thumb_width * 1.414)

        # Eén grid-container voor alle bestanden samen. Elke .docx krijgt
        # één kaart. Vaste minsize per kolom + weight=0 zorgt dat de kaarten
        # ALTIJD naast elkaar staan, ongeacht Tk-thema of DPI.
        grid_container = tk.Frame(self.inner, bg=COLOR_BG)
        grid_container.pack(anchor="nw", fill=tk.NONE, expand=False)

        for c in range(cols):
            grid_container.columnconfigure(c, weight=0, minsize=cell_outer)

        for index, docx in enumerate(self._files):
            row, col = divmod(index, cols)
            card = self._build_file_card(grid_container, docx,
                                         thumb_width, thumb_height)
            card.grid(row=row, column=col, padx=8, pady=8, sticky="nw")

    def _render_empty_state(self) -> None:
        placeholder = tk.Frame(self.inner, bg=COLOR_BG, padx=40, pady=40)
        placeholder.pack(fill=tk.BOTH, expand=True)
        tk.Label(
            placeholder, text="\U0001F4C4",  # 📄
            bg=COLOR_BG, fg=COLOR_TEXT_MUTED,
            font=(FONT_BASE, 56),
        ).pack(pady=(40, 8))
        tk.Label(
            placeholder, text="Nog geen Word-bestanden gevonden",
            bg=COLOR_BG, fg=COLOR_TEXT_PRIMARY,
            font=(FONT_BASE, 14, "bold"),
        ).pack()
        tk.Label(
            placeholder,
            text=(
                "Plaats je .docx-bestanden in de map word_files\\\n"
                "en klik op 'Vernieuwen' om de previews te genereren."
            ),
            bg=COLOR_BG, fg=COLOR_TEXT_SECONDARY,
            font=(FONT_BASE, 10), justify="center",
        ).pack(pady=(6, 16))
        ttk.Button(
            placeholder, text="Open word_files-map",
            style="Primary.TButton", command=self._open_word_dir,
        ).pack()

    def _build_file_card(
        self, parent: tk.Misc, docx: Path, thumb_w: int, thumb_h: int,
    ) -> tk.Widget:
        card_bg = COLOR_CARD_BG
        # 2px highlight blijft constant zodat de layout niet schuift bij
        # selectie; alleen de kleur verandert.
        card = tk.Frame(
            parent, bg=card_bg, highlightthickness=2,
            highlightbackground=COLOR_CARD_BORDER,
            highlightcolor=COLOR_CARD_BORDER_SELECTED,
            cursor="hand2",
        )

        # ---- Thumbnail ----
        thumb_frame = tk.Frame(
            card, bg=COLOR_THUMB_FRAME, width=thumb_w, height=thumb_h,
        )
        thumb_frame.pack(padx=12, pady=(12, 8))
        thumb_frame.pack_propagate(False)

        previews = _existing_previews(docx)
        thumb_widget: tk.Widget
        if previews:
            try:
                img = Image.open(previews[0]).convert("RGB")
                img.thumbnail((thumb_w, thumb_h), Image.LANCZOS)
                photo = ImageTk.PhotoImage(img)
                self._photo_refs.append(photo)
                thumb_widget = tk.Label(
                    thumb_frame, image=photo, bg=COLOR_THUMB_FRAME,
                    cursor="hand2", bd=0,
                )
            except OSError as exc:
                logger.exception("Kon preview niet laden: %s", previews[0])
                thumb_widget = tk.Label(
                    thumb_frame, text=f"Fout\n{exc}",
                    fg="#b91c1c", bg=COLOR_THUMB_FRAME, justify="center",
                )
        else:
            thumb_widget = tk.Label(
                thumb_frame,
                text="\u23F3\nNog geen preview\nKlik 'Vernieuwen'",
                fg=COLOR_TEXT_SECONDARY, bg=COLOR_THUMB_FRAME,
                justify="center", font=(FONT_BASE, 9),
            )
        thumb_widget.place(relx=0.5, rely=0.5, anchor="center")

        # ---- Naam ----
        display = self._display_name(docx)
        name_label = tk.Label(
            card, text=display, bg=card_bg, fg=COLOR_TEXT_PRIMARY,
            font=(FONT_BASE, 10, "bold"), wraplength=thumb_w,
            justify="center", cursor="hand2",
        )
        name_label.pack(padx=12, pady=(0, 2))

        subtitle: tk.Label | None = None
        if display != docx.stem:
            subtitle = tk.Label(
                card, text=docx.name, bg=card_bg, fg=COLOR_TEXT_SECONDARY,
                font=(FONT_BASE, 8), wraplength=thumb_w,
                justify="center", cursor="hand2",
            )
            subtitle.pack(padx=12, pady=(0, 4))

        # ---- Knoppenrij ----
        actions = tk.Frame(card, bg=card_bg)
        actions.pack(padx=10, pady=(6, 12))
        ttk.Button(
            actions, text="Naam", style="CardAction.TButton",
            command=lambda p=docx: self._rename_file(p),
        ).pack(side=tk.LEFT, padx=2)
        ttk.Button(
            actions, text="Vernieuwen", style="CardAction.TButton",
            command=lambda p=docx: self._regenerate_one(p),
        ).pack(side=tk.LEFT, padx=2)

        # ---- "Op klembord"-badge (alleen zichtbaar op geselecteerde card) ----
        badge = tk.Label(
            card, text=" \u2713 Op klembord ", bg=COLOR_BADGE_BG,
            fg=COLOR_BADGE_FG, font=(FONT_BASE, 8, "bold"),
            padx=6, pady=2,
        )
        # badge nog niet placen — pas tonen na klik

        self._selectable_cells.append(card)
        cmd = _build_include_command(docx)

        def apply_bg_recursive(target: tk.Widget, bg: str) -> None:
            try:
                target.configure(bg=bg)
            except tk.TclError:
                pass
            for child in target.winfo_children():
                if isinstance(child, ttk.Widget):
                    continue
                if child is thumb_frame or child is thumb_widget:
                    # Thumbnail-frame houdt altijd zijn eigen lichte tint.
                    continue
                if child is badge:
                    continue
                apply_bg_recursive(child, bg)

        def set_selected(selected: bool) -> None:
            if selected:
                card.configure(highlightbackground=COLOR_CARD_BORDER_SELECTED)
                apply_bg_recursive(card, COLOR_CARD_BG_SELECTED)
                badge.place(relx=1.0, rely=0.0, x=-10, y=10, anchor="ne")
                badge.lift()
            else:
                card.configure(highlightbackground=COLOR_CARD_BORDER)
                apply_bg_recursive(card, card_bg)
                try:
                    badge.place_forget()
                except tk.TclError:
                    pass

        def on_click(_event: tk.Event | None = None) -> None:  # type: ignore[type-arg]
            previous = self._selected_cell
            if previous is not None and previous is not card:
                handler = getattr(previous, "_set_selected", None)
                if callable(handler):
                    handler(False)
            set_selected(True)
            self._selected_cell = card
            self._selected_badge = badge
            self._copy_command(cmd, docx)

        # Stash handler op de cell zodat _render_all/andere callbacks hem
        # netjes kunnen aanroepen.
        card._set_selected = set_selected  # type: ignore[attr-defined]

        clickable = [card, thumb_frame, thumb_widget, name_label, actions]
        if subtitle is not None:
            clickable.append(subtitle)
        for widget in clickable:
            widget.bind("<Button-1>", on_click)

        def on_enter(_e: tk.Event) -> None:  # type: ignore[type-arg]
            if self._selected_cell is card:
                return
            card.configure(highlightbackground=COLOR_CARD_BORDER_HOVER)

        def on_leave(_e: tk.Event) -> None:  # type: ignore[type-arg]
            if self._selected_cell is card:
                return
            card.configure(highlightbackground=COLOR_CARD_BORDER)

        card.bind("<Enter>", on_enter)
        card.bind("<Leave>", on_leave)

        return card

    # ------------------------------------------------------------------
    # Custom names
    # ------------------------------------------------------------------

    def _display_name(self, docx: Path) -> str:
        return self._names.get(docx.stem) or docx.stem

    def _rename_file(self, docx: Path) -> None:
        if self._busy:
            return
        current = self._names.get(docx.stem, "")
        prompt = (
            f"Geef een eigen naam voor:\n{docx.name}\n\n"
            "Laat leeg om de bestandsnaam weer te gebruiken."
        )
        new_name = simpledialog.askstring(
            "Naam aanpassen", prompt, parent=self,
            initialvalue=current,
        )
        if new_name is None:
            return  # geannuleerd
        new_name = new_name.strip()
        _save_name(docx.stem, new_name or None)
        self._names = _load_names()
        self._render_all()
        if new_name:
            self.status_var.set(
                f"Naam voor {docx.name} aangepast naar '{new_name}'."
            )
        else:
            self.status_var.set(f"Aangepaste naam voor {docx.name} verwijderd.")

    # ------------------------------------------------------------------
    # Klembord + duidelijke bevestiging
    # ------------------------------------------------------------------

    def _copy_command(self, command: str, docx: Path, _page: int = 1) -> None:
        self.clipboard_clear()
        self.clipboard_append(command)
        self.update()  # zorgt dat de waarde echt op het systeem-klembord staat
        display = self._display_name(docx)
        self.status_var.set(f"Op klembord: {display}")
        self._show_toast(
            f"\u2713 INCLUDETEXT gekopieerd\n{display}"
        )

    # ------------------------------------------------------------------
    # Toast (bevestigingsmelding onderin)
    # ------------------------------------------------------------------

    def _ensure_toast(self) -> None:
        if self._toast is not None and self._toast.winfo_exists():
            return
        toast = tk.Toplevel(self)
        toast.overrideredirect(True)
        toast.transient(self)
        toast.attributes("-topmost", True)
        try:
            # Klikken/typen mag niet de toast als focus opeisen op Windows.
            toast.attributes("-toolwindow", True)
        except tk.TclError:
            pass
        toast.configure(bg=COLOR_TOAST_BG)
        label = tk.Label(
            toast, text="", bg=COLOR_TOAST_BG, fg=COLOR_TOAST_FG,
            font=(FONT_BASE, 11, "bold"), padx=28, pady=14, justify="center",
        )
        label.pack()
        self._toast = toast
        self._toast_label = label
        toast.withdraw()

    def _show_toast(self, message: str) -> None:
        self._ensure_toast()
        assert self._toast is not None and self._toast_label is not None
        self._toast_label.configure(text=message)
        self._toast.update_idletasks()

        # Plaats midden-onderaan het hoofdvenster.
        try:
            self.update_idletasks()
            root_x = self.winfo_rootx()
            root_y = self.winfo_rooty()
            root_w = self.winfo_width()
            root_h = self.winfo_height()
            tw = self._toast.winfo_reqwidth()
            th = self._toast.winfo_reqheight()
            x = root_x + (root_w - tw) // 2
            y = root_y + root_h - th - 60
            self._toast.geometry(f"+{x}+{y}")
        except tk.TclError:
            pass
        self._toast.deiconify()
        self._toast.lift()

        if self._toast_after_id is not None:
            try:
                self.after_cancel(self._toast_after_id)
            except (ValueError, tk.TclError):
                pass
        self._toast_after_id = self.after(2200, self._hide_toast)

    def _hide_toast(self) -> None:
        self._toast_after_id = None
        if self._toast is not None and self._toast.winfo_exists():
            try:
                self._toast.withdraw()
            except tk.TclError:
                pass


# ---------------------------------------------------------------------------
# Entrypoint
# ---------------------------------------------------------------------------


def main() -> None:
    app = WordPreviewApp()
    try:
        app.mainloop()
    finally:
        app.destroy()


if __name__ == "__main__":
    main()
