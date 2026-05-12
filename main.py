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
THUMB_MIN_WIDTH = 130


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
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Toolbar.TFrame", background="#1f4ed8")
        style.configure(
            "Toolbar.TButton",
            background="#1f4ed8",
            foreground="#ffffff",
            padding=(10, 6),
            relief="flat",
        )
        style.map(
            "Toolbar.TButton",
            background=[("active", "#1839a6")],
            foreground=[("disabled", "#cfd8ef")],
        )
        style.configure(
            "Toolbar.TLabel",
            background="#1f4ed8",
            foreground="#ffffff",
            padding=(6, 0),
        )
        style.configure(
            "Toolbar.TSpinbox",
            fieldbackground="#ffffff",
            arrowsize=14,
            padding=2,
        )
        style.configure("FileTitle.TLabel", font=("Segoe UI", 11, "bold"))
        style.configure("Status.TLabel", padding=(8, 4), foreground="#5b6471")

    def _build_ui(self) -> None:
        toolbar = ttk.Frame(self, style="Toolbar.TFrame", padding=(12, 8))
        toolbar.pack(side=tk.TOP, fill=tk.X)

        ttk.Label(
            toolbar, text="Word Preview", style="Toolbar.TLabel",
            font=("Segoe UI", 12, "bold"),
        ).pack(side=tk.LEFT, padx=(0, 16))

        self.refresh_btn = ttk.Button(
            toolbar, text="Vernieuwen", style="Toolbar.TButton",
            command=lambda: self.refresh(generate_missing=True),
        )
        self.refresh_btn.pack(side=tk.LEFT, padx=(0, 6))

        self.regen_btn = ttk.Button(
            toolbar, text="Alles opnieuw genereren", style="Toolbar.TButton",
            command=self.regenerate_all,
        )
        self.regen_btn.pack(side=tk.LEFT, padx=(0, 12))

        ttk.Label(toolbar, text="Kolommen:", style="Toolbar.TLabel").pack(side=tk.LEFT)
        self.col_spin = ttk.Spinbox(
            toolbar,
            from_=MIN_COLUMNS, to=MAX_COLUMNS, width=4,
            textvariable=self.columns_var, command=self._on_columns_change,
            style="Toolbar.TSpinbox",
        )
        self.col_spin.pack(side=tk.LEFT, padx=(6, 12))
        self.col_spin.bind("<FocusOut>", lambda _e: self._on_columns_change())
        self.col_spin.bind("<Return>", lambda _e: self._on_columns_change())

        self.open_dir_btn = ttk.Button(
            toolbar, text="Open word_files-map", style="Toolbar.TButton",
            command=self._open_word_dir,
        )
        self.open_dir_btn.pack(side=tk.LEFT)

        # Body: scrollable canvas
        body = ttk.Frame(self)
        body.pack(fill=tk.BOTH, expand=True)

        self.canvas = tk.Canvas(body, highlightthickness=0, background="#f4f5f7")
        v_scroll = ttk.Scrollbar(body, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=v_scroll.set)

        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.inner = ttk.Frame(self.canvas, padding=12)
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

        # Footer / status
        status_bar = ttk.Frame(self)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        ttk.Separator(status_bar, orient="horizontal").pack(fill=tk.X)
        ttk.Label(
            status_bar, textvariable=self.status_var, style="Status.TLabel",
            anchor="w",
        ).pack(side=tk.LEFT, fill=tk.X, expand=True)

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
        if not to_generate:
            self._render_all()
            removed_msg = f" Opgeruimd: {len(removed)}." if removed else ""
            self.status_var.set(
                f"{len(word_files)} bestand(en) geladen.{removed_msg}"
            )
            return

        self._render_all()  # toon huidige stand alvast
        self._run_generate_in_background(to_generate, label="Genereren")

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

        if not self._files:
            placeholder = ttk.Frame(self.inner, padding=24)
            placeholder.pack(fill=tk.BOTH, expand=True)
            ttk.Label(
                placeholder,
                text=(
                    "Geen .docx-bestanden gevonden in de map word_files/.\n"
                    "Plaats Word-bestanden in die map en klik op Vernieuwen."
                ),
                justify="center",
            ).pack()
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
        usable = max(canvas_width - 40, cols * (THUMB_MIN_WIDTH + 12))
        thumb_width = max(THUMB_MIN_WIDTH, (usable // cols) - 12)
        thumb_height = int(thumb_width * 1.414)

        # Eén grid-container voor alle bestanden samen. Elke .docx krijgt
        # één cel (eerste pagina als preview + naam + knoppen). Door elke
        # kolom een vaste minimale breedte te geven en weight=0 te houden
        # staan de cellen ALTIJD naast elkaar, ongeacht Tk-thema of DPI.
        grid_container = tk.Frame(self.inner, bg=self["bg"])
        grid_container.pack(anchor="nw", fill=tk.NONE, expand=False,
                            padx=4, pady=4)

        cell_w = thumb_width + 12
        for c in range(cols):
            grid_container.columnconfigure(c, weight=0, minsize=cell_w)

        for index, docx in enumerate(self._files):
            row, col = divmod(index, cols)
            card = self._build_file_card(grid_container, docx,
                                         thumb_width, thumb_height)
            card.grid(row=row, column=col, padx=6, pady=6, sticky="nw")

    def _build_file_card(
        self, parent: tk.Misc, docx: Path, thumb_w: int, thumb_h: int,
    ) -> tk.Widget:
        cell_bg = "#fafbfc"
        cell = tk.Frame(
            parent, bg=cell_bg, highlightthickness=3,
            highlightbackground="#d8dbe0", highlightcolor="#1f4ed8",
            cursor="hand2",
        )

        # Thumbnail van de eerste pagina (of placeholder als die ontbreekt)
        previews = _existing_previews(docx)
        thumb_widget: tk.Widget
        if previews:
            try:
                img = Image.open(previews[0]).convert("RGB")
                img.thumbnail((thumb_w, thumb_h), Image.LANCZOS)
                photo = ImageTk.PhotoImage(img)
                self._photo_refs.append(photo)
                thumb_widget = tk.Label(
                    cell, image=photo, bg=cell_bg, cursor="hand2",
                )
            except OSError as exc:
                logger.exception("Kon preview niet laden: %s", previews[0])
                thumb_widget = tk.Label(
                    cell, text=f"[fout]\n{exc}",
                    fg="#b91c1c", bg=cell_bg, width=20, height=10,
                )
        else:
            thumb_widget = tk.Label(
                cell, text="(nog geen preview)\nklik 'Opnieuw'",
                fg="#5b6471", bg=cell_bg, justify="center",
                width=max(16, thumb_w // 9),
                height=max(10, thumb_h // 20),
            )
        thumb_widget.pack(padx=6, pady=(6, 4))

        display = self._display_name(docx)
        name_label = tk.Label(
            cell, text=display, bg=cell_bg, fg="#0e1726",
            font=("Segoe UI", 10, "bold"), wraplength=thumb_w,
            justify="center", cursor="hand2",
        )
        name_label.pack(padx=6, pady=(0, 0))

        subtitle: tk.Label | None = None
        if display != docx.stem:
            subtitle = tk.Label(
                cell, text=docx.name, bg=cell_bg, fg="#5b6471",
                font=("Segoe UI", 8), wraplength=thumb_w,
                justify="center", cursor="hand2",
            )
            subtitle.pack(padx=6, pady=(0, 0))

        buttons = tk.Frame(cell, bg=cell_bg)
        buttons.pack(padx=4, pady=(6, 8))
        ttk.Button(
            buttons, text="Naam aanpassen",
            command=lambda p=docx: self._rename_file(p),
        ).pack(side=tk.LEFT, padx=2)
        ttk.Button(
            buttons, text="Opnieuw",
            command=lambda p=docx: self._regenerate_one(p),
        ).pack(side=tk.LEFT, padx=2)

        self._selectable_cells.append(cell)
        cmd = _build_include_command(docx)

        def reset_cell_colors(target: tk.Widget, bg: str) -> None:
            """Pas de achtergrondkleur recursief aan op cell + kinderen,
            zodat ook de buttons-frame meegaat in de selectie-highlight."""
            try:
                target.configure(bg=bg)
            except tk.TclError:
                pass
            for child in target.winfo_children():
                # ttk.Button heeft geen bg-attribute; daar gaan we niet in mee.
                if isinstance(child, ttk.Widget):
                    continue
                reset_cell_colors(child, bg)

        def on_click(_event: tk.Event | None = None) -> None:  # type: ignore[type-arg]
            self._copy_command(cmd, docx, 1)
            for other in self._selectable_cells:
                try:
                    other.configure(highlightbackground="#d8dbe0")
                except tk.TclError:
                    pass
                reset_cell_colors(other, cell_bg)
            # Markeer de aangeklikte cel met een groene rand + lichtgroene
            # achtergrond zodat helder is welke preview op het klembord staat.
            cell.configure(highlightbackground="#15803d")
            reset_cell_colors(cell, "#dcfce7")

        clickable = [cell, thumb_widget, name_label, buttons]
        if subtitle is not None:
            clickable.append(subtitle)
        for widget in clickable:
            widget.bind("<Button-1>", on_click)

        def on_enter(_e: tk.Event) -> None:  # type: ignore[type-arg]
            current = str(cell.cget("highlightbackground"))
            if current != "#15803d":
                cell.configure(highlightbackground="#1f4ed8")

        def on_leave(_e: tk.Event) -> None:  # type: ignore[type-arg]
            current = str(cell.cget("highlightbackground"))
            if current != "#15803d":
                cell.configure(highlightbackground="#d8dbe0")

        cell.bind("<Enter>", on_enter)
        cell.bind("<Leave>", on_leave)

        return cell

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
        self.status_var.set(
            f"INCLUDETEXT gekopieerd voor '{display}'."
        )
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
        toast.configure(bg="#15803d")
        label = tk.Label(
            toast, text="", bg="#15803d", fg="#ffffff",
            font=("Segoe UI", 12, "bold"), padx=24, pady=12, justify="center",
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
