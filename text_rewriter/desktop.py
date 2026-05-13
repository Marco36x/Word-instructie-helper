"""Text Rewriter - desktop (Tkinter).

Standalone Python-desktopapp die ruwe tekst herschrijft naar ondubbelzinnige,
duidelijke technische tekst via een gratis OpenRouter-model. Geen webbrowser
nodig.

Start:
    pip install -r requirements.txt
    python desktop.py

Tkinter zit ingebakken in de standaard Python-installatie op Windows, dus er
zijn geen extra GUI-dependencies. Voor het aanroepen van OpenRouter wordt
``httpx`` gebruikt (zie ``requirements.txt``).

Later met PyInstaller in een losse ``.exe`` pakken kan met:
    pyinstaller --onefile --noconsole --name "TextRewriter" desktop.py
"""

from __future__ import annotations

import logging
import os
import queue
import sys
import threading
import tkinter as tk
import webbrowser
from pathlib import Path
from tkinter import font as tkfont
from tkinter import messagebox, ttk

from openrouter_client import (
    DEFAULT_MODEL,
    FREE_MODEL_CHOICES,
    OpenRouterError,
    RewriteResult,
    load_dotenv,
    rewrite_text,
    save_api_key_to_dotenv,
)

logging.basicConfig(
    level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s"
)
logger = logging.getLogger("text_rewriter.desktop")


BASE_DIR = Path(__file__).parent.resolve()


class RewriterApp:
    """Hoofdklasse van de desktop-UI."""

    APP_TITLE = "Text Rewriter"
    INSTR_PLACEHOLDER = (
        "Bv. 'gebruik de wij-vorm' of 'maak er een stapsgewijze instructie van'"
    )

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(self.APP_TITLE)
        self.root.geometry("1100x680")
        self.root.minsize(820, 520)

        # Wordt vanuit de worker-thread gevuld; de UI-thread pollt 'm.
        self._result_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self._worker: threading.Thread | None = None

        load_dotenv(BASE_DIR)

        self._build_styles()
        self._build_widgets()

        # Periodiek de resultaat-queue legen op de UI-thread.
        self.root.after(100, self._drain_queue)
        # Ctrl+Enter shortcut binnen het invoerveld.
        self.input_text.bind("<Control-Return>", lambda _e: self._on_rewrite())
        # Esc om instructie-placeholder weg te halen wanneer focus erin valt.
        self.instructions_var.trace_add("write", self._instructions_var_changed)

    # -- UI ---------------------------------------------------------------
    def _build_styles(self) -> None:
        style = ttk.Style(self.root)
        try:
            # 'clam' staat overal beschikbaar en heeft duidelijke kleuren.
            style.theme_use("clam")
        except tk.TclError:
            pass
        base_font = tkfont.nametofont("TkDefaultFont")
        base_font.configure(size=10)
        self.root.option_add("*Font", base_font)

        style.configure("Header.TLabel", font=(base_font.actual("family"), 14, "bold"))
        style.configure("Sub.TLabel", foreground="#555555")
        style.configure("Status.TLabel", foreground="#0a7d3b")
        style.configure("StatusError.TLabel", foreground="#b00020")
        style.configure("Primary.TButton", padding=(14, 6))
        style.configure("TButton", padding=(10, 5))

    def _build_widgets(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(2, weight=1)

        # --- Header -----------------------------------------------------
        header = ttk.Frame(self.root, padding=(16, 12, 16, 4))
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)

        ttk.Label(header, text="Text Rewriter", style="Header.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            header,
            text="Maak ruwe tekst ondubbelzinnig en technisch helder.",
            style="Sub.TLabel",
        ).grid(row=1, column=0, sticky="w")

        # --- Configuratiebalk (API key + model + instructies) ----------
        config = ttk.LabelFrame(
            self.root, text="Configuratie", padding=(12, 8, 12, 10)
        )
        config.grid(row=1, column=0, sticky="ew", padx=16, pady=(8, 4))
        config.columnconfigure(1, weight=1)
        config.columnconfigure(3, weight=1)

        ttk.Label(config, text="API key").grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.api_key_var = tk.StringVar(
            value=os.environ.get("OPENROUTER_API_KEY", "")
        )
        self.api_key_entry = ttk.Entry(
            config, textvariable=self.api_key_var, show="*"
        )
        self.api_key_entry.grid(row=0, column=1, sticky="ew", padx=(0, 8))

        self.show_key_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            config,
            text="Toon",
            variable=self.show_key_var,
            command=self._toggle_show_key,
        ).grid(row=0, column=2, sticky="w", padx=(0, 8))

        self.save_key_btn = ttk.Button(
            config,
            text="Opslaan in .env",
            command=self._on_save_key,
        )
        self.save_key_btn.grid(row=0, column=3, sticky="w")

        ttk.Label(config, text="Model").grid(
            row=1, column=0, sticky="w", padx=(0, 8), pady=(8, 0)
        )
        self.model_var = tk.StringVar()
        self.model_id_by_label: dict[str, str] = {
            m["label"]: m["id"] for m in FREE_MODEL_CHOICES
        }
        model_labels = [m["label"] for m in FREE_MODEL_CHOICES]
        default_label = next(
            (m["label"] for m in FREE_MODEL_CHOICES if m["id"] == DEFAULT_MODEL),
            model_labels[0],
        )
        self.model_var.set(default_label)
        self.model_combo = ttk.Combobox(
            config,
            values=model_labels,
            textvariable=self.model_var,
            state="readonly",
        )
        self.model_combo.grid(
            row=1, column=1, columnspan=3, sticky="ew", pady=(8, 0)
        )

        ttk.Label(config, text="Extra aanwijzingen").grid(
            row=2, column=0, sticky="w", padx=(0, 8), pady=(8, 0)
        )
        self.instructions_var = tk.StringVar()
        self.instructions_entry = ttk.Entry(
            config, textvariable=self.instructions_var
        )
        self.instructions_entry.grid(
            row=2, column=1, columnspan=3, sticky="ew", pady=(8, 0)
        )
        self._set_instructions_placeholder()
        self.instructions_entry.bind("<FocusIn>", self._instructions_focus_in)
        self.instructions_entry.bind("<FocusOut>", self._instructions_focus_out)

        # --- Twee-koloms tekst ------------------------------------------
        panes = ttk.Frame(self.root, padding=(16, 4, 16, 4))
        panes.grid(row=2, column=0, sticky="nsew")
        panes.columnconfigure(0, weight=1, uniform="cols")
        panes.columnconfigure(1, weight=1, uniform="cols")
        panes.rowconfigure(1, weight=1)

        ttk.Label(panes, text="Invoer", style="Header.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(panes, text="Resultaat", style="Header.TLabel").grid(
            row=0, column=1, sticky="w", padx=(12, 0)
        )

        input_frame = ttk.Frame(panes)
        input_frame.grid(row=1, column=0, sticky="nsew", pady=(4, 0))
        input_frame.columnconfigure(0, weight=1)
        input_frame.rowconfigure(0, weight=1)
        self.input_text = tk.Text(
            input_frame, wrap="word", undo=True, font=("Segoe UI", 10)
        )
        self.input_text.grid(row=0, column=0, sticky="nsew")
        input_scroll = ttk.Scrollbar(
            input_frame, orient="vertical", command=self.input_text.yview
        )
        input_scroll.grid(row=0, column=1, sticky="ns")
        self.input_text.configure(yscrollcommand=input_scroll.set)

        output_frame = ttk.Frame(panes)
        output_frame.grid(row=1, column=1, sticky="nsew", padx=(12, 0), pady=(4, 0))
        output_frame.columnconfigure(0, weight=1)
        output_frame.rowconfigure(0, weight=1)
        self.output_text = tk.Text(
            output_frame,
            wrap="word",
            state="disabled",
            background="#fafafa",
            font=("Segoe UI", 10),
        )
        self.output_text.grid(row=0, column=0, sticky="nsew")
        output_scroll = ttk.Scrollbar(
            output_frame, orient="vertical", command=self.output_text.yview
        )
        output_scroll.grid(row=0, column=1, sticky="ns")
        self.output_text.configure(yscrollcommand=output_scroll.set)

        # --- Actiebalk --------------------------------------------------
        actions = ttk.Frame(self.root, padding=(16, 8, 16, 12))
        actions.grid(row=3, column=0, sticky="ew")
        actions.columnconfigure(99, weight=1)

        self.rewrite_btn = ttk.Button(
            actions,
            text="Herschrijf tekst",
            style="Primary.TButton",
            command=self._on_rewrite,
        )
        self.rewrite_btn.grid(row=0, column=0, padx=(0, 6))

        self.clear_btn = ttk.Button(
            actions, text="Wissen", command=self._on_clear
        )
        self.clear_btn.grid(row=0, column=1, padx=6)

        self.copy_btn = ttk.Button(
            actions,
            text="Kopieer resultaat",
            command=self._on_copy_result,
            state="disabled",
        )
        self.copy_btn.grid(row=0, column=2, padx=6)

        self.status_label = ttk.Label(actions, text="", style="Status.TLabel")
        self.status_label.grid(row=0, column=3, padx=(12, 0), sticky="w")

        self.model_info_label = ttk.Label(actions, text="", style="Sub.TLabel")
        self.model_info_label.grid(row=0, column=99, sticky="e")

        # --- Footer -----------------------------------------------------
        footer = ttk.Frame(self.root, padding=(16, 0, 16, 10))
        footer.grid(row=4, column=0, sticky="ew")
        footer.columnconfigure(0, weight=1)
        ttk.Label(
            footer,
            text=(
                "Geen API key? Maak een gratis aan op openrouter.ai/keys. "
                "De key wordt lokaal opgeslagen in text_rewriter/.env."
            ),
            style="Sub.TLabel",
        ).grid(row=0, column=0, sticky="w")
        link = ttk.Label(
            footer,
            text="openrouter.ai/keys",
            foreground="#0050b3",
            cursor="hand2",
        )
        link.grid(row=0, column=1, sticky="e")
        link.bind(
            "<Button-1>", lambda _e: webbrowser.open("https://openrouter.ai/keys")
        )

    # -- Placeholder-gedrag voor het instructie-veld ----------------------
    def _set_instructions_placeholder(self) -> None:
        if not self.instructions_var.get():
            self.instructions_entry.configure(foreground="#888888")
            # placeholder via insert want we willen niet dat de StringVar 'm
            # ziet als 'echte' inhoud.
            self.instructions_entry.delete(0, "end")
            self.instructions_entry.insert(0, self.INSTR_PLACEHOLDER)
            self._instructions_is_placeholder = True
        else:
            self.instructions_entry.configure(foreground="#000000")
            self._instructions_is_placeholder = False

    def _instructions_focus_in(self, _event: tk.Event) -> None:
        if getattr(self, "_instructions_is_placeholder", False):
            self.instructions_entry.delete(0, "end")
            self.instructions_entry.configure(foreground="#000000")
            self._instructions_is_placeholder = False

    def _instructions_focus_out(self, _event: tk.Event) -> None:
        if not self.instructions_entry.get():
            self._set_instructions_placeholder()

    def _instructions_var_changed(self, *_args: object) -> None:
        # Als de gebruiker iets typt, is het geen placeholder meer.
        if (
            getattr(self, "_instructions_is_placeholder", False)
            and self.instructions_entry.get() != self.INSTR_PLACEHOLDER
        ):
            self._instructions_is_placeholder = False
            self.instructions_entry.configure(foreground="#000000")

    # -- Acties -----------------------------------------------------------
    def _toggle_show_key(self) -> None:
        self.api_key_entry.configure(show="" if self.show_key_var.get() else "*")

    def _on_save_key(self) -> None:
        key = self.api_key_var.get().strip()
        if not key:
            messagebox.showwarning(
                self.APP_TITLE, "Vul eerst een API key in."
            )
            return
        try:
            env_path = save_api_key_to_dotenv(BASE_DIR, key)
        except OSError as exc:
            messagebox.showerror(
                self.APP_TITLE, f"Kon .env niet schrijven: {exc}"
            )
            return
        os.environ["OPENROUTER_API_KEY"] = key
        self._set_status(
            f"API key opgeslagen in {env_path.name}.", error=False
        )

    def _on_clear(self) -> None:
        if self._worker is not None and self._worker.is_alive():
            return
        self.input_text.delete("1.0", "end")
        self._set_output("")
        self.copy_btn.configure(state="disabled")
        self.model_info_label.configure(text="")
        self._set_status("", error=False)

    def _on_copy_result(self) -> None:
        text = self.output_text.get("1.0", "end-1c")
        if not text.strip():
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        # Tk verwacht dat de event-loop draait, anders verdwijnt de clipboard
        # zodra het venster sluit. Wij draaien gewoon, dus dit is voldoende.
        original = self.copy_btn.cget("text")
        self.copy_btn.configure(text="Gekopieerd!")
        self.root.after(1500, lambda: self.copy_btn.configure(text=original))

    def _on_rewrite(self) -> None:
        if self._worker is not None and self._worker.is_alive():
            return
        text = self.input_text.get("1.0", "end-1c").strip()
        if not text:
            self._set_status("Voer eerst tekst in.", error=True)
            return
        api_key = self.api_key_var.get().strip()
        if not api_key:
            self._set_status(
                "Geen API key. Vul 'm in en klik 'Opslaan in .env'.",
                error=True,
            )
            return

        model = self.model_id_by_label.get(self.model_var.get()) or DEFAULT_MODEL
        instructions = (
            ""
            if getattr(self, "_instructions_is_placeholder", False)
            else self.instructions_var.get().strip()
        )

        self._set_buttons_busy(True)
        self._set_status(f"Bezig met herschrijven via {model}...", error=False)
        self.copy_btn.configure(state="disabled")
        self.model_info_label.configure(text="")

        self._worker = threading.Thread(
            target=self._worker_rewrite,
            args=(text, api_key, model, instructions),
            daemon=True,
        )
        self._worker.start()

    def _worker_rewrite(
        self,
        text: str,
        api_key: str,
        model: str,
        instructions: str,
    ) -> None:
        def progress(current_model: str, attempt: int, total: int) -> None:
            self._result_queue.put(
                ("progress", (current_model, attempt, total))
            )

        try:
            result = rewrite_text(
                text,
                api_key=api_key,
                model=model,
                instructions=instructions or None,
                progress=progress,
            )
        except OpenRouterError as exc:
            self._result_queue.put(("error", exc))
            return
        except Exception as exc:  # noqa: BLE001 - alles eindigt in de queue
            logger.exception("Onverwachte fout in rewrite-thread")
            self._result_queue.put(("error", OpenRouterError(str(exc))))
            return
        self._result_queue.put(("done", result))

    def _drain_queue(self) -> None:
        try:
            while True:
                kind, payload = self._result_queue.get_nowait()
                if kind == "progress":
                    current_model, attempt, total = payload  # type: ignore[misc]
                    self._set_status(
                        f"Probeer {current_model} ({attempt}/{total})...",
                        error=False,
                    )
                elif kind == "done":
                    assert isinstance(payload, RewriteResult)
                    self._handle_done(payload)
                elif kind == "error":
                    assert isinstance(payload, OpenRouterError)
                    self._handle_error(payload)
        except queue.Empty:
            pass
        finally:
            self.root.after(100, self._drain_queue)

    def _handle_done(self, result: RewriteResult) -> None:
        self._set_output(result.rewritten)
        self.copy_btn.configure(state="normal")
        if result.fallback_used:
            self._set_status(
                "Klaar (gekozen model was rate-limited; fallback gebruikt).",
                error=False,
            )
            self.model_info_label.configure(
                text=f"Model: {result.model}  (fallback van {result.requested_model})"
            )
        else:
            self._set_status("Klaar.", error=False)
            self.model_info_label.configure(text=f"Model: {result.model}")
        self._set_buttons_busy(False)

    def _handle_error(self, exc: OpenRouterError) -> None:
        self._set_status(f"Fout: {exc}", error=True)
        self._set_buttons_busy(False)

    # -- Helpers ----------------------------------------------------------
    def _set_output(self, text: str) -> None:
        self.output_text.configure(state="normal")
        self.output_text.delete("1.0", "end")
        if text:
            self.output_text.insert("1.0", text)
        self.output_text.configure(state="disabled")

    def _set_status(self, text: str, *, error: bool) -> None:
        self.status_label.configure(
            text=text, style="StatusError.TLabel" if error else "Status.TLabel"
        )

    def _set_buttons_busy(self, busy: bool) -> None:
        state = "disabled" if busy else "normal"
        self.rewrite_btn.configure(state=state)
        self.clear_btn.configure(state=state)
        self.model_combo.configure(state="disabled" if busy else "readonly")


def main(argv: list[str] | None = None) -> int:
    _ = argv  # geen CLI-opties (yet)
    root = tk.Tk()
    RewriterApp(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
