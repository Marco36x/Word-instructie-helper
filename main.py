"""Word Preview App.

Een lokale FastAPI-applicatie die previews genereert van .docx-bestanden in
``word_files/`` (een submap van de programma-directory) en de previews opslaat
in ``previews/``. Via een browser-UI scrol je door de previews en kopieer je
met een klik de inhoud van een Word-document naar het klembord, zodat je
het direct in je eigen Word-document kunt plakken. Als alternatief kun je nog
steeds de Word INCLUDETEXT-instructie naar het klembord kopieren.

Werkt op Windows (gebruikt Microsoft Word via docx2pdf voor previews en via
pywin32/COM voor het kopieren van de inhoud met behoud van opmaak). Op
Linux/macOS is er een fallback (LibreOffice voor previews, platte tekst voor
het klembord) zodat het ook op een niet-Windows-machine te testen is.
"""

from __future__ import annotations

import logging
import platform
import shutil
import subprocess
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from typing import Any

from fastapi import FastAPI, HTTPException, Query
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
logger = logging.getLogger("word_preview")

BASE_DIR = Path(__file__).parent.resolve()
WORD_DIR = BASE_DIR / "word_files"
PREVIEW_DIR = BASE_DIR / "previews"
STATIC_DIR = BASE_DIR / "static"

WORD_DIR.mkdir(exist_ok=True)
PREVIEW_DIR.mkdir(exist_ok=True)

PREVIEW_DPI = 110

app = FastAPI(title="Word Preview App")


def _is_windows() -> bool:
    return platform.system() == "Windows"


def _convert_docx_to_pdf(docx_path: Path, pdf_path: Path) -> None:
    """Convert a .docx to .pdf.

    On Windows we use ``docx2pdf`` which drives Microsoft Word via COM. On
    other platforms we shell out to LibreOffice in headless mode.
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
    """Render every page of ``pdf_path`` to a PNG in ``output_dir``."""
    import fitz  # PyMuPDF

    output_dir.mkdir(parents=True, exist_ok=True)
    for stale in output_dir.glob(f"{base}_page_*.png"):
        stale.unlink()

    paths: list[Path] = []
    with fitz.open(str(pdf_path)) as doc:
        zoom = PREVIEW_DPI / 72
        matrix = fitz.Matrix(zoom, zoom)
        for index, page in enumerate(doc, start=1):
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            png_path = output_dir / f"{base}_page_{index:03d}.png"
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


def _relative_url(path: Path) -> str:
    rel = path.relative_to(PREVIEW_DIR).as_posix()
    return f"/previews/{rel}"


def _existing_previews(docx_path: Path) -> list[Path]:
    base = docx_path.stem
    out_dir = PREVIEW_DIR / base
    if not out_dir.exists():
        return []
    return sorted(out_dir.glob(f"{base}_page_*.png"))


def _list_word_files() -> list[Path]:
    return sorted(
        p
        for p in WORD_DIR.glob("*.docx")
        if not p.name.startswith("~$") and p.is_file()
    )


@app.get("/api/files")
def list_files(regenerate: bool = Query(default=False)) -> list[dict[str, Any]]:
    """Return metadata for every .docx-bestand inclusief preview-URLs."""
    results: list[dict[str, Any]] = []
    for docx in _list_word_files():
        info: dict[str, Any] = {
            "name": docx.name,
            "stem": docx.stem,
            "path": str(docx.resolve()),
            "include_command": _build_include_command(docx),
            "previews": [],
            "error": None,
        }
        try:
            previews = _existing_previews(docx)
            if regenerate or not previews:
                previews = _generate_previews(docx)
            info["previews"] = [_relative_url(p) for p in previews]
        except Exception as exc:  # noqa: BLE001 - surface to UI
            logger.exception("Preview mislukt voor %s", docx.name)
            info["error"] = str(exc)
        results.append(info)
    return results


@app.post("/api/regenerate/{stem}")
def regenerate_one(stem: str) -> dict[str, Any]:
    docx_path = WORD_DIR / f"{stem}.docx"
    if not docx_path.exists():
        raise HTTPException(status_code=404, detail=f"{docx_path.name} niet gevonden")
    try:
        previews = _generate_previews(docx_path)
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(status_code=500, detail=str(exc)) from exc
    return {
        "name": docx_path.name,
        "stem": docx_path.stem,
        "path": str(docx_path.resolve()),
        "include_command": _build_include_command(docx_path),
        "previews": [_relative_url(p) for p in previews],
    }


def _build_include_command(docx_path: Path) -> str:
    """Build the INCLUDETEXT body that goes inside Word's ``{ }``-veld.

    Word vereist dat backslashes in een pad verdubbeld worden, of je kan
    forward slashes gebruiken. We verdubbelen ze omdat dat de standaard
    notatie is op Windows.
    """
    absolute = str(docx_path.resolve())
    escaped = absolute.replace("\\", "\\\\")
    return f'INCLUDETEXT "{escaped}"'


WORD_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _extract_text_from_docx(docx_path: Path) -> str:
    """Extract plain text from a ``.docx`` by parsing ``word/document.xml``.

    Used as a cross-platform fallback when Microsoft Word (via COM) is not
    available. Behoudt geen opmaak; voor volledige opmaak is Windows + Word
    nodig.
    """
    with zipfile.ZipFile(docx_path) as archive:
        try:
            xml_bytes = archive.read("word/document.xml")
        except KeyError as exc:
            raise RuntimeError(
                f"{docx_path.name} bevat geen 'word/document.xml' "
                "en lijkt geen geldig .docx-bestand te zijn."
            ) from exc

    root = ET.fromstring(xml_bytes)
    paragraphs: list[str] = []
    for paragraph in root.iter(f"{WORD_NS}p"):
        runs: list[str] = []
        for node in paragraph.iter():
            tag = node.tag
            if tag == f"{WORD_NS}t" and node.text:
                runs.append(node.text)
            elif tag in (f"{WORD_NS}br", f"{WORD_NS}cr"):
                runs.append("\n")
            elif tag == f"{WORD_NS}tab":
                runs.append("\t")
        paragraphs.append("".join(runs))
    return "\n".join(paragraphs).strip()


def _set_clipboard_text(text: str) -> str:
    """Place ``text`` on the system clipboard. Returns the tool that was used.

    On Windows we use the ``win32clipboard`` API. On macOS we use ``pbcopy``.
    On Linux/BSD we try ``wl-copy``, ``xclip`` en ``xsel`` op volgorde.
    """
    system = platform.system()
    if system == "Windows":
        try:
            import win32clipboard  # type: ignore[import-not-found]
        except ImportError as exc:  # pragma: no cover - depends on platform
            raise RuntimeError(
                "pywin32 niet geinstalleerd. Voer 'pip install pywin32' uit."
            ) from exc
        win32clipboard.OpenClipboard()
        try:
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardText(text, win32clipboard.CF_UNICODETEXT)
        finally:
            win32clipboard.CloseClipboard()
        return "win32clipboard"

    candidates: list[list[str]]
    if system == "Darwin":
        candidates = [["pbcopy"]]
    else:
        candidates = [
            ["wl-copy"],
            ["xclip", "-selection", "clipboard"],
            ["xsel", "--clipboard", "--input"],
        ]

    last_error: Exception | None = None
    for cmd in candidates:
        if not shutil.which(cmd[0]):
            continue
        try:
            subprocess.run(cmd, input=text, text=True, check=True)
            return cmd[0]
        except (subprocess.CalledProcessError, OSError) as exc:
            last_error = exc

    hint = (
        "Installeer 'xclip' (sudo apt install xclip) of 'wl-clipboard'."
        if system == "Linux"
        else "Installeer een clipboard-tool voor je systeem."
    )
    raise RuntimeError(
        f"Geen werkende clipboard-tool gevonden. {hint}"
    ) from last_error


def _copy_docx_to_clipboard(docx_path: Path) -> dict[str, Any]:
    """Plaats de inhoud van ``docx_path`` op het klembord zodat de gebruiker
    het direct in zijn eigen Word-document kan plakken.

    Op Windows openen we Word via COM en kopieren we ``Document.Content``,
    waardoor het volledige document (tekst, opmaak, tabellen, afbeeldingen)
    op het klembord komt in alle formats die Word normaal gebruikt. Na de
    Copy roepen we ``OleFlushClipboard`` aan zodat de klembord-inhoud blijft
    bestaan nadat we het document sluiten.

    Op andere platforms valt de functie terug op platte tekst (alleen voor
    ontwikkelen/testen; opmaak gaat verloren).
    """
    if _is_windows():
        try:
            import pythoncom  # type: ignore[import-not-found]
            import win32com.client  # type: ignore[import-not-found]
        except ImportError as exc:  # pragma: no cover - depends on platform
            raise RuntimeError(
                "pywin32 niet geinstalleerd. Voer 'pip install pywin32' uit."
            ) from exc

        pythoncom.CoInitialize()
        try:
            word = win32com.client.Dispatch("Word.Application")
            previous_alerts = word.DisplayAlerts
            word.DisplayAlerts = 0  # wdAlertsNone
            doc = word.Documents.Open(
                FileName=str(docx_path),
                ReadOnly=True,
                AddToRecentFiles=False,
                Visible=False,
            )
            try:
                doc.Content.Copy()
                # Zorg dat de klembord-inhoud blijft nadat we het document
                # sluiten (anders verdwijnt het bij Close).
                pythoncom.OleFlushClipboard()
            finally:
                doc.Close(SaveChanges=0)  # wdDoNotSaveChanges
                word.DisplayAlerts = previous_alerts
        finally:
            pythoncom.CoUninitialize()
        return {
            "mode": "word_com",
            "format": "rich",
            "note": (
                "Volledige Word-inhoud (met opmaak) staat op het klembord. "
                "Schakel over naar je doel-document en plak met Ctrl+V."
            ),
        }

    text = _extract_text_from_docx(docx_path)
    if not text:
        raise RuntimeError(
            f"Geen tekst gevonden in {docx_path.name}; niets gekopieerd."
        )
    tool = _set_clipboard_text(text)
    return {
        "mode": "plain_text",
        "format": "text",
        "tool": tool,
        "note": (
            "Alleen platte tekst gekopieerd (fallback zonder Microsoft Word). "
            "Op Windows wordt de volledige opmaak meegenomen."
        ),
        "chars": len(text),
    }


@app.post("/api/copy-content/{stem}")
def copy_content(stem: str) -> dict[str, Any]:
    """Plaats de inhoud van het opgegeven .docx-bestand op het klembord."""
    docx_path = WORD_DIR / f"{stem}.docx"
    if not docx_path.exists():
        raise HTTPException(
            status_code=404, detail=f"{docx_path.name} niet gevonden"
        )
    try:
        result = _copy_docx_to_clipboard(docx_path)
    except Exception as exc:  # noqa: BLE001 - surface to UI
        logger.exception("Kopieren naar klembord mislukt voor %s", docx_path.name)
        raise HTTPException(status_code=500, detail=str(exc)) from exc
    return {
        "name": docx_path.name,
        "stem": docx_path.stem,
        "path": str(docx_path.resolve()),
        **result,
    }


@app.get("/api/health")
def health() -> dict[str, str]:
    return {"status": "ok", "platform": platform.system()}


@app.get("/", response_class=HTMLResponse)
def index() -> str:
    index_html = STATIC_DIR / "index.html"
    if not index_html.exists():
        raise HTTPException(status_code=500, detail="static/index.html ontbreekt")
    return index_html.read_text(encoding="utf-8")


app.mount("/previews", StaticFiles(directory=PREVIEW_DIR), name="previews")
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


def main() -> None:
    import argparse
    import webbrowser

    import uvicorn

    parser = argparse.ArgumentParser(description="Word Preview App")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument(
        "--no-browser",
        action="store_true",
        help="Open de browser niet automatisch.",
    )
    args = parser.parse_args()

    url = f"http://{args.host}:{args.port}"
    logger.info("Word-bestanden directory: %s", WORD_DIR)
    logger.info("Previews directory:        %s", PREVIEW_DIR)
    logger.info("Server op %s", url)

    if not args.no_browser:
        try:
            webbrowser.open(url)
        except Exception:  # noqa: BLE001
            pass

    uvicorn.run(app, host=args.host, port=args.port, log_level="info")


if __name__ == "__main__":
    main()
