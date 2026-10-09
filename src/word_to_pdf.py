"""Convert Microsoft Word documents to PDF using an installed office suite."""

from pathlib import Path
import shutil
import subprocess
import tempfile


SUPPORTED_WORD_EXTENSIONS = {".doc", ".docx"}


class WordToPdfError(RuntimeError):
    """Raised when no available backend can convert a Word document."""


class BackendUnavailable(RuntimeError):
    """Raised when an optional conversion backend is not installed."""


def convert_word_to_pdf(source_path, output_path):
    """Convert *source_path* to *output_path*, preferring Microsoft Word."""
    source = Path(source_path).resolve()
    output = Path(output_path).resolve()

    if source.suffix.lower() not in SUPPORTED_WORD_EXTENSIONS:
        raise ValueError("Select a .doc or .docx Word document.")
    if not source.is_file():
        raise FileNotFoundError(f"Word document not found: {source}")
    if output.suffix.lower() != ".pdf":
        raise ValueError("The output file must use the .pdf extension.")

    output.parent.mkdir(parents=True, exist_ok=True)
    errors = []
    for converter in (_convert_with_microsoft_word, _convert_with_libreoffice):
        try:
            converter(source, output)
            if not output.is_file() or output.stat().st_size == 0:
                raise RuntimeError("The converter did not create a PDF file.")
            return output
        except Exception as exc:
            errors.append(str(exc))

    details = "\n\n".join(error for error in errors if error)
    raise WordToPdfError(
        "Could not convert the Word document. Install Microsoft Word or LibreOffice."
        + (f"\n\nDetails:\n{details}" if details else "")
    )


def _convert_with_microsoft_word(source, output):
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise BackendUnavailable("Microsoft Word integration is not available.") from exc

    pythoncom.CoInitialize()
    word = None
    document = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        document = word.Documents.Open(str(source), ReadOnly=True)
        document.ExportAsFixedFormat(str(output), 17)  # wdExportFormatPDF
    except Exception as exc:
        raise RuntimeError(f"Microsoft Word failed: {exc}") from exc
    finally:
        if document is not None:
            document.Close(False)
        if word is not None:
            word.Quit()
        pythoncom.CoUninitialize()


def _find_libreoffice():
    executable = shutil.which("soffice") or shutil.which("libreoffice")
    if executable:
        return executable

    for candidate in (
        Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
        Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
    ):
        if candidate.is_file():
            return str(candidate)
    return None


def _convert_with_libreoffice(source, output):
    executable = _find_libreoffice()
    if executable is None:
        raise BackendUnavailable("LibreOffice is not installed or was not found.")

    with tempfile.TemporaryDirectory(prefix="word-to-pdf-") as temp_directory:
        result = subprocess.run(
            [
                executable,
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                temp_directory,
                str(source),
            ],
            capture_output=True,
            text=True,
            timeout=180,
            check=False,
        )
        converted = Path(temp_directory) / f"{source.stem}.pdf"
        if result.returncode != 0 or not converted.is_file():
            message = (result.stderr or result.stdout).strip()
            raise RuntimeError(f"LibreOffice failed: {message or 'no PDF was created'}")
        shutil.copy2(converted, output)
