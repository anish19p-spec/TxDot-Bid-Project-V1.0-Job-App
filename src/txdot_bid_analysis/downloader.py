"""Validated, content-addressed storage of original public bid PDFs."""
from dataclasses import replace
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import re
from tempfile import NamedTemporaryFile
from urllib.parse import urlsplit

import pdfplumber
from requests import RequestException, Session

from .config import Settings
from .errors import DownloadError
from .models import ProjectSource


def download_pdf(session: Session, project: ProjectSource, settings: Settings) -> ProjectSource:
    """Stream, structurally validate and preserve source bytes without duplicate copies."""
    source = urlsplit(project.pdf_url or "")
    if (source.scheme != "https" or source.netloc != urlsplit(settings.listing_url).netloc
            or source.username or source.password):
        raise DownloadError("PDF URL must use the configured TxDOT HTTPS origin.")
    if not re.fullmatch(r"[A-Za-z0-9_-]+", project.project_id):
        raise DownloadError("Invalid project identifier for source storage.")
    temporary: Path | None = None
    try:
        folder = settings.data_dir / "raw" / project.project_id
        folder.mkdir(parents=True, exist_ok=True)
        digest = hashlib.sha256()
        size = 0
        with NamedTemporaryFile(dir=folder, suffix=".tmp", delete=False) as stream:
            temporary = Path(stream.name)
            with session.get(project.pdf_url, stream=True, timeout=settings.timeout,
                             allow_redirects=False) as response:
                response.raise_for_status()
                if 300 <= response.status_code < 400:
                    raise DownloadError("PDF redirected; review the new source URL before retrying.")
                for chunk in response.iter_content(chunk_size=65536):
                    size += len(chunk)
                    if size > settings.max_pdf_bytes:
                        raise DownloadError("PDF exceeds the configured download size limit.")
                    digest.update(chunk)
                    stream.write(chunk)
            stream.flush()
            os.fsync(stream.fileno())
        with temporary.open("rb") as stream:
            if stream.read(5) != b"%PDF-":
                raise DownloadError("Downloaded content does not have a PDF header.")
        try:
            with pdfplumber.open(temporary) as pdf:
                if not pdf.pages:
                    raise ValueError("no pages")
                for page in pdf.pages:
                    _ = page.width, page.height, page.objects
        except Exception as exc:
            raise DownloadError("Downloaded PDF is corrupt, encrypted or unreadable.") from exc
        sha256 = digest.hexdigest()
        path = folder / f"{sha256}.pdf"
        # Hard-link publication is atomic and refuses to overwrite an existing original.
        try:
            os.link(temporary, path)
        except FileExistsError:
            if hashlib.sha256(path.read_bytes()).hexdigest() != sha256:
                raise DownloadError("Existing archived PDF failed its content-hash check.")
        retrieved_at = datetime.now(timezone.utc)
        manifest = path.with_suffix(".json")
        metadata = {"project_id": project.project_id, "source_url": project.pdf_url,
                    "listing_url": project.listing_url, "sha256": sha256,
                    "retrieved_at": retrieved_at.isoformat(), "bytes": size}
        with NamedTemporaryFile(mode="w", dir=folder, suffix=".tmp", delete=False) as stream:
            manifest_temp = Path(stream.name)
            try:
                json.dump(metadata, stream, indent=2)
                stream.flush()
                os.fsync(stream.fileno())
                try:
                    os.link(manifest_temp, manifest)
                except FileExistsError:
                    previous = json.loads(manifest.read_text())
                    if (previous["sha256"] != sha256 or previous["project_id"] != project.project_id
                            or previous["source_url"] != project.pdf_url):
                        raise DownloadError("Archived PDF provenance conflicts with this download.")
                    retrieved_at = datetime.fromisoformat(previous["retrieved_at"])
            finally:
                manifest_temp.unlink(missing_ok=True)
        return replace(project, path=path, sha256=sha256, retrieved_at=retrieved_at)
    except DownloadError:
        raise
    except (RequestException, OSError, ValueError, KeyError) as exc:
        raise DownloadError(f"Could not preserve PDF for {project.project_id}; check network and folder permissions.") from exc
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)
