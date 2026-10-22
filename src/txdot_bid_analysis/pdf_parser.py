"""Convert Gemini's untrusted transcriptions to validated domain records."""
from dataclasses import asdict
from datetime import date, datetime, timezone
from decimal import Decimal
import hashlib
import json
import os
from pathlib import Path
from uuid import NAMESPACE_URL, uuid5

from dotenv import load_dotenv
import pdfplumber
from pydantic import ValidationError as SchemaError

from .config import DEFAULT_GEMINI_MODEL
from .errors import ParseError
from .gemini_extractor import ExtractedDocument, request_extraction
from .models import Bidder, BidItem, BidRecord, Project, ProjectSource, SourceRef
from .validator import normalize_decimal, validate_project


def _require_download(document: ProjectSource) -> None:
    if not document.pdf_url or document.path is None or not document.sha256 or document.retrieved_at is None:
        raise ParseError("Download the project PDF before extraction; file provenance is incomplete.")


def _snapshot(project: Project) -> dict[str, object]:
    """Keep the saved audit format compatible with existing reviewed extractions."""
    result = asdict(project)
    document = project.document
    result["document"] = {"project_id": document.project_id, "source_url": document.pdf_url,
                          "path": document.path, "sha256": document.sha256,
                          "retrieved_at": document.retrieved_at}
    return result


def project_from_extraction(raw_json: str, document: ProjectSource, *, page_count: int) -> Project:
    """Reject schema errors, uncertain cells, disagreement and invalid monetary totals."""
    _require_download(document)
    try:
        extracted = ExtractedDocument.model_validate_json(raw_json)
        if extracted.uncertainties:
            raise ParseError("Gemini reported uncertain or unsupported source data; review the audit response.")
        if extracted.airport_name and extracted.airport_name.casefold().split() != document.airport_name.casefold().split():
            raise ParseError("PDF airport name disagrees with the listing; review source metadata.")
        if extracted.bid_date and date.fromisoformat(extracted.bid_date) != document.bid_date:
            raise ParseError("PDF bid date disagrees with the listing.")
        bidders: list[Bidder] = []
        items: dict[str, BidItem] = {}
        records: list[BidRecord] = []
        def source(page: int, raw_text: str) -> SourceRef:
            if not 1 <= page <= page_count or not raw_text.strip():
                raise ParseError("Missing or out-of-range source page/evidence.")
            return SourceRef(page, "gemini", raw_text)
        def identity(value: str) -> str:
            return str(uuid5(NAMESPACE_URL, f"txdot-bid-analysis/{document.project_id}/{value}"))
        for ordinal, bidder in enumerate(extracted.bidders, 1):
            if not bidder.company_name or not bidder.company_name.strip():
                raise ParseError("Missing bidder name.")
            bidder_id = identity(f"bidder/{ordinal}")
            bidders.append(Bidder(document.project_id, bidder_id, bidder.company_name.strip(),
                normalize_decimal(bidder.submitted_total), None,
                source(bidder.total_page, bidder.total_raw_text)))
            for row in bidder.items:
                if not row.item_number.strip() or not row.unit or not row.description:
                    raise ParseError("Missing required bid-item text.")
                number = row.item_number.strip().upper()
                item_id = identity(f"base/{number}")
                evidence = source(row.page, row.raw_text)
                item = BidItem(document.project_id, item_id, number, normalize_decimal(row.quantity),
                               row.unit.strip().upper(), row.description.strip(), evidence)
                if item_id in items:
                    previous = items[item_id]
                    if (previous.quantity, previous.unit, previous.description.casefold().split()) != (
                            item.quantity, item.unit, item.description.casefold().split()):
                        raise ParseError("Bidders disagree on item quantity, unit or description.")
                else:
                    items[item_id] = item
                records.append(BidRecord(document.project_id, bidder_id, item_id,
                    normalize_decimal(row.unit_price), normalize_decimal(row.extended_amount), evidence))
        project = Project(document.project_id, document.airport_name, document.bid_date,
                          document.source_url, document, tuple(bidders), tuple(items.values()), tuple(records),
                          document.listing_url, extracted.raw_project_number)
        validation = validate_project(project)
        if validation.issues:
            codes = ", ".join(sorted({issue.code for issue in validation.issues}))
            raise ParseError(f"Extracted bid data failed validation ({codes}); review audit response and PDF.")
        return project
    except ParseError:
        raise
    except (SchemaError, ValueError, TypeError, ArithmeticError):
        raise ParseError("Gemini response contains missing, malformed or unsupported values; review the audit response.") from None


def parse_pdf(document: ProjectSource, *, allow_ocr: bool = False) -> Project:
    """Explicitly opt into sending the PDF to Gemini; preserve response before validation."""
    if not allow_ocr:
        raise ParseError("Gemini extraction requires explicit --extract consent to send this public PDF to Google.")
    _require_download(document)
    load_dotenv(Path.cwd() / ".env", override=False)
    key = os.environ.get("GEMINI_API_KEY", "").strip()
    model = os.environ.get("GEMINI_MODEL", "").strip() or DEFAULT_GEMINI_MODEL
    if not key:
        raise ParseError("Set GEMINI_API_KEY (optionally GEMINI_MODEL) in the project .env or environment; see .env.example.")
    try:
        content = document.path.read_bytes()
        if len(content) > 15 * 1024 * 1024:
            raise ParseError("PDF exceeds the current 15 MiB inline Gemini limit; file-upload support is not implemented.")
        if hashlib.sha256(content).hexdigest() != document.sha256:
            raise ParseError("PDF identity/hash does not match its listing and provenance.")
        with pdfplumber.open(document.path) as pdf:
            page_count = len(pdf.pages)
        raw = request_extraction(content, api_key=key, model=model)
        # Each paid attempt gets an audit record; invalid model output is preserved too.
        audit_dir = document.path.parent / "extractions"
        audit_dir.mkdir(exist_ok=True)
        timestamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S%fZ")
        audit = audit_dir / f"{document.sha256}-{timestamp}.json"
        with audit.open("x") as stream:
            json.dump({"source_sha256": document.sha256, "source_url": document.source_url,
                       "listing_url": document.listing_url, "model": model, "prompt_version": "1",
                       "retrieved_at": timestamp, "raw_response": raw,
                       "status": "unreviewed_model_transcription"}, stream, indent=2)
        project = project_from_extraction(raw, document, page_count=page_count)
        save_structured_extraction(project, audit.with_suffix(".structured.json"))
        return project
    except ParseError:
        raise
    except OSError:
        raise ParseError("Could not read the source or preserve extraction audit; check folder permissions.") from None


def save_structured_extraction(project: Project, path: Path) -> None:
    """Save a unique per-attempt snapshot, preserving exact Decimal strings for review."""
    def encode(value: object) -> str:
        if isinstance(value, (date, datetime)):
            return value.isoformat()
        if isinstance(value, (Decimal, Path)):
            return str(value)
        raise TypeError("Unsupported structured extraction value.")
    with path.open("x") as stream:
        json.dump({"status": "arithmetic_checked_source_review_required",
                   "metadata_sources": {"airport_name": project.listing_url, "bid_date": project.listing_url},
                   "project": _snapshot(project)}, stream, default=encode, indent=2)


def verify_saved_extraction(audit_path: Path) -> Project:
    """Replay saved JSON locally and check its snapshot/hash; never call Gemini or HTTP."""
    try:
        audit = json.loads(audit_path.read_text())
        snapshot = json.loads(audit_path.with_suffix(".structured.json").read_text())
        stored = snapshot["project"]
        doc = stored["document"]
        sha256 = audit["source_sha256"]
        if (not isinstance(sha256, str) or len(sha256) != 64
                or any(c not in "0123456789abcdef" for c in sha256)):
            raise ParseError("Saved response has an invalid source hash.")
        path = audit_path.parent.parent / f"{sha256}.pdf"
        if hashlib.sha256(path.read_bytes()).hexdigest() != sha256:
            raise ParseError("Saved PDF failed its content-hash check.")
        if (doc["sha256"] != sha256 or doc["source_url"] != audit["source_url"]
                or stored["source_url"] != audit["source_url"]
                or stored["listing_url"] != audit["listing_url"]):
            raise ParseError("Saved response and snapshot provenance disagree.")
        if doc["project_id"] != stored["project_id"]:
            raise ParseError("Saved document and project identifiers disagree.")
        document = ProjectSource(stored["project_id"], stored["airport_name"],
            date.fromisoformat(stored["bid_date"]), stored["listing_url"], stored["source_url"],
            path, sha256, datetime.fromisoformat(doc["retrieved_at"]))
        with pdfplumber.open(path) as pdf:
            project = project_from_extraction(audit["raw_response"], document,
                                              page_count=len(pdf.pages))
        rebuilt = json.loads(json.dumps(_snapshot(project), default=str))
        # Paths may be relative to the original run; compare their resolved destinations.
        if Path(stored["document"]["path"]).resolve() != path.resolve():
            raise ParseError("Saved snapshot refers to a different source PDF path.")
        rebuilt["document"]["path"] = stored["document"]["path"]
        rebuilt["document"]["retrieved_at"] = document.retrieved_at.isoformat()
        if rebuilt != stored:
            raise ParseError("Saved structured values differ from the original response; no changes accepted.")
        return project
    except ParseError:
        raise
    except (OSError, ValueError, KeyError, TypeError):
        raise ParseError("Cannot verify saved extraction: require its raw audit, structured snapshot and original PDF.") from None
