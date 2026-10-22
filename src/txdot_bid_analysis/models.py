"""Immutable domain records shared across pipeline stages."""
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import Literal

@dataclass(frozen=True)
class ProjectSource:
    project_id: str
    airport_name: str
    bid_date: date
    listing_url: str
    pdf_url: str | None
    path: Path | None = None
    sha256: str | None = None
    retrieved_at: datetime | None = None

    @property
    def source_url(self) -> str | None:
        return self.pdf_url

@dataclass(frozen=True)
class SourceRef:
    page: int
    method: Literal["text", "ocr", "gemini"]
    raw_text: str
    confidence: Decimal | None = None

@dataclass(frozen=True)
class Bidder:
    project_id: str
    bidder_id: str
    company_name: str
    submitted_total: Decimal
    opening_order: int | None
    source: SourceRef

@dataclass(frozen=True)
class BidItem:
    project_id: str
    item_id: str
    item_number: str
    quantity: Decimal
    unit: str
    description: str
    source: SourceRef

@dataclass(frozen=True)
class BidRecord:
    project_id: str
    bidder_id: str
    item_id: str
    unit_price: Decimal
    extended_amount: Decimal
    source: SourceRef

@dataclass(frozen=True)
class Project:
    project_id: str
    airport_name: str
    bid_date: date
    source_url: str
    document: ProjectSource
    bidders: tuple[Bidder, ...]
    items: tuple[BidItem, ...]
    records: tuple[BidRecord, ...]
    listing_url: str = ""
    raw_project_number: str | None = None

@dataclass(frozen=True)
class ValidationIssue:
    severity: Literal["error", "warning"]
    code: str
    message: str
    reference: str | None = None

@dataclass(frozen=True)
class ValidationResult:
    issues: tuple[ValidationIssue, ...]

@dataclass(frozen=True)
class BidRanking:
    bidder_id: str
    rank: int
    submitted_total: Decimal
    difference: Decimal
    percentage_above_low: Decimal | None

@dataclass(frozen=True)
class ItemBidRanking:
    bidder_id: str
    rank: int
    unit_price: Decimal
    position: str

@dataclass(frozen=True)
class ItemComparison:
    item_id: str
    lowest_bidder_ids: tuple[str, ...]
    middle_bidder_ids: tuple[str, ...]
    highest_bidder_ids: tuple[str, ...]
    rankings: tuple[ItemBidRanking, ...]

@dataclass(frozen=True)
class AnalysisResult:
    project_id: str
    source_sha256: str
    lowest_submitted_bidder_ids: tuple[str, ...]
    rankings: tuple[BidRanking, ...]
    warnings: tuple[ValidationIssue, ...]
    item_comparisons: tuple[ItemComparison, ...] = ()

@dataclass(frozen=True)
class ReportPaths:
    excel: Path
    pdf: Path

@dataclass(frozen=True)
class MasterUpdateResult:
    path: Path
    status: Literal["inserted", "unchanged", "upgraded"]
    records_added: int


@dataclass(frozen=True)
class HistoricalPrice:
    project_id: str
    bidder_id: str
    bid_date: date
    item_number: str
    description: str
    unit: str
    quantity: Decimal
    unit_price: Decimal
    source_url: str
    source_hash: str
    source_page: int
    airport_name: str = ""
    region: str = "Unknown"
    county: str = ""
    geography_sources: str = ""
