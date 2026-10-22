"""Explicit defaults; importing this module performs no I/O."""
from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path

LISTING_URL = "https://www.dot.state.tx.us/AVN/avninfo/status/construct/index.htm"
DEFAULT_PROJECT_ID = "NTB2614LHART00034"
DEFAULT_GEMINI_MODEL = "gemini-3.1-flash-lite"

@dataclass(frozen=True)
class Settings:
    listing_url: str = LISTING_URL
    project_id: str = DEFAULT_PROJECT_ID
    data_dir: Path = Path("data")
    output_dir: Path = Path("output")
    timeout: tuple[float, float] = (5.0, 30.0)
    max_retries: int = 3
    max_pdf_bytes: int = 25 * 1024 * 1024
    reconciliation_tolerance: Decimal = Decimal("0.01")
