"""Stage-specific error contract; CLI handling is planned for M2-M7."""
class BidAnalysisError(Exception):
    """Base for actionable application failures."""

class RetrievalError(BidAnalysisError):
    """Listing retrieval failed."""

class DiscoveryError(BidAnalysisError):
    """Project selection or PDF link is missing or ambiguous."""

class DownloadError(BidAnalysisError):
    """Source download or PDF validation failed."""

class ParseError(BidAnalysisError):
    """Document layout or required extraction is unsupported."""

class ValidationError(BidAnalysisError):
    """Critical inconsistencies prevent analysis or persistence."""

class ReportError(BidAnalysisError):
    """A report could not be safely written."""

class MasterUpdateError(BidAnalysisError):
    """Master is locked, invalid, or conflicts with existing history."""
