"""Numeric normalization and explicit bid completeness/arithmetic checks."""
from decimal import Decimal, ROUND_HALF_UP
import re

from .models import Project, ValidationIssue, ValidationResult


def normalize_decimal(raw: str) -> Decimal:
    """Accept explicit nonnegative US currency/quantity strings; never fill blanks."""
    if not isinstance(raw, str):
        raise ValueError("Numeric source values must be strings.")
    value = raw.strip().removeprefix("$").strip()
    if not re.fullmatch(r"(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d+)?", value):
        raise ValueError("Invalid or ambiguous numeric source value.")
    return Decimal(value.replace(",", ""))


def validate_project(project: Project, *, tolerance: Decimal = Decimal("0.01")) -> ValidationResult:
    """Check foreign keys, complete bidder/item coverage and submitted arithmetic."""
    if not tolerance.is_finite() or tolerance < 0:
        raise ValueError("Tolerance must be finite and nonnegative.")
    issues: list[ValidationIssue] = []
    def error(code: str, message: str, reference: str | None = None) -> None:
        issues.append(ValidationIssue("error", code, message, reference))
    bidders = {b.bidder_id: b for b in project.bidders}
    items = {i.item_id: i for i in project.items}
    if not bidders or not items or len(bidders) != len(project.bidders) or len(items) != len(project.items):
        error("IDENTITY", "Missing or duplicate bidders/items.")
    if project.document.project_id != project.project_id:
        error("PROJECT", "Document belongs to a different project.")
    seen: set[tuple[str, str]] = set()
    submitted = {key: Decimal(0) for key in bidders}
    calculated = dict(submitted)
    for obj in (*project.bidders, *project.items, *project.records):
        if obj.project_id != project.project_id:
            error("PROJECT", "Record belongs to a different project.")
        if obj.source.page < 1 or not obj.source.raw_text.strip():
            error("SOURCE", "Missing page or source evidence.")
    for item in project.items:
        if not item.quantity.is_finite() or item.quantity < 0 or not all((item.item_number.strip(), item.unit.strip(), item.description.strip())):
            error("ITEM", "Invalid item metadata/quantity.", item.item_id)
    for bidder in project.bidders:
        if not bidder.company_name.strip() or not bidder.submitted_total.is_finite() or bidder.submitted_total <= 0:
            error("BIDDER", "Invalid bidder name or submitted total.", bidder.bidder_id)
    for record in project.records:
        key = record.bidder_id, record.item_id
        if key in seen:
            error("DUPLICATE", "Duplicate bidder/item record.", str(key))
        seen.add(key)
        if record.bidder_id not in bidders or record.item_id not in items:
            error("REFERENCE", "Unresolved bidder/item reference.", str(key))
            continue
        values = (record.unit_price, record.extended_amount, items[record.item_id].quantity)
        if any(not value.is_finite() or value < 0 for value in values):
            error("NUMBER", "Non-finite or negative amount.", str(key))
            continue
        amount = (values[0] * values[2]).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        submitted[record.bidder_id] += record.extended_amount
        calculated[record.bidder_id] += amount
        if abs(amount - record.extended_amount) > tolerance:
            error("ITEM_TOTAL", "Quantity × unit price differs from submitted extension.", str(key))
    if seen != {(b, i) for b in bidders for i in items}:
        error("COVERAGE", "Bidders do not have exactly the same complete set of items.")
    for key, bidder in bidders.items():
        if bidder.submitted_total.is_finite() and any(abs(total[key] - bidder.submitted_total) > tolerance for total in (submitted, calculated)):
            error("BID_TOTAL", "Item amounts do not reconcile to submitted bid total.", key)
    return ValidationResult(tuple(issues))
