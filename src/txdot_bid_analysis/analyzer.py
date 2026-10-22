"""Deterministic total and item rankings from validated submitted values."""
from decimal import Decimal

from .errors import ValidationError
from .models import (
    AnalysisResult, BidRanking, ItemBidRanking, ItemComparison, Project, ValidationResult,
)
from .validator import validate_project


def analyze_bids(project: Project, validation: ValidationResult) -> AnalysisResult:
    """Revalidate current data and rank exact submitted totals/unit prices, retaining ties."""
    current = validate_project(project)
    issues = tuple(dict.fromkeys((*validation.issues, *current.issues)))
    if any(issue.severity == "error" for issue in issues):
        raise ValidationError("Bid ranking requires complete, reconciled project data.")
    bidders = sorted(project.bidders, key=lambda b: (b.submitted_total, b.bidder_id))
    low = bidders[0].submitted_total
    totals = [bidder.submitted_total for bidder in bidders]
    rankings = tuple(BidRanking(b.bidder_id, 1 + sum(total < b.submitted_total for total in totals),
                               b.submitted_total, b.submitted_total - low,
                               (b.submitted_total - low) / low * Decimal(100)) for b in bidders)
    records = {(r.bidder_id, r.item_id): r for r in project.records}
    comparisons = []
    for item in project.items:
        prices = {b.bidder_id: records[b.bidder_id, item.item_id].unit_price for b in bidders}
        minimum, maximum = min(prices.values()), max(prices.values())
        lowest = tuple(key for key in prices if prices[key] == minimum)
        highest = tuple(key for key in prices if prices[key] == maximum)
        middle = tuple(key for key in prices if minimum < prices[key] < maximum)
        item_rankings = []
        for bidder_id, price in sorted(prices.items(), key=lambda pair: (pair[1], pair[0])):
            if minimum == maximum:
                position = "Only bid" if len(prices) == 1 else "Tied lowest / highest"
            elif price == minimum:
                position = "Tied lowest" if len(lowest) > 1 else "Lowest"
            elif price == maximum:
                position = "Tied highest" if len(highest) > 1 else "Highest"
            else:
                position = "Middle" if len(prices) == 3 else "Intermediate"
            item_rankings.append(ItemBidRanking(bidder_id,
                1 + sum(other < price for other in prices.values()), price, position))
        comparisons.append(ItemComparison(item.item_id, lowest, middle, highest, tuple(item_rankings)))
    return AnalysisResult(project.project_id, project.document.sha256,
        tuple(b.bidder_id for b in bidders if b.submitted_total == low), rankings,
        tuple(issue for issue in issues if issue.severity == "warning"), tuple(comparisons))
