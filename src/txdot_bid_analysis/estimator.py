"""Prepare one historical unit price per project/item for an interactive estimate."""
from collections import defaultdict
from dataclasses import dataclass
from datetime import date
from decimal import Decimal
from statistics import median

from .models import HistoricalPrice


@dataclass(frozen=True)
class EstimateObservation:
    item_key: str
    project_id: str
    bid_date: date
    quantity: Decimal
    unit_price: Decimal
    sources: tuple[HistoricalPrice, ...]


def prepare_observations(prices: tuple[HistoricalPrice, ...]) -> tuple[EstimateObservation, ...]:
    groups = defaultdict(list)
    seen = set()
    for row in prices:
        key = (row.item_number.strip().upper(), ' '.join(row.description.split()).upper(), row.unit.strip().upper())
        identity = (key, row.project_id, row.bidder_id)
        if identity in seen:
            raise ValueError('Duplicate historical price record.')
        seen.add(identity)
        if any(not isinstance(v, Decimal) or not v.is_finite() or v < 0
               for v in (row.quantity, row.unit_price)):
            raise ValueError('Historical quantities and prices must be finite nonnegative Decimals.')
        groups[(key, row.project_id)].append(row)
    observations = []
    for (key, project), rows in sorted(groups.items()):
        if len({(r.bid_date, r.quantity, r.source_hash, r.source_url, r.region, r.county, r.geography_sources) for r in rows}) != 1:
            raise ValueError('Conflicting dates, quantities or sources for one project/item.')
        first = rows[0]
        # Tied low bidders share one project observation, avoiding extra project weight.
        observations.append(EstimateObservation(' | '.join(key), project, first.bid_date,
                            first.quantity, median(r.unit_price for r in rows), tuple(rows)))
    if not observations:
        raise ValueError('No historical records are available.')
    return tuple(observations)
