"""Explicit, source-backed airport-to-TxDOT-district mappings; no name guessing."""
import csv
from dataclasses import dataclass, replace
from datetime import date
from pathlib import Path
from urllib.parse import urlparse

from .models import HistoricalPrice


@dataclass(frozen=True)
class AirportRegion:
    county: str
    region: str
    airport_source: str
    region_source: str
    verified_on: date


def read_regions(path: Path) -> dict[str, AirportRegion]:
    expected = ['airport_name', 'county', 'region', 'airport_source', 'region_source', 'verified_on']
    result = {}
    with path.open(newline='', encoding='utf-8') as handle:
        reader = csv.DictReader(handle)
        if reader.fieldnames != expected:
            raise ValueError('Airport region file has unsupported headers.')
        for row in reader:
            if None in row or any(not value or not value.strip() for value in row.values()):
                raise ValueError('Airport region rows must contain all fields and source URLs.')
            key = ' '.join(row['airport_name'].split()).casefold()
            if key in result:
                raise ValueError('Duplicate airport in region mapping.')
            region = row['region'].strip()
            if region.casefold() in ('all regions', 'unknown'):
                raise ValueError('Region mapping uses a reserved selector name.')
            for field in ('airport_source', 'region_source'):
                url = urlparse(row[field])
                if url.scheme != 'https' or not url.netloc:
                    raise ValueError('Region evidence requires HTTPS source URLs.')
            result[key] = AirportRegion(row['county'].strip(), region, row['airport_source'],
                                       row['region_source'], date.fromisoformat(row['verified_on']))
    return result


def enrich_regions(prices: tuple[HistoricalPrice, ...], path: Path) -> tuple[HistoricalPrice, ...]:
    mappings = read_regions(path)
    enriched = []
    for price in prices:
        mapping = mappings.get(' '.join(price.airport_name.split()).casefold())
        if mapping is None:
            enriched.append(replace(price, region='Unknown', county='', geography_sources=''))
        else:
            evidence = f'{mapping.airport_source}\n{mapping.region_source}\nVerified: {mapping.verified_on}'
            enriched.append(replace(price, region=mapping.region, county=mapping.county,
                                    geography_sources=evidence))
    return tuple(enriched)
