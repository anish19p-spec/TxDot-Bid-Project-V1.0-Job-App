# TxDOT Aviation Bid Analysis Tool

A Python tool for extracting TxDOT aviation bids, comparing contractors, and tracking historical unit prices.

## Completed

- Discover project listings, download PDFs, and extract bids with Gemini.
- Validate bid data and generate Excel/PDF comparisons with bidder rankings.
- Maintain a duplicate-free master workbook of overall low-bid unit prices.
- Explore historical estimates in Excel by date, quantity, and TxDOT district.

## To do

- Create a linear regression model based on previous data consisting of data like geographical preference, units, and the date
- Create a chatbot for exploring bid comparisons and historical prices.

## Install from source

Requires Python 3.11+. Run from the project folder on macOS/Linux:

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -e .
python -m txdot_bid_analysis --help
```

## Guide

- [Setup, commands, outputs, and limitations](docs/USER_GUIDE.md)
- https://docs.google.com/document/d/1tyh6H2ZtWiufpxzmb1T6wt_iHFg5n_ynEeaA2n71YpQ/edit?usp=sharing
- For Lockhart airport:

source .venv/bin/activate

# Find Lockhart, download its PDF, and extract the bids with Gemini.
python -m txdot_bid_analysis --extract --project NTB2614LHART00034

# Select the latest saved extraction.
AUDIT_FILE=$(python -c 'from pathlib import Path; print(max((p for p in Path("data/raw/NTB2614LHART00034/extractions").glob("*.json") if not p.name.endswith(".structured.json")), key=lambda p: p.stat().st_mtime))')

# Create the Excel and PDF comparisons.
python -m txdot_bid_analysis --report-saved "$AUDIT_FILE"
