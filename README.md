# TxDOT Aviation Bid Analysis Tool

A Python tool for extracting TxDOT aviation bids, comparing contractors, and tracking historical unit prices.



## Completed

- Discover project listings, download PDFs, and extract bids with Gemini.
- Validate bid data and generate Excel/PDF comparisons with bidder rankings.
- Maintain a duplicate-free master workbook of overall low-bid unit prices.
- Explore historical estimates in Excel by date, quantity, and TxDOT district.

## To do

- Create a linear regression model based on previous data consisting of data like geographical preference, units, and the date
- Create a chatbot so people can conveniently

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
