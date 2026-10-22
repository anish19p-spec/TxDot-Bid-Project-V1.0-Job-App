# Usage guide

## Setup

Open a terminal in the project folder. Install Python 3.11+ and run the README installation commands. On Windows, create the environment with `py -3 -m venv .venv` and activate it with `.venv\Scripts\Activate.ps1` in PowerShell.

Copy `.env.example` to `.env` and enter `GEMINI_API_KEY` locally. Never commit the key. `GEMINI_MODEL` selects a model available to your Google account; the configured default is `gemini-3.1-flash-lite`. Shell environment variables override `.env`.

## Discover and extract

```bash
python -m txdot_bid_analysis --find-project --project NTB2614LHART00034
python -m txdot_bid_analysis --download --project NTB2614LHART00034
python -m txdot_bid_analysis --extract --project NTB2614LHART00034
```

Change the project ID to select another listing. Extraction sends the public PDF to Gemini and may incur API charges. Each extraction makes a new API request.

PDFs are saved in `data/raw/<project-id>/`. Extraction responses and structured snapshots are saved in its `extractions/` subfolder. Review the transcription against the PDF before using the results.

## Generate reports from saved data

Replace `PATH_TO_AUDIT_JSON` below with the raw extraction JSON path. Use the file without `.structured` in its name. Keep its companion structured JSON and original PDF.

```bash
python -m txdot_bid_analysis --verify-saved PATH_TO_AUDIT_JSON
python -m txdot_bid_analysis --report-saved PATH_TO_AUDIT_JSON
python -m txdot_bid_analysis --master-saved PATH_TO_AUDIT_JSON
python -m txdot_bid_analysis --estimate
```

These commands use saved data without API calls.

| Output | Location |
| --- | --- |
| Side-by-side Excel and PDF comparisons | `output/<project-id>/bid_comparison.xlsx` and `.pdf` |
| Historical low-bid workbook | `output/master_unit_prices.xlsx` |
| Interactive historical estimate | `output/estimates/estimate.xlsx` |

The master stores prices from the overall lowest submitted bidder, including ties. Repeat imports do not add duplicate history. Changed source documents for an existing project produce a conflict.

## Interactive estimates

Open the estimate workbook in Excel and edit the yellow cells: item, planned quantity, date, rolling month window, quantity tolerance, and TxDOT district. Excel calculates a median historical reference from matching projects. These values are not verified current market quotes or predictions. Region mappings live in `references/airport_regions.csv`; only Lockhart is currently mapped, and other airports remain Unknown.

## Optional Gmail sending

Set `EMAIL_FROM`, `EMAIL_TO`, and `GMAIL_APP_PASSWORD` in `.env`. Use a Gmail app password. Generate all four report files first, then explicitly send:

```bash
python -c 'from txdot_bid_analysis.email_sender import send_email; send_email()'
```

The default is Lockhart. Pass `project_id="YOUR_PROJECT_ID"` to send another project's comparison. There is no automatic sending or scheduler.

## Design

Scraping, downloading, extraction, validation, ranking, reporting, and workbook persistence are separate modules in `src/txdot_bid_analysis/`. Immutable records carry source metadata. Monetary values use `Decimal`. Downloaded PDFs are preserved by SHA-256 hash; saved extraction audits support offline replay.

## Limitations and recovery

- Base bids are the supported scope; new formats need independent review.
- Automated schema and arithmetic checks do not prove source-cell accuracy.
- Lowest submitted bidder does not mean awarded contractor.
- Historical estimates need comparable projects; regression is not implemented.
- Gemini failures may require checking credentials, model access, billing, or quota. Never patch extracted amounts to force validation to pass.
- Run from the repository root so configuration and region paths resolve.
- Source PDFs, test fixtures, credentials, and generated reports are intentionally not distributed. Generate your own artifacts using the commands above.
