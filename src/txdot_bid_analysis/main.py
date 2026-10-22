"""CLI for project-link discovery; document processing remains planned."""
import argparse
from collections.abc import Sequence
import webbrowser
from pathlib import Path

from requests import Session
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

from . import __version__
from .config import DEFAULT_PROJECT_ID, Settings
from .errors import BidAnalysisError
from .downloader import download_pdf
from .pdf_parser import parse_pdf, verify_saved_extraction
from .analyzer import analyze_bids
from .report_writer import write_reports
from .master_sheet import update_master
from .validator import validate_project
from .scraper import discover_projects, fetch_listing, select_project


def create_session(settings: Settings) -> Session:
    """Bound retries to idempotent TxDOT GETs; never retry paid Gemini requests here."""
    session = Session()
    session.headers.update({"User-Agent": "TxDOTBidAnalysis/0.1"})
    session.mount("https://", HTTPAdapter(max_retries=Retry(
        total=settings.max_retries, backoff_factor=0.5, backoff_max=10,
        status_forcelist=(429, 500, 502, 503, 504), allowed_methods=frozenset({"GET"}),
        respect_retry_after_header=False,
    )))
    return session


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Discover TxDOT aviation project PDF links")
    parser.add_argument("--version", action="version", version=__version__)
    # Change --project to another listing ID in the future; discovery is already reusable.
    parser.add_argument("--project", default=DEFAULT_PROJECT_ID, help="exact TxDOT project ID")
    actions = parser.add_mutually_exclusive_group()
    actions.add_argument("--find-project", action="store_true", help="read listing and print the PDF link")
    actions.add_argument("--open-project", action="store_true", help="find the project and open its PDF in your browser")
    actions.add_argument("--download", action="store_true", help="download and validate the selected PDF")
    actions.add_argument("--extract", action="store_true", help="download and send PDF to Gemini; API charges may apply")
    actions.add_argument("--verify-saved", type=Path, metavar="AUDIT_JSON",
                         help="check a saved extraction and PDF locally; no API calls")
    actions.add_argument("--report-saved", type=Path, metavar="AUDIT_JSON",
                         help="rank saved data and generate Excel/PDF reports without API calls")
    parser.add_argument("--output-dir", type=Path, help="report output folder (default output/<project-id>)")
    actions.add_argument("--master-saved", type=Path, metavar="AUDIT_JSON",
                         help="add overall low-bid unit prices to the master workbook; offline")
    parser.add_argument("--master-path", type=Path, default=Path("output/master_unit_prices.xlsx"),
                        help="historical master workbook path")
    actions.add_argument("--estimate", action="store_true", help="create interactive Excel estimation workbook")
    parser.add_argument("--estimate-output", type=Path, default=Path("output/estimates/estimate.xlsx"))
    parser.add_argument("--regions-file", type=Path, default=Path("references/airport_regions.csv"),
                        help="verified airport-to-TxDOT-district mapping for estimates")
    args = parser.parse_args(argv)
    if args.estimate:
        from .master_sheet import read_historical_prices
        from .report_writer import write_estimate_workbook
        from .geography import enrich_regions
        if args.estimate_output.resolve() == args.master_path.resolve():
            parser.error("Estimate output must be separate from the master workbook.")
        try:
            prices = enrich_regions(read_historical_prices(args.master_path), args.regions_file)
            path = write_estimate_workbook(prices, args.estimate_output)
            print(f"Interactive estimate: {path}. Edit the yellow cells in Excel. No API calls; master unchanged.")
            return 0
        except (BidAnalysisError, ValueError, OSError) as exc:
            print(f"Estimation failed: {exc}")
            return 1
    if args.master_saved is not None:
        try:
            project = verify_saved_extraction(args.master_saved)
            analysis = analyze_bids(project, validate_project(project))
            result = update_master(project, analysis, args.master_path)
            print(f"Master {result.status}: {result.path}; {result.records_added} records added. "
                  f"Bid date: {project.bid_date.isoformat()}. No API calls.")
            return 0
        except BidAnalysisError as exc:
            print(f"Master update failed: {exc}")
            return 1
    if args.report_saved is not None:
        try:
            result = verify_saved_extraction(args.report_saved)
            analysis = analyze_bids(result, validate_project(result))
            destination = args.output_dir or Settings().output_dir / result.project_id
            paths = write_reports(result, analysis, destination)
            names = {bidder.bidder_id: bidder.company_name for bidder in result.bidders}
            print("Lowest submitted bidder: " + "; ".join(names[key] for key in analysis.lowest_submitted_bidder_ids))
            print(f"Excel: {paths.excel}\nPDF: {paths.pdf}\nNo API calls. Historical master not updated.")
            return 0
        except BidAnalysisError as exc:
            print(f"Report generation failed: {exc}")
            return 1
    if args.verify_saved is not None:
        try:
            result = verify_saved_extraction(args.verify_saved)
            print(f"Local verification passed: {len(result.bidders)} bidders, {len(result.items)} items, "
                  f"{len(result.records)} records. No API calls. Source transcription requires separate review.")
            return 0
        except BidAnalysisError as exc:
            print(f"Local verification failed: {exc}")
            return 1
    if not (args.find_project or args.open_project or args.download or args.extract):
        print("Choose --find-project, --open-project, --download, --extract, --verify-saved, --report-saved or --master-saved.")
        return 0
    settings = Settings()
    try:
        with create_session(settings) as session:
            session.headers.update({"User-Agent": "TxDOTBidAnalysis/0.1"})
            html = fetch_listing(session, settings.listing_url, timeout=settings.timeout)
        project = select_project(discover_projects(html, settings.listing_url), args.project)
        print(f"{project.airport_name} | {project.project_id} | {project.bid_date.isoformat()}")
        print(project.pdf_url)
        if args.download or args.extract:
            with create_session(settings) as session:
                document = download_pdf(session, project, settings)
            print(f"Source PDF: {document.path}")
            if args.extract:
                result = parse_pdf(document, allow_ocr=True)
                print(f"Arithmetic checks passed: {len(result.bidders)} bidders, {len(result.items)} items. "
                      "Source-cell review is still required; no reports or master updates were generated.")
        if args.open_project and not webbrowser.open(project.pdf_url):
            print("Could not launch a browser. Open the printed PDF URL manually.")
            return 1
    except (BidAnalysisError, webbrowser.Error) as exc:
        print(f"Project processing failed: {exc}")
        return 1
    return 0
