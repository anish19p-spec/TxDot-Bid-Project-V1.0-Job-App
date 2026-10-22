"""Read TxDOT project listings and resolve the selected project's PDF link."""
from datetime import datetime
from urllib.parse import urljoin, urlsplit

from bs4 import BeautifulSoup
from requests import RequestException, Session

from .errors import DiscoveryError, RetrievalError
from .models import ProjectSource


def fetch_listing(session: Session, url: str, *, timeout: tuple[float, float]) -> str:
    """Fetch HTML without redirects so relative links keep their original base URL."""
    try:
        with session.get(url, timeout=timeout, allow_redirects=False) as response:
            response.raise_for_status()
            if 300 <= response.status_code < 400:
                raise RetrievalError(f"Listing redirected: {url}; review its new URL before retrying.")
            if "html" not in response.headers.get("Content-Type", "").lower():
                raise RetrievalError(f"Expected HTML from listing: {url}")
            # Decode bytes ourselves: the source contains conflicting charset declarations.
            encoding = response.apparent_encoding or "utf-8"
            return response.content.decode(encoding, errors="strict")
    except (RequestException, UnicodeError, LookupError) as exc:
        raise RetrievalError(f"Could not read TxDOT listing {url}: {exc}") from exc


def discover_projects(html: str, listing_url: str) -> tuple[ProjectSource, ...]:
    """Read project rows and resolve PDF links; no project-specific parsing rules."""
    soup = BeautifulSoup(html, "html.parser")
    projects: list[ProjectSource] = []
    for row in soup.find_all("tr"):
        cells = row.find_all("td", recursive=False)
        # Nested layout tables also have rows; only the three-column project rows qualify.
        if len(cells) != 3:
            continue
        project_id = cells[1].get_text(" ", strip=True).upper()
        if not project_id.startswith("NTB"):
            continue
        airport = cells[0].get_text(" ", strip=True)
        try:
            bid_date = datetime.strptime(
                cells[2].get_text(" ", strip=True), "%m/%d/%Y %I:%M:%S %p"
            ).date()
        except ValueError as exc:
            raise DiscoveryError(f"Invalid opening date for project {project_id}.") from exc
        if not airport:
            raise DiscoveryError(f"Missing airport name for project {project_id}.")
        links = {
            urljoin(listing_url, anchor["href"])
            for anchor in cells[0].find_all("a", href=True)
            if urlsplit(anchor["href"]).path.lower().endswith(".pdf")
        }
        if len(links) > 1:
            raise DiscoveryError(f"Multiple PDF links for project {project_id}.")
        pdf_url = next(iter(links), None)
        if pdf_url:
            source, target = urlsplit(listing_url), urlsplit(pdf_url)
            if (target.scheme != "https" or target.netloc != source.netloc
                    or target.username or target.password):
                raise DiscoveryError(f"Unexpected PDF URL for project {project_id}: {pdf_url}")
        projects.append(ProjectSource(project_id, airport, bid_date, listing_url, pdf_url))
    if not projects:
        raise DiscoveryError("No TxDOT project rows found; the listing layout may have changed.")
    return tuple(projects)


def select_project(projects: tuple[ProjectSource, ...], project_id: str) -> ProjectSource:
    """Select an exact project ID; callers can choose other projects without parser changes."""
    matches = [project for project in projects if project.project_id == project_id.strip().upper()]
    if len(matches) != 1:
        raise DiscoveryError(f"Expected one listing for {project_id}; found {len(matches)}.")
    if matches[0].pdf_url is None:
        raise DiscoveryError(f"No PDF link available for project {project_id}.")
    return matches[0]
