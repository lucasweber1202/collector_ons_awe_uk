"""Selected raw Average Weekly Earnings series from the official ONS EARN01 workbook."""

from __future__ import annotations

import hashlib
import logging
import math
import re
from dataclasses import dataclass
from datetime import UTC, date, datetime
from typing import Any
from urllib.parse import urljoin

import httpx
import xlrd

from scripts.config import (
    MAX_DOWNLOAD_BYTES,
    MAX_STALE_MONTHS,
    MIN_HISTORY_YEARS,
    REQUEST_TIMEOUT,
    USER_AGENT,
)
from scripts.snapshots import Snapshot, build_snapshot
from scripts.time_series import Observation

logger = logging.getLogger(__name__)


# -- series_id contract (GUIDELINES.md 4) ---------------------------------
# series_id is uppercase, underscore-separated and ordered coarse -> fine. The
# pair below is the canonical public surface: parse splits an id into its
# components, build rejoins them, and build(*parse(sid)) == sid for every id
# this collector emits. Only economic identity is encoded -- never a delivery
# provider or any other detail of how the value reached us.


def parse_series_id(series_id: str) -> tuple[str, ...]:
    """Split a series_id into its underscore-delimited components.

    Raises ValueError on anything this collector would not have produced:
    lowercase, empty components, or an id with no structure at all.
    """
    if not series_id or series_id != series_id.upper():
        raise ValueError(f"series_id must be uppercase: {series_id!r}")
    components = tuple(series_id.split("_"))
    if any(not component for component in components):
        raise ValueError(f"series_id has an empty component: {series_id!r}")
    return components


def build_series_id(*components: str) -> str:
    """Rejoin the tuple parse_series_id returned into the original id."""
    if not components:
        raise ValueError("series_id needs at least one component")
    if any(not component or component != component.upper() for component in components):
        raise ValueError(f"invalid series_id components: {components!r}")
    return "_".join(components)


# -- 5.1 usable-series filtering ------------------------------------------


@dataclass(frozen=True)
class UsabilityReport:
    """What the filter removed, for logging and for tests to assert on."""

    kept: tuple[str, ...]
    stale: tuple[str, ...]
    short_history: tuple[str, ...]
    empty: tuple[str, ...]

    @property
    def dropped(self) -> tuple[str, ...]:
        return tuple(sorted(set(self.stale) | set(self.short_history) | set(self.empty)))


def _months_between(earlier: date, later: date) -> int:
    """Whole months from ``earlier`` to ``later``, day-of-month aware."""
    months = (later.year - earlier.year) * 12 + (later.month - earlier.month)
    if later.day < earlier.day:
        months -= 1
    return months


def _is_valid(value: Any) -> bool:
    """A real observation: present, numeric and finite."""
    if value is None:
        return False
    try:
        numeric = float(value)
    except (TypeError, ValueError):
        return False
    return math.isfinite(numeric)


def assess_series(
    reference_dates: list[date],
    today: date,
    max_stale_months: int = MAX_STALE_MONTHS,
    min_history_years: float = MIN_HISTORY_YEARS,
) -> str:
    """Classify one series from the reference dates of its valid observations.

    Returns ``"keep"``, ``"empty"``, ``"stale"`` or ``"short_history"``.
    Recency is judged at the period end and over non-null values only: a source
    that keeps listing a discontinued series with empty recent cells must not
    look live because of those blanks.
    """
    if not reference_dates:
        return "empty"
    first, last = min(reference_dates), max(reference_dates)
    if _months_between(last, today) > max_stale_months:
        return "stale"
    if _months_between(first, last) < round(min_history_years * 12):
        return "short_history"
    return "keep"


def filter_usable_series(
    observations: list[Any],
    catalog: dict[str, dict[str, Any]],
    today: date,
    max_stale_months: int = MAX_STALE_MONTHS,
    min_history_years: float = MIN_HISTORY_YEARS,
) -> tuple[list[Any], dict[str, dict[str, Any]], UsabilityReport]:
    """Drop obsolete and history-less series before anything is persisted.

    Runs after parsing and before the time_series / metadata upsert, so the
    standardized tables never carry a dead or stub series, and prunes the
    catalog alongside the observations so metadata can never describe a series
    the database does not hold (GUIDELINES.md 5.1).
    """
    valid_dates: dict[str, list[date]] = {}
    for observation in observations:
        if _is_valid(observation.value):
            valid_dates.setdefault(observation.series_id, []).append(observation.reference_date)

    verdicts: dict[str, str] = {}
    for series_id in set(catalog) | {o.series_id for o in observations}:
        verdicts[series_id] = assess_series(
            valid_dates.get(series_id, []), today, max_stale_months, min_history_years
        )

    keep = {series_id for series_id, verdict in verdicts.items() if verdict == "keep"}
    report = UsabilityReport(
        kept=tuple(sorted(keep)),
        stale=tuple(sorted(s for s, v in verdicts.items() if v == "stale")),
        short_history=tuple(sorted(s for s, v in verdicts.items() if v == "short_history")),
        empty=tuple(sorted(s for s, v in verdicts.items() if v == "empty")),
    )

    if report.dropped:
        logger.info(
            "Usable-series filter: kept %d, dropped %d "
            "(stale=%d short_history=%d empty=%d; max_stale_months=%d min_history_years=%s)",
            len(report.kept),
            len(report.dropped),
            len(report.stale),
            len(report.short_history),
            len(report.empty),
            max_stale_months,
            min_history_years,
        )
        for series_id in report.stale:
            logger.info(
                "Dropped %s: last valid observation older than %d months",
                series_id,
                max_stale_months,
            )
        for series_id in report.short_history:
            logger.info(
                "Dropped %s: valid history shorter than %s years", series_id, min_history_years
            )
        for series_id in report.empty:
            logger.info("Dropped %s: no valid observations", series_id)
    else:
        logger.info("Usable-series filter: all %d series usable", len(report.kept))

    kept_observations = [o for o in observations if o.series_id in keep]
    kept_catalog = {sid: fields for sid, fields in catalog.items() if sid in keep}
    return kept_observations, kept_catalog, report


LANDING = "https://www.ons.gov.uk/employmentandlabourmarket/peopleinwork/earningsandworkinghours/datasets/averageweeklyearningsearn01"
SHEETS = {"1. AWE Total Pay": "TOTAL", "3. AWE Regular Pay": "REGULAR"}
SELECTED = {
    "KAB9",
    "KAC4",
    "KAC7",
    "K5BZ",
    "K5C4",
    "K5CA",
    "K5CD",
    "K5CG",
    "KAI7",
    "KAJ2",
    "KAJ5",
    "K5DL",
    "K5DO",
    "K5DU",
    "K5DX",
    "K5E2",
}


@dataclass(frozen=True)
class ExtractedData:
    observations: list[Observation]
    snapshots: list[Snapshot]
    catalog: dict[str, dict[str, Any]]
    releases: list[datetime]
    availability_by_key: dict[tuple[str, date], tuple[datetime, str, date | None]]
    min_lag_days: int = 0
    max_lag_days: int = 62
    inferred_lag_days: int = 31


def parse_xls(
    body: bytes, snapshot_id: str, source_url: str
) -> tuple[list[Observation], dict[str, dict[str, Any]], datetime]:
    book = xlrd.open_workbook(file_contents=body)
    observations: list[Observation] = []
    catalog: dict[str, dict[str, Any]] = {}
    keys: set[tuple[str, date]] = set()
    release = None
    for sheet_name, measure in SHEETS.items():
        if sheet_name not in book.sheet_names():
            raise ValueError(f"ONS AWE workbook missing sheet {sheet_name}")
        sheet = book.sheet_by_name(sheet_name)
        match = re.search(r"(\d{1,2} \w+ \d{4})", str(sheet.cell_value(1, 0)))
        if not match:
            raise ValueError("ONS AWE publication date missing")
        candidate = datetime.strptime(match.group(1), "%d %B %Y").replace(hour=7, tzinfo=UTC)
        release = candidate if release is None else release
        if release != candidate:
            raise ValueError("ONS AWE sheets disagree on publication date")
        chosen = {
            col: str(sheet.cell_value(7, col)).strip()
            for col in range(sheet.ncols)
            if str(sheet.cell_value(7, col)).strip() in SELECTED
        }
        if len(chosen) != 8:
            raise ValueError(f"ONS AWE CDID set drifted on {sheet_name}: {sorted(chosen.values())}")
        for col, cdid in chosen.items():
            sector = str(sheet.cell_value(4, col)).strip()
            series_id = f"ONS_AWE_{cdid}_{measure}"
            catalog[series_id] = {
                "source_id": "ons_awe_earn01",
                "name": f"{sector} {measure.lower()} pay",
                "description": f"Raw seasonally adjusted weekly earnings in pounds; official ONS CDID {cdid}.",
                "frequency": "monthly",
                "unit": "currency",
                "eco_group": "wages",
                "source_url": source_url,
                "last_publish_date": candidate.date(),
            }
            for row in range(8, sheet.nrows):
                label = sheet.cell_value(row, 0)
                if not isinstance(label, (int, float)):
                    continue
                raw = sheet.cell_value(row, col)
                if not isinstance(raw, (int, float)):
                    continue
                excel_date = xlrd.xldate_as_datetime(label, book.datemode)
                reference = excel_date.date().replace(day=1)
                key = (series_id, reference)
                if key in keys:
                    raise ValueError(f"Duplicate ONS AWE key {key}")
                keys.add(key)
                observations.append(Observation(series_id, reference, float(raw), snapshot_id))
    if release is None or len(catalog) != 16 or len(observations) < 3000:
        raise ValueError("ONS AWE workbook unexpectedly short")
    return observations, catalog, release


def collect() -> ExtractedData:
    fetched = datetime.now(UTC)
    with httpx.Client(
        timeout=REQUEST_TIMEOUT, headers={"User-Agent": USER_AGENT}, follow_redirects=True
    ) as client:
        page = client.get(LANDING)
        page.raise_for_status()
        matches = re.findall(
            r'href=["\']([^"\']*file\?uri=[^"\']+\.xls)["\']', page.text, re.IGNORECASE
        )
        if len(matches) != 1:
            raise ValueError(f"Expected one current EARN01 workbook, found {len(matches)}")
        url = urljoin(LANDING, matches[0])
        response = client.get(url)
        response.raise_for_status()
    body = response.content
    if not body or len(body) > MAX_DOWNLOAD_BYTES:
        raise ValueError(f"Invalid ONS AWE artifact size {len(body)}")
    digest = hashlib.sha256(body).hexdigest()
    observations, catalog, released = parse_xls(body, digest, url)
    snapshot = build_snapshot(
        "ons_awe_earn01",
        url,
        "earn01.xls",
        body,
        digest,
        response.headers.get("etag"),
        response.headers.get("last-modified"),
        fetched,
        released.date(),
    )
    latest_by_series = {
        series_id: max(o.reference_date for o in observations if o.series_id == series_id)
        for series_id in catalog
    }
    availability = {
        (o.series_id, o.reference_date): (
            released if o.reference_date == latest_by_series[o.series_id] else fetched,
            "official_timestamp"
            if o.reference_date == latest_by_series[o.series_id]
            else "first_seen",
            released.date() if o.reference_date == latest_by_series[o.series_id] else None,
        )
        for o in observations
    }
    return ExtractedData(observations, [snapshot], catalog, [released], availability)
