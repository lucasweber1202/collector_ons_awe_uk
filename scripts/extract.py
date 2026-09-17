"""Selected raw Average Weekly Earnings series from the official ONS EARN01 workbook."""

from __future__ import annotations

import hashlib
import re
from dataclasses import dataclass
from datetime import UTC, date, datetime
from typing import Any
from urllib.parse import urljoin

import httpx
import xlrd

from scripts.config import MAX_DOWNLOAD_BYTES, REQUEST_TIMEOUT, USER_AGENT
from scripts.snapshots import Snapshot, build_snapshot
from scripts.time_series import Observation

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
