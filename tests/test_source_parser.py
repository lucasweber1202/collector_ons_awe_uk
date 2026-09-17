import pytest
from xlrd.biffh import XLRDError

from scripts.extract import parse_xls


def test_parser_fails_closed_on_invalid_workbook() -> None:
    with pytest.raises(XLRDError):
        parse_xls(b"not an xls", "snapshot", "https://ons.gov.uk")
