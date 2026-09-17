# collector_ons_awe_uk

Standalone collector for 16 official ONS EARN01 Average Weekly Earnings CDIDs: total and regular pay for the whole economy, private/public sectors, services, finance/business services, manufacturing, construction and wholesale/retail/hospitality. It contains 5,104 monthly observations from 2000-01 through 2026-07.

The latest point in each series uses the official release timestamp. Earlier values in the current workbook are `first_seen` unless an archived release establishes availability, preventing revision look-ahead.

## Install and run (PowerShell)

```powershell
py -3.11 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install .
Copy-Item .env.example .env
pytest -q
python main.py
```

Set `COLLECTOR_DB_URL` and allow `ons.gov.uk`. The official XLS is parsed by the declared `xlrd` dependency. Databricks is optional via `.[databricks]`. Source smoke: `python -c "from scripts.extract import collect; x=collect(); print(len(x.catalog), len(x.observations))"`.

See [METHODOLOGY.md](METHODOLOGY.md) and [POINT_IN_TIME.md](POINT_IN_TIME.md).
