# Game Collection Manager

Interactive app to manage:
- Owned collection (`PS1`, `PS2`, `PS4`, `DS WII`)
- Wishlist with priority, target price, order tracking, and replacement-copy support
- Private, manually entered PriceCharting references and market-value estimates

When a wishlist item is received, the app collects purchase details and moves it to the collection. Destructive actions require confirmation.

## Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Import data from your Excel files

```bash
source .venv/bin/activate
python scripts/import_excel_data.py
```

This generates `data/seed.json` from:
- `Ps2 games.xlsx`
- `wishlist_videogiochi.xlsx`

## Restructure Excel layout (functional format)

```bash
source .venv/bin/activate
python scripts/restructure_excel_layout.py
```

What it does:
- rewrites `Ps2 games.xlsx` into:
  - `Collection` sheet (single normalized table)
  - `Summary` sheet
- rewrites `wishlist_videogiochi.xlsx` into:
  - `Wishlist` sheet (single normalized table with `In Transit` / `Received`)
  - `Summary` sheet
- creates timestamped backup files before overwrite

## Run the app

Use a local server (recommended, so `fetch` works correctly):

```bash
python3 -m http.server 8000
```

Then open:
- `http://localhost:8000`

## Two-way sync

### App -> Excel

1. In the app, click `Export App State` (downloads a JSON snapshot).
2. Sync that file back into Excel:

```bash
source .venv/bin/activate
python scripts/sync_state_to_excel.py --state /path/to/game-collection-state-YYYY-MM-DDTHH-MM-SS.json
```

This updates:
- `Ps2 games.xlsx`
- `wishlist_videogiochi.xlsx`
- `data/seed.json`

It also creates timestamped Excel backups before overwrite.

### Excel -> App

```bash
source .venv/bin/activate
python scripts/import_excel_data.py
```

Then in the app:
- either clear localStorage key `gameCollectionManager.v2` and refresh
- or import a previously exported JSON snapshot using `Import App State`

Imports are validated and previewed before they replace browser data. The previous state is retained in `gameCollectionManager.backup.v2` for recovery.

On startup, the app compares the browser's base revision with the current repository seed. If both versions changed, it offers to export the browser state, keep it, or load the repository version. Loading the repository version creates a local backup first.

## Private pricing references

Use the `Market` button beside a collection or wishlist game to open a matching PriceCharting search, save the exact product link, and manually enter Loose, CIB, New, box-only, or manual-only values.

Pricing is intentionally isolated in the browser under `gameCollectionManager.priceCharting.v1`. It is not included in the normal app-state export, Excel files, or `data/seed.json`. Use `Export Private Prices` and `Import Private Prices` to move or back up these values separately. The app does not scrape PriceCharting or use its paid API.

## Data fields

- Collection: platform, title, version, disc/manual condition, price, extra, note, acquired date, and source
- Wishlist: platform, title, note, priority, target price, in-transit/received status, ordered/received dates, listing URL, and replacement-copy flag

All collection and wishlist fields round-trip between JSON and Excel. Private pricing references do not.

## Tests

```bash
source .venv/bin/activate
python -m unittest discover -s tests -v
node --test tests/*.test.js
node --check app.js
```

## Notes

- App data is persisted in browser `localStorage` and automatically migrates from the legacy `gameCollectionManager.v1` key.
- Private pricing data uses a separate browser storage key and must be backed up separately.
- Seed revisions are content-based, so regenerating an unchanged seed does not create a false conflict.
- If you update Excel files and want fresh seed data, re-run:
  - `python scripts/import_excel_data.py`
- To reset app state, clear `localStorage` for the page or use browser dev tools.
