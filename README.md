# Game Collection Manager

Interactive app to manage:
- your owned collection (`PS1`, `PS2`, `PS4`, `DS WII`)
- your wishlist with `In transit` and `Received` status

When a wishlist item is marked as `Received`, it is automatically moved to the collection.

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

## Notes

- App data is persisted in browser `localStorage`.
- If you update Excel files and want fresh seed data, re-run:
  - `python scripts/import_excel_data.py`
- To reset app state, clear `localStorage` for the page or use browser dev tools.
