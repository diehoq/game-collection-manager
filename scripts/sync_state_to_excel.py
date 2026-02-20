#!/usr/bin/env python3
"""Sync exported app state JSON back into collection and wishlist Excel files."""

from __future__ import annotations

import argparse
import json
import shutil
from pathlib import Path
from typing import Any, Dict, Iterable, List, Tuple

from import_excel_data import (
    DEFAULT_COLLECTION_FILE,
    DEFAULT_OUTPUT_FILE,
    DEFAULT_WISHLIST_FILE,
    CollectionRecord,
    WishlistRecord,
    build_collection_key,
    to_json_payload,
    transfer_received_items,
)
from restructure_excel_layout import backup_path, build_collection_workbook, build_wishlist_workbook


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Sync app-exported JSON state back into Excel workbooks."
    )
    parser.add_argument(
        "--state",
        required=True,
        help="Path to JSON file exported from the web app.",
    )
    parser.add_argument(
        "--collection",
        default=DEFAULT_COLLECTION_FILE,
        help="Collection workbook path.",
    )
    parser.add_argument(
        "--wishlist",
        default=DEFAULT_WISHLIST_FILE,
        help="Wishlist workbook path.",
    )
    parser.add_argument(
        "--seed-out",
        default=DEFAULT_OUTPUT_FILE,
        help="Seed JSON path to update after sync.",
    )
    parser.add_argument(
        "--no-backup",
        action="store_true",
        help="Overwrite Excel files without creating timestamped backups first.",
    )
    parser.add_argument(
        "--no-update-seed",
        action="store_true",
        help="Do not regenerate seed JSON.",
    )
    return parser.parse_args()


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def normalize_bool(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    text = normalize_text(value).casefold()
    return text in {"true", "1", "yes", "y", "x"}


def row_text(row: Dict[str, Any], *keys: str) -> str:
    for key in keys:
        if key in row:
            return normalize_text(row.get(key))
    return ""


def row_bool(row: Dict[str, Any], *keys: str) -> bool:
    for key in keys:
        if key in row:
            return normalize_bool(row.get(key))
    return False


def parse_collection_rows(rows: Iterable[Dict[str, Any]]) -> List[CollectionRecord]:
    parsed: List[CollectionRecord] = []
    seen = set()
    for row in rows:
        if not isinstance(row, dict):
            continue
        platform = row_text(row, "platform")
        title = row_text(row, "title")
        if not platform or not title:
            continue
        key = build_collection_key(platform, title)
        if key in seen:
            continue
        seen.add(key)
        parsed.append(
            CollectionRecord(
                id="",
                platform=platform,
                title=title,
                version=row_text(row, "version"),
                cd_condition=row_text(row, "cdCondition", "cd_condition"),
                manual_condition=row_text(row, "manualCondition", "manual_condition"),
                price=row_text(row, "price"),
                extra=row_text(row, "extra"),
                note=row_text(row, "note"),
            )
        )

    return [
        CollectionRecord(
            id=f"c{idx}",
            platform=row.platform,
            title=row.title,
            version=row.version,
            cd_condition=row.cd_condition,
            manual_condition=row.manual_condition,
            price=row.price,
            extra=row.extra,
            note=row.note,
        )
        for idx, row in enumerate(parsed, start=1)
    ]


def parse_wishlist_rows(rows: Iterable[Dict[str, Any]]) -> List[WishlistRecord]:
    parsed: List[WishlistRecord] = []
    seen = set()
    for row in rows:
        if not isinstance(row, dict):
            continue
        platform = row_text(row, "platform")
        title = row_text(row, "title")
        if not platform or not title:
            continue
        key = build_collection_key(platform, title)
        if key in seen:
            continue
        seen.add(key)
        parsed.append(
            WishlistRecord(
                id="",
                platform=platform,
                title=title,
                note=row_text(row, "note"),
                in_transit=row_bool(row, "inTransit", "in_transit"),
                received=row_bool(row, "received"),
            )
        )

    return [
        WishlistRecord(
            id=f"w{idx}",
            platform=row.platform,
            title=row.title,
            note=row.note,
            in_transit=row.in_transit,
            received=row.received,
        )
        for idx, row in enumerate(parsed, start=1)
    ]


def load_state(path: Path) -> Tuple[List[CollectionRecord], List[WishlistRecord]]:
    if not path.exists():
        raise FileNotFoundError(f"State file not found: {path}")
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise ValueError("State JSON root must be an object.")

    collection_raw = data.get("collection")
    wishlist_raw = data.get("wishlist")
    if not isinstance(collection_raw, list) or not isinstance(wishlist_raw, list):
        raise ValueError("State JSON must contain 'collection' and 'wishlist' arrays.")

    collection = parse_collection_rows(collection_raw)
    wishlist = parse_wishlist_rows(wishlist_raw)
    collection, wishlist = transfer_received_items(collection, wishlist)
    return collection, wishlist


def write_seed(path: Path, collection: List[CollectionRecord], wishlist: List[WishlistRecord]) -> None:
    payload = to_json_payload(collection, wishlist)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")


def main() -> int:
    args = parse_args()
    state_path = Path(args.state)
    collection_path = Path(args.collection)
    wishlist_path = Path(args.wishlist)
    seed_out = Path(args.seed_out)

    collection_rows, wishlist_rows = load_state(state_path)

    backups: List[Path] = []
    if not args.no_backup:
        collection_backup = backup_path(collection_path)
        wishlist_backup = backup_path(wishlist_path)
        shutil.copy2(collection_path, collection_backup)
        shutil.copy2(wishlist_path, wishlist_backup)
        backups.extend([collection_backup, wishlist_backup])

    collection_wb = build_collection_workbook(collection_rows)
    wishlist_wb = build_wishlist_workbook(wishlist_rows)
    collection_wb.save(collection_path)
    wishlist_wb.save(wishlist_path)

    if not args.no_update_seed:
        write_seed(seed_out, collection_rows, wishlist_rows)

    print(f"Synced state from {state_path}")
    print(f"Updated {collection_path} ({len(collection_rows)} collection items)")
    print(f"Updated {wishlist_path} ({len(wishlist_rows)} wishlist items)")
    if not args.no_update_seed:
        print(f"Updated {seed_out}")
    for backup in backups:
        print(f"Backup created: {backup}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
