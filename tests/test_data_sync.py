from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from import_excel_data import (  # noqa: E402
    CollectionRecord,
    WishlistRecord,
    parse_collection_sheet,
    parse_wishlist_sheet,
    transfer_received_items,
)
from restructure_excel_layout import (  # noqa: E402
    build_collection_workbook,
    build_wishlist_workbook,
)
from sync_state_to_excel import parse_collection_rows, parse_wishlist_rows  # noqa: E402


class DataSyncTests(unittest.TestCase):
    def test_json_fields_survive_excel_round_trip(self) -> None:
        collection = [
            CollectionRecord(
                id="c7",
                platform="PS2",
                title="Example Game",
                version="Collector's Edition",
                cd_condition="4.5",
                manual_condition="5",
                price="24.99",
                extra="Slipcase",
                note="Complete",
                acquired_date="2026-08-01",
                source="Local shop",
            )
        ]
        wishlist = [
            WishlistRecord(
                id="w4",
                platform="PS4",
                title="Wanted Game",
                note="PAL copy",
                in_transit=True,
                received=False,
                priority="High",
                target_price="30",
                ordered_date="2026-07-31",
                listing_url="https://example.com/game",
                replacement=True,
            )
        ]

        with tempfile.TemporaryDirectory() as directory:
            collection_path = Path(directory) / "collection.xlsx"
            wishlist_path = Path(directory) / "wishlist.xlsx"
            build_collection_workbook(collection).save(collection_path)
            build_wishlist_workbook(wishlist).save(wishlist_path)

            parsed_collection = parse_collection_sheet(collection_path)
            parsed_wishlist = parse_wishlist_sheet(wishlist_path)

        self.assertEqual(parsed_collection[0].acquired_date, "2026-08-01")
        self.assertEqual(parsed_collection[0].source, "Local shop")
        self.assertEqual(parsed_wishlist[0].priority, "High")
        self.assertEqual(parsed_wishlist[0].target_price, "30")
        self.assertEqual(parsed_wishlist[0].ordered_date, "2026-07-31")
        self.assertEqual(parsed_wishlist[0].listing_url, "https://example.com/game")
        self.assertTrue(parsed_wishlist[0].replacement)

    def test_received_item_moves_to_collection_with_received_date(self) -> None:
        wishlist = [
            WishlistRecord(
                id="w1",
                platform="PS4",
                title="Delivered Game",
                note="Sealed",
                in_transit=True,
                received=True,
                received_date="2026-08-01",
            )
        ]

        collection, remaining = transfer_received_items([], wishlist)

        self.assertEqual(remaining, [])
        self.assertEqual(collection[0].title, "Delivered Game")
        self.assertEqual(collection[0].acquired_date, "2026-08-01")

    def test_owned_wishlist_item_is_marked_as_replacement(self) -> None:
        collection = parse_collection_rows([{"platform": "PS2", "title": "Same Game"}])
        wishlist = parse_wishlist_rows([{"platform": "PS2", "title": "Same Game"}])

        _, wishlist = transfer_received_items(collection, wishlist)

        self.assertTrue(wishlist[0].replacement)


if __name__ == "__main__":
    unittest.main()
