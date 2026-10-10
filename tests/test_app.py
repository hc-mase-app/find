from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

import app


class ApplicationFlowTest(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        app.USE_POSTGRES = False
        app.DATA_DIR = Path(self.temp_dir.name)
        app.DB_PATH = app.DATA_DIR / "test.db"
        app.BACKUP_DIR = app.DATA_DIR / "backup"
        app.init_db()

    def tearDown(self) -> None:
        self.temp_dir.cleanup()

    def test_inbound_sale_inventory_and_report(self) -> None:
        app.manual_inbound(
            {
                "imei": "123456789012345",
                "tipe": "REALME C53 6/128",
                "hpp": 480000,
                "modal_awal": 480000,
                "grade": "A",
                "spek": "6/128",
                "supplier": "Test",
            }
        )
        before = app.summary()
        self.assertEqual(before["stok"], 1)

        sale = app.sell_item(
            {
                "imei": "012345",
                "harga_jual": 600000,
                "customer": "Andi",
                "catatan": "Tes",
            }
        )
        self.assertGreater(sale["id"], 0)
        self.assertEqual(sale["laba_rugi"], 120000)

        after = app.summary()
        self.assertEqual(after["stok"], 0)
        self.assertEqual(after["jumlah_terjual_periode"], 1)
        self.assertEqual(after["total_jual"], 600000)
        self.assertEqual(app.list_inventory(status="TERJUAL")[0]["imei"], "123456789012345")


if __name__ == "__main__":
    unittest.main()
