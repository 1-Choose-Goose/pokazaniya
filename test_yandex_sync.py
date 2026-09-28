import hashlib
from pathlib import Path
import sqlite3
import tempfile
import time
import unittest
from unittest.mock import patch

from main import Database, Entry
from yandex_sync import Conflict, SyncEngine, digest, load_token, save_token, snapshot, validate


class FakeDisk:
    def __init__(self):
        self.data = None
        self.uploads = 0
        self.offline = False

    def metadata(self):
        if self.offline:
            raise OSError("offline")
        return {"sha256": hashlib.sha256(self.data).hexdigest()} if self.data else None

    def upload(self, path):
        self.metadata()
        self.data = path.read_bytes()
        self.uploads += 1
        return self.metadata()

    def download(self, path):
        meta = self.metadata()
        path.write_bytes(self.data)
        db = sqlite3.connect(path)
        try:
            validate(db)
            return digest(db), meta
        finally:
            db.close()


class SyncTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.db = Database(self.root / "data.db")
        self.api = FakeDisk()
        self.engine = SyncEngine(self.root / "data.db", self.root / "state", lambda: self.api)

    def tearDown(self):
        self.db.close()
        self.temp.cleanup()

    def test_saved_add_edit_delete_and_settings_replace_one_cloud_file(self):
        self.engine.once()
        item = Entry("Газ", "01.09.2026", None)
        item.id = self.db.add(item)
        self.engine.once()
        item.note = "changed"
        self.db.update(item)
        self.engine.once()
        self.db.delete(item.id)
        self.engine.once()
        self.db.set_setting("gas_meter", "123")
        self.engine.once()
        self.assertEqual(self.api.uploads, 5)
        self.engine.once()
        self.assertEqual(self.api.uploads, 5)
        restored = self.root / "restored.db"
        self.api.download(restored)
        check = Database(restored)
        self.assertEqual(check.all(), [])
        self.assertEqual(check.setting("gas_meter"), "123")
        check.close()

    def test_offline_change_survives_restart(self):
        self.engine.once()
        self.db.add(Entry("Газ", "01.09.2026", None))
        self.api.offline = True
        with self.assertRaises(OSError):
            self.engine.once()
        self.api.offline = False
        restarted = SyncEngine(self.root / "data.db", self.root / "state", lambda: self.api)
        restarted.once()
        self.assertEqual(self.api.uploads, 2)

    def test_different_cloud_database_is_not_overwritten(self):
        self.engine.once()
        other = Database(self.root / "other.db")
        other.add(Entry("Вода", "01.09.2026", None))
        other.close()
        self.api.data = (self.root / "other.db").read_bytes()
        old = self.api.data
        self.db.add(Entry("Газ", "01.09.2026", None))
        with self.assertRaises(Conflict):
            self.engine.once()
        self.assertEqual(self.api.data, old)
        self.engine.once(force=True)
        self.assertNotEqual(self.api.data, old)

    def test_fresh_empty_database_cannot_overwrite_cloud(self):
        self.db.add(Entry("Газ", "01.09.2026", None))
        self.engine.once()
        empty = Database(self.root / "empty.db")
        empty.close()
        new = SyncEngine(self.root / "empty.db", self.root / "newstate", lambda: self.api)
        with self.assertRaises(Conflict):
            new.once()
        self.assertEqual(self.api.uploads, 1)

    def test_transferred_database_adopts_identical_cloud(self):
        self.db.add(Entry("Газ", "01.09.2026", None))
        self.engine.once()
        moved = self.root / "moved.db"
        snapshot(self.root / "data.db", moved)
        engine = SyncEngine(moved, self.root / "newstate", lambda: self.api)
        engine.once()
        self.assertEqual(self.api.uploads, 1)

    def test_missing_file_is_not_created_or_uploaded(self):
        engine = SyncEngine(self.root / "missing.db", self.root / "newstate", lambda: self.api)
        with self.assertRaises(sqlite3.OperationalError):
            engine.once()
        self.assertFalse((self.root / "missing.db").exists())
        self.assertEqual(self.api.uploads, 0)

    def test_corrupt_cloud_database_rejected(self):
        self.api.data = b"not a database"
        with self.assertRaises(sqlite3.DatabaseError):
            self.api.download(self.root / "bad.db")

    def test_wal_snapshot_includes_committed_data(self):
        self.db.connection.execute("PRAGMA journal_mode=WAL")
        self.db.add(Entry("Газ", "01.09.2026", None))
        self.engine.once()
        target = self.root / "restored.db"
        self.api.download(target)
        restored = Database(target)
        self.assertEqual(len(restored.all()), 1)
        restored.close()

    def test_plain_token_is_portable(self):
        with patch("yandex_sync.TOKEN_PATH", self.root / "token.json"):
            save_token("test-token")
            self.assertEqual(load_token(), "test-token")
            self.assertIn("test-token", (self.root / "token.json").read_text())

    def test_background_worker_uploads_only_after_save_request(self):
        self.engine.start()
        time.sleep(0.05)
        self.assertEqual(self.api.uploads, 0)
        self.db.add(Entry("Газ", "01.09.2026", None))
        self.engine.request_sync()
        self.assertIn("отправка после сохранения", self.engine.status)
        self.assertTrue(self.engine.idle.wait(2))
        self.assertEqual(self.api.uploads, 1)
        self.assertTrue(self.engine.status.startswith("На Диске:"))
        self.engine.stop()

    def test_save_during_upload_is_not_lost(self):
        original_upload = self.api.upload

        def upload(path):
            if self.api.uploads == 0:
                concurrent = Database(self.root / "data.db")
                concurrent.add(Entry("ТКО", "02.09.2026", None))
                concurrent.close()
                self.engine.request_sync()
            return original_upload(path)

        self.api.upload = upload
        self.db.add(Entry("Газ", "01.09.2026", None))
        self.engine.start()
        self.engine.request_sync()
        self.assertTrue(self.engine.idle.wait(3))
        self.assertEqual(self.api.uploads, 2)
        restored = self.root / "restored-after-race.db"
        self.api.download(restored)
        check = Database(restored)
        self.assertEqual(len(check.all()), 2)
        check.close()
        self.engine.stop()


if __name__ == "__main__":
    unittest.main()
