import hashlib
import io
import json
import tempfile
import unittest
import zipfile
from pathlib import Path
from urllib.error import HTTPError
from unittest.mock import patch

import updater
import updates


class Response(io.BytesIO):
    def __init__(self, value: bytes):
        super().__init__(value)
        self.headers = {"Content-Length": str(len(value))}

    def __enter__(self):
        return self

    def __exit__(self, *_arguments):
        self.close()


class UpdateClientTests(unittest.TestCase):
    def test_semantic_version_comparison(self):
        self.assertTrue(updates.is_newer_version("v1.2.0", "1.1.9"))
        self.assertFalse(updates.is_newer_version("1.2.0", "1.2.0"))
        self.assertFalse(updates.is_newer_version("1.1.9", "1.2.0"))
        with self.assertRaises(ValueError):
            updates.version_tuple("latest")

    def test_latest_release_selects_verified_windows_asset(self):
        digest = "a" * 64
        payload = {
            "tag_name": "v2.0.0",
            "body": "Исправления",
            "html_url": "https://github.com/1-Choose-Goose/pokazaniya/releases/tag/v2.0.0",
            "assets": [{
                "name": updates.WINDOWS_ASSET_NAME,
                "browser_download_url": (
                    "https://github.com/1-Choose-Goose/pokazaniya/releases/"
                    "download/v2.0.0/Pokazaniya-Windows-x64.zip"
                ),
                "digest": f"sha256:{digest}",
                "size": 123,
            }],
        }
        with patch.object(
            updates, "_open", return_value=Response(json.dumps(payload).encode())
        ):
            result = updates.check_for_update("1.0.0")
        self.assertEqual(result.version, "2.0.0")
        self.assertEqual(result.sha256, digest)
        self.assertEqual(result.size, 123)

    def test_release_without_digest_is_rejected(self):
        payload = {
            "tag_name": "v2.0.0",
            "assets": [{
                "name": updates.WINDOWS_ASSET_NAME,
                "browser_download_url": "https://github.com/example/update.zip",
            }],
        }
        with patch.object(
            updates, "_open", return_value=Response(json.dumps(payload).encode())
        ), self.assertRaisesRegex(updates.UpdateError, "SHA-256"):
            updates.check_for_update("1.0.0")

    def test_inaccessible_private_releases_are_reported(self):
        error = HTTPError(updates.LATEST_RELEASE_API, 404, "Not Found", {}, None)
        with patch.object(updates, "_open", side_effect=error), self.assertRaisesRegex(
            updates.UpdateError, "публичным"
        ):
            updates.check_for_update("1.0.0")

    def test_download_is_atomic_and_checksum_is_verified(self):
        content = b"verified update archive"
        information = updates.UpdateInfo(
            "2.0.0",
            "https://github.com/example/update.zip",
            hashlib.sha256(content).hexdigest(),
            len(content),
            "",
            "",
        )
        with tempfile.TemporaryDirectory() as folder:
            destination = Path(folder) / "update.zip"
            progress = []
            with patch.object(updates, "_open", return_value=Response(content)):
                updates.download_update(
                    information,
                    destination,
                    progress=lambda downloaded, total: progress.append(
                        (downloaded, total)
                    ),
                )
            self.assertEqual(destination.read_bytes(), content)
            self.assertFalse(destination.with_suffix(".zip.part").exists())
            self.assertEqual(progress[-1], (len(content), len(content)))

            bad = updates.UpdateInfo(
                "2.0.0", information.download_url, "0" * 64, len(content), "", ""
            )
            with patch.object(
                updates, "_open", return_value=Response(content)
            ), self.assertRaisesRegex(updates.UpdateError, "сумма"):
                updates.download_update(bad, destination)
            self.assertEqual(destination.read_bytes(), content)

    def test_stale_update_lock_is_removed(self):
        with tempfile.TemporaryDirectory() as folder:
            lock = Path(folder) / "update.lock"
            lock.write_text(
                json.dumps({"pid": 1234, "created_at": 1}), encoding="utf-8"
            )
            with (
                patch.object(updates, "update_lock_path", return_value=lock),
                patch.object(updates.time, "time", return_value=3600),
                patch.object(updates, "_process_is_running", return_value=False),
            ):
                self.assertFalse(updates.is_update_in_progress())
            self.assertFalse(lock.exists())


class ApplyUpdateTests(unittest.TestCase):
    @staticmethod
    def make_archive(path: Path) -> None:
        with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as package:
            package.writestr("Pokazaniya.exe", b"new executable")
            package.writestr("_internal/PokazaniyaUpdater.exe", b"new updater")
            package.writestr("_internal/runtime.dll", b"new runtime")

    def make_install(self, root: Path, name: str = "Pokazaniya") -> Path:
        install = root / name
        (install / "_internal").mkdir(parents=True)
        (install / "data").mkdir()
        (install / ".yandex-sync").mkdir()
        (install / "Pokazaniya.exe").write_bytes(b"old executable")
        (install / "_internal" / "PokazaniyaUpdater.exe").write_bytes(b"old updater")
        (install / "_internal" / "runtime.dll").write_bytes(b"old runtime")
        (install / "data" / "pokazaniya.db").write_bytes(b"user readings")
        (install / "yandex-disk.json").write_text('{"token":"secret"}', encoding="utf-8")
        (install / ".yandex-sync" / "state.json").write_text("{}", encoding="utf-8")
        return install

    def test_update_replaces_program_and_preserves_user_data(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            install = self.make_install(root)
            download = root / "download"
            download.mkdir()
            archive = download / "update.zip"
            self.make_archive(archive)

            updater.apply_update(archive, install, restart=False)

            self.assertEqual((install / "Pokazaniya.exe").read_bytes(), b"new executable")
            self.assertEqual(
                (install / "_internal" / "runtime.dll").read_bytes(), b"new runtime"
            )
            self.assertEqual(
                (install / "data" / "pokazaniya.db").read_bytes(), b"user readings"
            )
            self.assertTrue((install / "yandex-disk.json").is_file())
            self.assertTrue((install / ".yandex-sync" / "state.json").is_file())
            self.assertFalse(download.exists())
            self.assertFalse(list(root.glob(".Pokazaniya-*")))

    def test_update_allows_a_dedicated_folder_with_any_name(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            install = self.make_install(root, "Новая папка (2)")
            archive = root / "update.zip"
            self.make_archive(archive)

            updater.apply_update(archive, install, restart=False)

            self.assertEqual((install / "Pokazaniya.exe").read_bytes(), b"new executable")
            self.assertEqual(
                (install / "data" / "pokazaniya.db").read_bytes(), b"user readings"
            )

    def test_update_preserves_unrelated_files_in_the_program_folder(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            install = self.make_install(root, "mixed-files")
            (install / "unrelated-document.txt").write_text("keep", encoding="utf-8")
            archive = root / "update.zip"
            self.make_archive(archive)

            updater.apply_update(archive, install, restart=False)

            self.assertEqual(
                (install / "unrelated-document.txt").read_text(encoding="utf-8"),
                "keep",
            )

    def test_archive_cannot_write_outside_staging_directory(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            archive = root / "bad.zip"
            with zipfile.ZipFile(archive, "w") as package:
                package.writestr("../outside.txt", b"unsafe")
            with self.assertRaisesRegex(updater.ApplyUpdateError, "опасный путь"):
                updater._safe_extract(archive, root / "staging")
            self.assertFalse((root / "outside.txt").exists())

    def test_failed_restart_rolls_back_program_and_user_data(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            install = self.make_install(root)
            archive = root / "update.zip"
            self.make_archive(archive)
            with patch.object(
                updater.subprocess, "Popen", side_effect=OSError("cannot start")
            ), self.assertRaisesRegex(OSError, "cannot start"):
                updater.apply_update(archive, install)
            self.assertEqual((install / "Pokazaniya.exe").read_bytes(), b"old executable")
            self.assertEqual(
                (install / "data" / "pokazaniya.db").read_bytes(), b"user readings"
            )
            self.assertFalse(list(root.glob(".Pokazaniya-*")))

    def test_refuses_to_replace_a_source_checkout(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            install = self.make_install(root)
            (install / ".git").mkdir()
            archive = root / "update.zip"
            self.make_archive(archive)
            with self.assertRaisesRegex(updater.ApplyUpdateError, "установка"):
                updater.apply_update(archive, install, restart=False)
            self.assertTrue((install / ".git").is_dir())


if __name__ == "__main__":
    unittest.main()
