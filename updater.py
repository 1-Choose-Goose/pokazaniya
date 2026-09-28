"""Отдельный атомарный обновлятор Windows-сборки PyInstaller."""

from __future__ import annotations

import argparse
import ctypes
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
import uuid
import zipfile
from collections.abc import Callable
from pathlib import Path


ProgressCallback = Callable[[int, str], None]


class ApplyUpdateError(RuntimeError):
    pass


class UpdateProgressWindow:
    """Окно установки в серо-зелёном стиле приложения «Мои показания»."""

    def __init__(self) -> None:
        self.root = None
        self.status = None
        self.progress = None
        if os.name != "nt":
            return
        import tkinter as tk
        from tkinter import ttk

        root = tk.Tk()
        root.withdraw()
        root.title("Обновление — Мои показания")
        root.resizable(False, False)
        root.protocol("WM_DELETE_WINDOW", lambda: None)
        root.attributes("-topmost", True)
        root.configure(background="#e9ebe8")
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
                "MyMeterReadings.Updater.1"
            )
        except (AttributeError, OSError):
            pass

        style = ttk.Style(root)
        style.theme_use("clam")
        style.configure("Update.TFrame", background="#f7f7f4")
        style.configure(
            "Update.Horizontal.TProgressbar",
            troughcolor="#d8e0da",
            background="#49645d",
            bordercolor="#d8e0da",
            lightcolor="#49645d",
            darkcolor="#49645d",
            thickness=16,
        )

        outer = tk.Frame(root, background="#e9ebe8", padx=12, pady=12)
        outer.pack(fill="both", expand=True)
        card = tk.Frame(
            outer,
            background="#f7f7f4",
            highlightthickness=1,
            highlightbackground="#d0d6d1",
            padx=26,
            pady=22,
        )
        card.pack(fill="both", expand=True)
        tk.Label(
            card,
            text="Устанавливаем обновление",
            background="#f7f7f4",
            foreground="#263b33",
            font=("Segoe UI", 16, "bold"),
            anchor="w",
        ).pack(fill="x")
        self.status = tk.StringVar(value="Подготовка…")
        tk.Label(
            card,
            textvariable=self.status,
            background="#f7f7f4",
            foreground="#40554d",
            font=("Segoe UI", 10),
            anchor="w",
        ).pack(fill="x", pady=(14, 8))
        self.progress = ttk.Progressbar(
            card,
            style="Update.Horizontal.TProgressbar",
            maximum=100,
            mode="determinate",
        )
        self.progress.pack(fill="x")
        tk.Label(
            card,
            text="После установки программа откроется автоматически",
            background="#f7f7f4",
            foreground="#66756d",
            font=("Segoe UI", 9),
            anchor="w",
        ).pack(fill="x", pady=(12, 0))

        root.update_idletasks()
        width, height = 520, 218
        x = max(0, (root.winfo_screenwidth() - width) // 2)
        y = max(0, (root.winfo_screenheight() - height) // 2)
        root.geometry(f"{width}x{height}+{x}+{y}")
        root.deiconify()
        self.root = root
        self.update(3, "Закрываем программу…")

    def update(self, value: int, text: str) -> None:
        if self.root is None:
            return
        self.status.set(text)
        self.progress["value"] = max(0, min(100, value))
        self.root.update_idletasks()
        self.root.update()

    def close(self) -> None:
        if self.root is not None:
            self.root.destroy()
            self.root = None


def wait_for_process(pid: int, timeout: float = 120.0) -> None:
    if pid <= 0 or os.name != "nt":
        return
    handle = ctypes.windll.kernel32.OpenProcess(0x00100000, False, pid)
    if not handle:
        return
    try:
        result = ctypes.windll.kernel32.WaitForSingleObject(
            handle, max(0, round(timeout * 1000))
        )
        if result == 0x00000102:
            raise ApplyUpdateError("Программа не завершилась вовремя")
    finally:
        ctypes.windll.kernel32.CloseHandle(handle)


def _safe_extract(
    archive: Path,
    destination: Path,
    progress: ProgressCallback | None = None,
) -> None:
    destination_root = destination.resolve()
    with zipfile.ZipFile(archive) as package:
        members = package.infolist()
        for item in members:
            mode = item.external_attr >> 16
            if mode & 0o170000 == 0o120000:
                raise ApplyUpdateError("Архив обновления содержит символическую ссылку")
            target = (destination / item.filename).resolve()
            if target != destination_root and destination_root not in target.parents:
                raise ApplyUpdateError("Архив обновления содержит опасный путь")

        total_size = max(1, sum(item.file_size for item in members))
        extracted = 0
        last_value = -1
        for item in members:
            target = destination / item.filename
            if item.is_dir():
                target.mkdir(parents=True, exist_ok=True)
                continue
            target.parent.mkdir(parents=True, exist_ok=True)
            with package.open(item) as source, target.open("wb") as output:
                while chunk := source.read(1024 * 1024):
                    output.write(chunk)
                    extracted += len(chunk)
                    value = 20 + min(35, int(extracted * 35 / total_size))
                    if progress is not None and value != last_value:
                        progress(value, "Распаковываем новую версию…")
                        last_value = value


def _set_windows_background_mode(enabled: bool) -> bool:
    if os.name != "nt":
        return False
    kernel32 = ctypes.windll.kernel32
    if enabled:
        process_changed = bool(
            kernel32.SetPriorityClass(kernel32.GetCurrentProcess(), 0x00004000)
        )
        thread_changed = bool(
            kernel32.SetThreadPriority(kernel32.GetCurrentThread(), 0x00010000)
        )
        return process_changed or thread_changed
    kernel32.SetThreadPriority(kernel32.GetCurrentThread(), 0x00020000)
    return True


def _payload_root(staging: Path, executable: str) -> Path:
    if (staging / executable).is_file():
        return staging
    children = [item for item in staging.iterdir() if item.is_dir()]
    if len(children) == 1 and (children[0] / executable).is_file():
        return children[0]
    raise ApplyUpdateError("В архиве нет новой версии Pokazaniya.exe")


def _move_preserved_items(backup: Path, installed: Path) -> list[str]:
    """Вернуть пользовательские данные и посторонние файлы в новую установку."""
    moved = []
    for source in backup.iterdir():
        name = source.name
        destination = installed / name
        if destination.exists():
            if name in {"data", "pokazaniya.db", ".yandex-sync"}:
                if destination.is_dir():
                    shutil.rmtree(destination)
                else:
                    destination.unlink()
            else:
                continue
        shutil.move(str(source), str(destination))
        moved.append(name)
    return moved


def apply_update(
    archive: Path,
    install_dir: Path,
    *,
    executable: str = "Pokazaniya.exe",
    parent_pid: int = 0,
    restart: bool = True,
    progress: ProgressCallback | None = None,
) -> None:
    archive = archive.resolve()
    install_dir = install_dir.resolve()
    if not archive.is_file() or not zipfile.is_zipfile(archive):
        raise ApplyUpdateError("Архив обновления повреждён")
    if (
        not install_dir.is_dir()
        or not (install_dir / executable).is_file()
        or not (install_dir / "_internal" / "PokazaniyaUpdater.exe").is_file()
        or not (install_dir / "_internal").is_dir()
        or (install_dir / ".git").exists()
    ):
        raise ApplyUpdateError("Не найдена текущая установка программы")

    report = progress or (lambda _value, _text: None)
    report(5, "Ожидаем завершения программы…")
    wait_for_process(parent_pid)
    suffix = uuid.uuid4().hex[:10]
    staging = install_dir.parent / f".Pokazaniya-update-{suffix}"
    backup = install_dir.parent / f".Pokazaniya-backup-{suffix}"
    failed_install: Path | None = None
    moved_items: list[str] = []
    try:
        staging.mkdir()
        report(20, "Распаковываем новую версию…")
        _safe_extract(archive, staging, progress=report)
        payload = _payload_root(staging, executable)
        if not (payload / "_internal").is_dir():
            raise ApplyUpdateError("В архиве нет зависимостей программы")

        report(55, "Заменяем файлы программы…")
        install_dir.rename(backup)
        try:
            payload.rename(install_dir)
            report(78, "Сохраняем показания и настройки…")
            moved_items = _move_preserved_items(backup, install_dir)
            if restart:
                subprocess.Popen(
                    [str(install_dir / executable)], close_fds=True, cwd=install_dir
                )
        except Exception:
            if install_dir.exists():
                for name in moved_items:
                    source = install_dir / name
                    destination = backup / name
                    if source.exists() and not destination.exists():
                        shutil.move(str(source), str(destination))
                failed_install = install_dir.with_name(
                    install_dir.name + f".failed-{suffix}"
                )
                install_dir.rename(failed_install)
            backup.rename(install_dir)
            raise
        report(90, "Завершаем установку…")
        shutil.rmtree(backup, ignore_errors=True)
        archive.unlink(missing_ok=True)
        try:
            archive.parent.rmdir()
        except OSError:
            pass
    finally:
        shutil.rmtree(staging, ignore_errors=True)
        if failed_install is not None:
            shutil.rmtree(failed_install, ignore_errors=True)


def _show_error(message: str) -> None:
    log = Path(tempfile.gettempdir()) / "Pokazaniya-updater-error.txt"
    log.write_text(message, encoding="utf-8")
    if os.name == "nt":
        ctypes.windll.user32.MessageBoxW(
            None,
            message + f"\n\nПодробности: {log}",
            "Мои показания — ошибка обновления",
            0x10,
        )


def _schedule_self_cleanup() -> None:
    if os.name != "nt" or not getattr(sys, "frozen", False):
        return
    executable = Path(sys.executable).resolve()
    folder = executable.parent
    temporary_root = Path(tempfile.gettempdir()).resolve()
    if temporary_root not in folder.parents or not folder.name.startswith(
        "Pokazaniya-updater-"
    ):
        return
    descriptor, cleanup_name = tempfile.mkstemp(
        prefix="Pokazaniya-cleanup-", suffix=".cmd", dir=temporary_root
    )
    cleanup_script = Path(cleanup_name)
    with os.fdopen(descriptor, "w", encoding="mbcs", newline="\r\n") as script:
        script.write(
            "@echo off\n"
            "for /l %%i in (1,1,60) do (\n"
            f'  rmdir /s /q "{folder}" 2>nul\n'
            f'  if not exist "{folder}" goto cleanup_done\n'
            "  ping 127.0.0.1 -n 2 >nul\n"
            ")\n"
            ":cleanup_done\n"
            'del /f /q "%~f0"\n'
        )
    creation_flags = subprocess.CREATE_NEW_PROCESS_GROUP | subprocess.CREATE_NO_WINDOW
    try:
        subprocess.Popen(
            ["cmd.exe", "/d", "/c", str(cleanup_script)],
            close_fds=True,
            creationflags=creation_flags,
            cwd=temporary_root,
        )
    except OSError:
        cleanup_script.unlink(missing_ok=True)
        raise


def _claim_update_lock(lock_file: Path | None) -> None:
    if lock_file is None:
        return
    lock_file.parent.mkdir(parents=True, exist_ok=True)
    temporary = lock_file.with_suffix(".tmp")
    temporary.write_text(
        json.dumps({"pid": os.getpid(), "created_at": time.time()}),
        encoding="utf-8",
    )
    temporary.replace(lock_file)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--archive", required=True, type=Path)
    parser.add_argument("--install-dir", required=True, type=Path)
    parser.add_argument("--pid", required=True, type=int)
    parser.add_argument("--executable", default="Pokazaniya.exe")
    parser.add_argument("--lock-file", type=Path)
    parser.add_argument("--no-restart", action="store_true", help=argparse.SUPPRESS)
    arguments: argparse.Namespace | None = None
    progress_window: UpdateProgressWindow | None = None
    background_mode = False
    try:
        arguments = parser.parse_args()
        _claim_update_lock(arguments.lock_file)
        progress_window = UpdateProgressWindow()
        background_mode = _set_windows_background_mode(True)
        time.sleep(0.25)
        apply_update(
            arguments.archive,
            arguments.install_dir,
            executable=arguments.executable,
            parent_pid=arguments.pid,
            restart=False,
            progress=progress_window.update,
        )
        if background_mode:
            _set_windows_background_mode(False)
            background_mode = False
        if arguments.lock_file is not None:
            arguments.lock_file.unlink(missing_ok=True)
        if not arguments.no_restart:
            progress_window.update(100, "Готово. Запускаем программу…")
            time.sleep(0.4)
            progress_window.close()
            subprocess.Popen(
                [str(arguments.install_dir / arguments.executable)],
                close_fds=True,
                cwd=arguments.install_dir,
            )
        return 0
    except Exception as error:  # noqa: BLE001 — последняя граница обновлятора
        _show_error(str(error))
        return 1
    finally:
        if background_mode:
            _set_windows_background_mode(False)
        if progress_window is not None:
            progress_window.close()
        if arguments is not None:
            if arguments.lock_file is not None:
                arguments.lock_file.unlink(missing_ok=True)
            arguments.archive.unlink(missing_ok=True)
            try:
                arguments.archive.parent.rmdir()
            except OSError:
                pass
        _schedule_self_cleanup()


if __name__ == "__main__":
    raise SystemExit(main())
