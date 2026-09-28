"""Single-file Yandex Disk sync. Credentials never travel with the database."""
from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path
import queue
import sqlite3
import sys
import tempfile
import threading
import time
import tkinter as tk
from tkinter import messagebox, ttk
import urllib.error
import urllib.parse
import urllib.request
import webbrowser

CLIENT_ID = "fb49160dd62e4b918edaffa444988b25"
REMOTE = "disk:/MyDate/Показания/pokazaniya.db"
AUTH_URL = "https://oauth.yandex.ru/authorize?response_type=token&client_id=" + CLIENT_ID
RESOURCE_DIR = Path(__file__).resolve().parent
APP_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else RESOURCE_DIR
CONFIG_DIR = APP_DIR / ".yandex-sync"
TOKEN_PATH = APP_DIR / "yandex-disk.json"


def center_window(window):
    """Place a fully laid-out Tk window in the exact center of the screen."""
    window.update_idletasks()
    width = window.winfo_reqwidth()
    height = window.winfo_reqheight()
    x = max(0, (window.winfo_screenwidth() - width) // 2)
    y = max(0, (window.winfo_screenheight() - height) // 2)
    window.geometry(f"{width}x{height}+{x}+{y}")


def show_centered(window):
    center_window(window)
    window.deiconify()
    window.grab_set()
    window.focus_set()


def atomic_write(path: Path, data: bytes):
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, name = tempfile.mkstemp(dir=path.parent, prefix=".pending-")
    try:
        with os.fdopen(fd, "wb") as stream:
            stream.write(data)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(name, path)
    finally:
        Path(name).unlink(missing_ok=True)


def save_token(token: str):
    token = token.strip()
    if not token or any(c.isspace() for c in token) or len(token) > 4096:
        raise ValueError("Вставьте токен, показанный Яндексом после входа.")
    atomic_write(TOKEN_PATH, json.dumps({"token": token}).encode())


def load_token() -> str:
    try:
        token = json.loads(TOKEN_PATH.read_text(encoding="utf-8-sig"))["token"]
        if not isinstance(token, str) or not token.strip():
            raise ValueError
        return token.strip()
    except (FileNotFoundError, ValueError, KeyError):
        raise ValueError("Подключите Яндекс Диск: рядом с программой нет сохранённого доступа.") from None


def digest(connection: sqlite3.Connection) -> str:
    return hashlib.sha256("\n".join(connection.iterdump()).encode()).hexdigest()


def validate(connection: sqlite3.Connection):
    if connection.execute("PRAGMA quick_check").fetchone()[0] != "ok":
        raise ValueError("Файл базы повреждён.")
    columns = {r[1] for r in connection.execute("PRAGMA table_info(entries)")}
    required = {"id", "service", "period", "amount", "payment_date", "paid", "meter_1",
                "meter_2", "value_1", "value_2", "note", "created_at"}
    if not required <= columns:
        raise ValueError("На Диске находится не база программы «Мои показания».")
    if not {"key", "value"} <= {r[1] for r in connection.execute("PRAGMA table_info(settings)")}:
        raise ValueError("В базе отсутствуют настройки программы.")


def snapshot(path: Path, target: Path) -> str:
    # mode=ro prevents accidentally recreating a deleted local database.
    source = sqlite3.connect(path.resolve().as_uri() + "?mode=ro", uri=True, timeout=5)
    dest = sqlite3.connect(target)
    try:
        source.backup(dest)
        validate(dest)
        return digest(dest)
    finally:
        dest.close()
        source.close()


class Conflict(ValueError):
    pass


class DiskAPI:
    def __init__(self, token: str):
        self.token = token

    def request(self, endpoint, method="GET", **params):
        url = "https://cloud-api.yandex.net/v1/disk/" + endpoint
        if params:
            url += "?" + urllib.parse.urlencode(params)
        request = urllib.request.Request(url, method=method,
                                         headers={"Authorization": "OAuth " + self.token})
        try:
            with urllib.request.urlopen(request, timeout=20) as response:
                body = response.read()
                return json.loads(body) if body else {}
        except urllib.error.HTTPError as exc:
            if exc.code in (401, 403):
                raise ValueError("Доступ отклонён. Подключите Яндекс заново с правами чтения и записи.") from None
            raise

    def metadata(self):
        try:
            return self.request("resources", path=REMOTE)
        except urllib.error.HTTPError as exc:
            if exc.code == 404:
                return None
            raise

    def folders(self):
        for folder in ("disk:/MyDate", REMOTE.rsplit("/", 1)[0]):
            try:
                self.request("resources", "PUT", path=folder)
            except urllib.error.HTTPError as exc:
                if exc.code != 409:
                    raise

    @staticmethod
    def transfer_url(url):
        parsed = urllib.parse.urlparse(url)
        if parsed.scheme != "https" or not parsed.hostname or not any(
            parsed.hostname == domain or parsed.hostname.endswith("." + domain)
            for domain in ("yandex.net", "yandex.ru", "yandex.com")
        ):
            raise ValueError("Яндекс вернул неизвестный адрес передачи файла.")
        return url

    def upload(self, path):
        self.folders()
        link = self.request("resources/upload", path=REMOTE, overwrite="true")
        data = path.read_bytes()
        request = urllib.request.Request(self.transfer_url(link["href"]), data=data, method="PUT")
        with urllib.request.urlopen(request, timeout=45) as response:
            if response.status not in (200, 201, 204):
                raise ValueError("Яндекс не подтвердил загрузку.")
        meta = self.metadata()
        if not meta or meta.get("sha256") != hashlib.sha256(data).hexdigest():
            raise ValueError("Не удалось подтвердить целостность загруженной базы. Повторим проверку.")
        return meta

    def download(self, path):
        meta = self.metadata()
        if not meta:
            raise ValueError("На Диске пока нет базы.")
        link = self.request("resources/download", path=REMOTE)
        with urllib.request.urlopen(self.transfer_url(link["href"]), timeout=45) as response:
            data = response.read()
        if meta.get("sha256") != hashlib.sha256(data).hexdigest():
            raise ValueError("Файл изменился во время скачивания. Повторите восстановление.")
        path.write_bytes(data)
        connection = sqlite3.connect(path)
        try:
            validate(connection)
            local_hash = digest(connection)
        finally:
            connection.close()
        return local_hash, meta


class SyncEngine:
    def __init__(self, path: Path, config_dir=None, api_factory=None):
        self.path = path.resolve()
        self.config_dir = Path(config_dir or CONFIG_DIR)
        self.state_path = self.config_dir / ("state.json")
        self.api_factory = api_factory or (lambda: DiskAPI(load_token()))
        self.lock = threading.Lock()
        self.stop_event = threading.Event()
        self.wake = threading.Event()
        self.idle = threading.Event()
        self.idle.set()
        self.request_guard = threading.Lock()
        self.requested = 0
        self.completed = 0
        previous = self.state().get("time")
        self.status = "На Диске: " + previous if previous else "Яндекс Диск: отправка после сохранения"
        self.synced = False
        self.thread = None

    def state(self):
        try:
            return json.loads(self.state_path.read_text(encoding="utf-8"))
        except (FileNotFoundError, ValueError):
            return {}

    def record(self, local_hash, meta):
        atomic_write(self.state_path, json.dumps({"local": local_hash, "remote": meta["sha256"],
                    "time": time.strftime("%d.%m.%Y %H:%M:%S")}).encode())

    def once(self, force=False):
        with self.lock, tempfile.TemporaryDirectory(prefix="pokazaniya-") as temp:
            local = Path(temp) / "upload.db"
            current = snapshot(self.path, local)
            state = self.state()
            api = self.api_factory()
            remote = api.metadata()
            remote_hash = remote.get("sha256") if remote else None
            # Adopt an identical cloud file after a move or an uncertain previous upload.
            if remote_hash and (remote_hash != state.get("remote") or not state):
                downloaded = Path(temp) / "compare.db"
                remote_digest, remote = api.download(downloaded)
                if remote_digest == current:
                    self.record(current, remote)
                    return
                if not force:
                    raise Conflict("База на Диске отличается. Откройте «Яндекс Диск» и выберите нужную версию.")
            if not force and remote and current == state.get("local") and remote_hash == state.get("remote"):
                return
            self.synced = False
            self.status = "Яндекс Диск: отправка…"
            meta = api.upload(local)
            self.record(current, meta)

    def run(self):
        while True:
            self.wake.wait()
            self.wake.clear()
            if self.stop_event.is_set():
                self.idle.set()
                return
            with self.request_guard:
                target_request = self.requested
            try:
                self.once()
                self.synced = True
                self.status = "На Диске: " + self.state().get("time", "")
            except Conflict as exc:
                self.synced = False
                self.status = str(exc)
            except ValueError as exc:
                self.synced = False
                self.status = str(exc)
            except Exception:
                self.synced = False
                self.status = "Не отправлено: нет связи или ошибка Диска. Нажмите «Сохранить» ещё раз."
            finally:
                with self.request_guard:
                    self.completed = target_request
                    if self.completed == self.requested:
                        self.idle.set()
                    else:
                        self.wake.set()

    def start(self):
        if self.thread is not None and self.thread.is_alive():
            return
        self.stop_event.clear()
        self.thread = threading.Thread(target=self.run, daemon=True)
        self.thread.start()

    def request_sync(self):
        """Queue one upload after a successful local Save operation."""
        with self.request_guard:
            self.requested += 1
            self.synced = False
            self.status = "Яндекс Диск: отправка после сохранения…"
            self.idle.clear()
            self.wake.set()

    def stop(self, wait=True):
        self.stop_event.set()
        self.wake.set()
        if wait and self.thread is not None:
            self.thread.join()

    def wait_until_idle(self):
        if not self.idle.wait(75):
            raise ValueError("Яндекс Диск не завершил отправку за 75 секунд.")


def background(parent, title, action):
    """Modal progress with an active Tk event loop; network never runs on UI thread."""
    dialog = tk.Toplevel(parent)
    dialog.withdraw()
    dialog.title(title)
    dialog.transient(parent)
    dialog.protocol("WM_DELETE_WINDOW", lambda: None)
    ttk.Label(dialog, text=title + "…", padding=25).pack()
    show_centered(dialog)
    result = queue.Queue()
    def work():
        try:
            result.put((True, action()))
        except Exception as exc:
            result.put((False, exc))
    threading.Thread(target=work, daemon=True).start()
    outcome = []
    def poll():
        try:
            outcome.extend(result.get_nowait())
            dialog.destroy()
        except queue.Empty:
            dialog.after(100, poll)
    poll()
    parent.wait_window(dialog)
    if not outcome[0]:
        error = outcome[1]
        if isinstance(error, (ValueError, Conflict)):
            raise error
        raise ValueError("Не удалось связаться с Диском. Проверьте интернет и повторите.") from None
    return outcome[1]


def connect(parent):
    dialog = tk.Toplevel(parent)
    dialog.withdraw()
    dialog.title("Подключить Яндекс Диск")
    dialog.transient(parent)
    ttk.Label(dialog, text="1. Откройте вход и разрешите чтение и запись Диска.\n"
              "2. Скопируйте показанный Яндексом токен и вставьте ниже.\n"
              "При переносе всей папки программы повторный вход не нужен.", padding=15).pack()
    ttk.Button(dialog, text="Открыть вход в Яндекс", command=lambda: webbrowser.open(AUTH_URL)).pack(pady=5)
    token = tk.StringVar()
    ttk.Entry(dialog, textvariable=token, show="•", width=65).pack(padx=15, pady=10)
    success = []
    def submit():
        try:
            value = token.get().strip()
            # Read a resource rather than Disk capacity: no disk.info permission needed.
            background(dialog, "Проверка доступа", lambda: DiskAPI(value).request("resources", path="disk:/"))
            save_token(value)
            token.set("")
            success.append(True)
            dialog.destroy()
        except ValueError as exc:
            messagebox.showerror("Яндекс Диск", str(exc), parent=dialog)
    ttk.Button(dialog, text="Подключить", command=submit).pack(pady=15)
    show_centered(dialog)
    parent.wait_window(dialog)
    return bool(success)


def prepare_database(parent, path: Path) -> bool:
    """Never silently create an empty DB when a database was lost or not transferred."""
    if path.exists():
        return True
    while True:
        answer = messagebox.askyesnocancel("База не найдена",
            "Локальной базы нет. Скачать её с Яндекс Диска?\n\n"
            "Да — подключиться и восстановить.\nНет — создать новую локальную базу.\n"
            "Отмена — закрыть программу. Существующая база на Диске не будет заменена автоматически.", parent=parent)
        if answer is None:
            return False
        if not answer:
            return True
        try:
            try:
                load_token()
            except (ValueError, OSError):
                if not connect(parent):
                    continue
            with tempfile.TemporaryDirectory(dir=path.parent, prefix="restore-") as temp:
                target = Path(temp) / "restore.db"
                local_hash, meta = background(
                    parent,
                    "Скачивание базы",
                    lambda target=target: DiskAPI(load_token()).download(target),
                )
                os.replace(target, path)
                SyncEngine(path).record(local_hash, meta)
            return True
        except (ValueError, sqlite3.Error, OSError) as exc:
            messagebox.showerror("Восстановление", str(exc), parent=parent)


class CloudControls:
    def __init__(self, app, parent, path):
        self.app = app
        self.engine = SyncEngine(path)
        self.label = tk.StringVar(value="Яндекс Диск: проверка…")
        row = ttk.Frame(parent)
        row.pack(side="right")
        ttk.Button(row, text="Яндекс Диск", command=self.open).pack(anchor="e")
        ttk.Label(row, textvariable=self.label, wraplength=360).pack(anchor="e")
        self.engine.start()
        self.timer = app.after(500, self.poll)

    def saved(self):
        self.engine.request_sync()

    def poll(self):
        self.label.set(self.engine.status)
        self.timer = self.app.after(500, self.poll)

    def open(self):
        dialog = tk.Toplevel(self.app)
        dialog.withdraw()
        dialog.title("Яндекс Диск — Мои показания")
        dialog.transient(self.app)
        ttk.Label(dialog, text="MyDate / Показания / pokazaniya.db\n"
                  "Один файл с заменой после сохранения изменений.", padding=15).pack()
        ttk.Label(dialog, textvariable=self.label, wraplength=470, padding=10).pack()
        def reconnect():
            # Serialize credential changes with transfers.
            try:
                background(dialog, "Ожидание текущей отправки", self.engine.wait_until_idle)
                self.engine.stop()
                connect(dialog)
            except ValueError as exc:
                messagebox.showerror("Яндекс Диск", str(exc), parent=dialog)
            finally:
                self.engine.start()
        def send():
            if not messagebox.askyesno("Заменить базу на Диске?",
                "Текущая локальная база заменит файл на Диске, даже если он изменён с другого компьютера.", parent=dialog):
                return
            try:
                background(dialog, "Отправка базы", lambda: self.engine.once(force=True))
                self.engine.synced = True
                self.engine.status = "На Диске: " + self.engine.state().get("time", "")
            except ValueError as exc:
                messagebox.showerror("Яндекс Диск", str(exc), parent=dialog)
        def restore():
            if not messagebox.askyesno("Восстановить с Диска?",
                "Локальные записи будут заменены базой с Диска. Несохранённые поля форм будут очищены.", parent=dialog):
                return
            try:
                background(dialog, "Ожидание текущей отправки", self.engine.wait_until_idle)
                self.engine.stop()
                with tempfile.TemporaryDirectory(prefix="pokazaniya-") as temp:
                    target = Path(temp) / "restore.db"
                    local_hash, meta = background(dialog, "Скачивание базы", lambda: self.engine.api_factory().download(target))
                    source = sqlite3.connect(target)
                    try:
                        source.backup(self.app.db.connection)
                    finally:
                        source.close()
                    self.engine.record(local_hash, meta)
                for form in self.app.forms.values():
                    form.clear()
                self.app.refresh()
            except (ValueError, sqlite3.Error, OSError) as exc:
                messagebox.showerror("Восстановление", str(exc), parent=dialog)
            finally:
                self.engine.start()
        for text, command in (("Подключить / сменить аккаунт", reconnect),
                              ("Проверить и отправить сейчас", self.engine.request_sync),
                              ("Восстановить с Диска", restore),
                              ("Заменить облачную базу текущей", send)):
            ttk.Button(dialog, text=text, command=command).pack(fill="x", padx=15, pady=5)
        ttk.Label(dialog, text="При переносе: скопируйте всю папку вместе с yandex-disk.json,\n"
                  "выполните восстановление при отсутствии базы. Не редактируйте\n"
                  "на двух компьютерах: автоматического объединения нет.", padding=15).pack()
        show_centered(dialog)

    def close(self):
        abandon_upload = False
        if not self.engine.idle.is_set():
            try:
                background(self.app, "Завершение начатой отправки", self.engine.wait_until_idle)
            except ValueError as exc:
                if not messagebox.askyesno(
                    "Отправка не завершена",
                    str(exc) + "\n\nЗакрыть программу? Локальная база сохранена.",
                    parent=self.app,
                ):
                    return False
                abandon_upload = True
        self.engine.stop(wait=not abandon_upload)
        self.app.after_cancel(self.timer)
        return True
