"""Tkinter-интерфейс проверки и загрузки обновлений."""

from __future__ import annotations

import queue
import shutil
import threading
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk

import updates
from version import APP_VERSION


class UpdateControls:
    """Фоновая проверка GitHub Releases и диалог загрузки обновления."""

    def __init__(self, app: tk.Tk, parent: tk.Misc) -> None:
        self.app = app
        self.events: queue.Queue[tuple] = queue.Queue()
        self.checking = False
        self.downloading = False
        self.cancel_download = threading.Event()
        self.download_path: Path | None = None
        self.dialog: tk.Toplevel | None = None
        self.progress: ttk.Progressbar | None = None
        self.progress_text: tk.StringVar | None = None

        self.button = ttk.Button(
            parent,
            text=f"Версия {APP_VERSION}",
            command=lambda: self.check(manual=True),
        )
        self.button.pack(side="right", padx=(8, 0))
        self.poll_timer = self.app.after(150, self._poll)
        self.app.bind("<Destroy>", self._on_destroy, add="+")
        if updates.updates_supported():
            self.app.after(2500, self.check)

    def _on_destroy(self, event: tk.Event) -> None:
        if event.widget is self.app:
            self.close()

    def check(self, *, manual: bool = False) -> None:
        if self.checking or self.downloading:
            return
        if not updates.updates_supported():
            if manual:
                messagebox.showinfo(
                    "Обновления",
                    "Автообновление доступно в собранной Windows-версии.",
                    parent=self.app,
                )
            return
        self.checking = True
        self.button.state(["disabled"])
        self.button.configure(text="Проверяю…")

        def worker() -> None:
            try:
                result = updates.check_for_update()
                self.events.put(("checked", result, None, manual))
            except Exception as error:  # noqa: BLE001 — граница фонового потока
                self.events.put(("checked", None, error, manual))

        threading.Thread(target=worker, daemon=True, name="update-check").start()

    def _poll(self) -> None:
        try:
            while True:
                event = self.events.get_nowait()
                if event[0] == "checked":
                    self._check_finished(*event[1:])
                elif event[0] == "download-progress":
                    self._set_progress(event[1])
                elif event[0] == "downloaded":
                    self._download_finished(event[1], event[2])
        except queue.Empty:
            pass
        if self.app.winfo_exists():
            self.poll_timer = self.app.after(150, self._poll)

    def _check_finished(
        self,
        result: updates.UpdateInfo | None,
        error: Exception | None,
        manual: bool,
    ) -> None:
        self.checking = False
        self.button.state(["!disabled"])
        self.button.configure(text=f"Версия {APP_VERSION}")
        if error is not None:
            if manual:
                messagebox.showerror(
                    "Не удалось проверить обновления", str(error), parent=self.app
                )
            return
        if result is None:
            if manual:
                messagebox.showinfo(
                    "Обновления",
                    f"У вас уже установлена актуальная версия {APP_VERSION}.",
                    parent=self.app,
                )
            return
        text = (
            f"Доступна новая версия {result.version}.\n\n"
            "Скачать и установить её сейчас?"
        )
        if result.notes:
            notes = result.notes.strip()
            if len(notes) > 1200:
                notes = notes[:1197] + "…"
            text += f"\n\nЧто нового:\n{notes}"
        if messagebox.askyesno("Доступно обновление", text, parent=self.app):
            self._download(result)

    def _download(self, update: updates.UpdateInfo) -> None:
        self.downloading = True
        self.cancel_download.clear()
        self.download_path = updates.create_download_path(update.version)
        self.button.state(["disabled"])
        self._show_download_dialog(update.version)

        def worker() -> None:
            try:
                result = updates.download_update(
                    update,
                    self.download_path,
                    progress=lambda downloaded, total: self.events.put(
                        ("download-progress", round(downloaded * 100 / total))
                    ),
                    cancelled=self.cancel_download.is_set,
                )
                self.events.put(("downloaded", result, None))
            except Exception as error:  # noqa: BLE001 — граница фонового потока
                self.events.put(("downloaded", None, error))

        threading.Thread(target=worker, daemon=True, name="update-download").start()

    def _show_download_dialog(self, version: str) -> None:
        dialog = tk.Toplevel(self.app)
        dialog.withdraw()
        dialog.title("Обновление — Мои показания")
        dialog.transient(self.app)
        dialog.resizable(False, False)
        dialog.configure(background="#e9ebe8")
        dialog.protocol("WM_DELETE_WINDOW", self.cancel_download.set)
        if getattr(self.app, "_icons", None):
            dialog.iconphoto(False, *self.app._icons)

        outer = tk.Frame(dialog, background="#e9ebe8", padx=12, pady=12)
        outer.pack(fill="both", expand=True)
        card = tk.Frame(
            outer,
            background="#f7f7f4",
            highlightthickness=1,
            highlightbackground="#d0d6d1",
            padx=24,
            pady=20,
        )
        card.pack(fill="both", expand=True)
        tk.Label(
            card,
            text=f"Загрузка версии {version}",
            background="#f7f7f4",
            foreground="#263b33",
            font=("Segoe UI", 15, "bold"),
            anchor="w",
        ).pack(fill="x")
        self.progress_text = tk.StringVar(value="Загружено 0%")
        tk.Label(
            card,
            textvariable=self.progress_text,
            background="#f7f7f4",
            foreground="#40554d",
            font=("Segoe UI", 10),
            anchor="w",
        ).pack(fill="x", pady=(13, 7))
        style = ttk.Style(dialog)
        style.configure(
            "Download.Horizontal.TProgressbar",
            troughcolor="#d8e0da",
            background="#49645d",
            bordercolor="#d8e0da",
            lightcolor="#49645d",
            darkcolor="#49645d",
            thickness=15,
        )
        self.progress = ttk.Progressbar(
            card,
            maximum=100,
            style="Download.Horizontal.TProgressbar",
            mode="determinate",
        )
        self.progress.pack(fill="x")
        ttk.Button(card, text="Отмена", command=self.cancel_download.set).pack(
            anchor="e", pady=(14, 0)
        )

        dialog.update_idletasks()
        width, height = 500, 230
        x = max(0, (dialog.winfo_screenwidth() - width) // 2)
        y = max(0, (dialog.winfo_screenheight() - height) // 2)
        dialog.geometry(f"{width}x{height}+{x}+{y}")
        dialog.deiconify()
        dialog.grab_set()
        self.dialog = dialog

    def _set_progress(self, value: int) -> None:
        if self.progress is not None:
            self.progress["value"] = value
        if self.progress_text is not None:
            self.progress_text.set(f"Загружено {value}%")

    def _close_dialog(self) -> None:
        if self.dialog is not None:
            try:
                self.dialog.grab_release()
            except tk.TclError:
                pass
            self.dialog.destroy()
        self.dialog = None
        self.progress = None
        self.progress_text = None

    def _download_finished(self, result: Path | None, error: Exception | None) -> None:
        download_path = self.download_path
        self.download_path = None
        self.downloading = False
        self._close_dialog()
        self.button.state(["!disabled"])
        if error is not None:
            if download_path is not None:
                shutil.rmtree(download_path.parent, ignore_errors=True)
            if "отменена" not in str(error).lower():
                messagebox.showerror(
                    "Не удалось скачать обновление", str(error), parent=self.app
                )
            return
        if result is None:
            return
        if not messagebox.askyesno(
            "Всё готово к обновлению",
            "Программа закроется, установит новую версию и запустится снова.\n\n"
            "Продолжить?",
            parent=self.app,
        ):
            shutil.rmtree(result.parent, ignore_errors=True)
            return
        try:
            if not self.app.prepare_for_update():
                shutil.rmtree(result.parent, ignore_errors=True)
                return
            updates.launch_updater(result)
        except (OSError, updates.UpdateError) as error:
            shutil.rmtree(result.parent, ignore_errors=True)
            messagebox.showerror(
                "Не удалось запустить обновление", str(error), parent=self.app
            )
            return
        self.app.finish_for_update()

    def close(self) -> None:
        self.cancel_download.set()
        try:
            self.app.after_cancel(self.poll_timer)
        except (tk.TclError, ValueError):
            pass
