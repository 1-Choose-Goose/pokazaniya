"""Локальный журнал показаний и оплаты коммунальных услуг."""

from __future__ import annotations

import csv
import calendar
import ctypes
import sqlite3
import re
import sys
import tkinter as tk
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
import updates
from update_ui import UpdateControls
from yandex_sync import CloudControls, prepare_database


RESOURCE_DIR = Path(__file__).resolve().parent
APP_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else RESOURCE_DIR
DATA_DIR = APP_DIR / "data"
DB_PATH = DATA_DIR / "pokazaniya.db"
LEGACY_DB_PATH = APP_DIR / "pokazaniya.db"
ICON_PATH = RESOURCE_DIR / "app_icon.png"
ICON_ICO_PATH = RESOURCE_DIR / "app_icon.ico"
ICON_SIZES = (16, 32, 48, 256)
SERVICES = ("Вода", "Отопление", "Электроэнергия", "Газ", "ТКО", "Домофон", "Домашний интернет")
METERS = {"Вода": ("ХВС", "ГВС"), "Электроэнергия": ("День", "Ночь")}


def prepare_database_location(
    target: Path = DB_PATH,
    legacy: Path = LEGACY_DB_PATH,
) -> Path:
    """Создать data и перенести базу из старого расположения рядом с программой."""
    target.parent.mkdir(parents=True, exist_ok=True)
    if not target.exists() and legacy.is_file():
        legacy.replace(target)
    return target


def valid_date(raw: str, *, required: bool = True) -> str:
    value = raw.strip()
    if not value and not required:
        return ""
    try:
        return datetime.strptime(value, "%d.%m.%Y").strftime("%d.%m.%Y")
    except ValueError as exc:
        raise ValueError("Укажите дату в формате ДД.ММ.ГГГГ") from exc


def valid_payment_time(raw: str, *, required: bool = False) -> str:
    value = raw.strip()
    if not value and not required:
        return ""
    try:
        day_part, time_part = value.rsplit(None, 1)
        match = re.fullmatch(r"(\d{1,2})[:.,](\d{2})", time_part)
        if match is None:
            raise ValueError
        hours, minutes = map(int, match.groups())
        parsed_day = datetime.strptime(day_part, "%d.%m.%Y")
        return parsed_day.replace(
            hour=hours, minute=minutes
        ).strftime("%d.%m.%Y %H:%M")
    except ValueError as exc:
        raise ValueError(
            "Укажите дату и время с чека: ДД.ММ.ГГГГ и ЧЧ:ММ "
            "(можно также ЧЧ.ММ или ЧЧ,ММ)"
        ) from exc


def number(raw: str, label: str, *, required: bool = True) -> Decimal | None:
    value = raw.strip().replace(" ", "").replace("\u00a0", "").replace(",", ".")
    if not value and not required:
        return None
    try:
        result = Decimal(value)
    except InvalidOperation as exc:
        raise ValueError(f"Укажите число для поля «{label}»") from exc
    if not result.is_finite() or result < 0:
        raise ValueError(f"Поле «{label}» должно быть неотрицательным числом")
    return result


def money(value: Decimal) -> str:
    return f"{value:,.2f}".replace(",", " ").replace(".", ",") + " ₽"


def reading(value: Decimal | None) -> str:
    if value is None:
        return "—"
    return format(value, "f").rstrip("0").rstrip(".") if "." in format(value, "f") else str(value)


def export_excel_workbook(path: Path, db: "Database") -> None:
    """Экспорт полного журнала и настроек счётчиков в оформленный XLSX."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.workbook.properties import CalcProperties

    entries = sorted(
        db.all(), key=lambda item: (
            datetime.strptime(item.period, "%d.%m.%Y"), item.id or 0,
        ),
    )
    created = {
        row["id"]: row["created_at"]
        for row in db.connection.execute("SELECT id, created_at FROM entries")
    }
    settings = {
        row["key"]: row["value"]
        for row in db.connection.execute("SELECT key, value FROM settings")
    }
    book = Workbook()
    summary = book.active
    summary.title = "Сводка"
    journal = book.create_sheet("Все записи")
    meters = book.create_sheet("Счётчики")
    book.calculation = CalcProperties(calcMode="auto", fullCalcOnLoad=True)
    ink, muted, accent = "293431", "66756D", "49645D"
    pale, stripe, line = "E9EFEB", "F6F8F6", "D8E0DA"
    red, red_fill = "A33D3D", "F9E8E6"
    green, green_fill = "376A4D", "E8F3EA"
    currency_format = '#,##0.00 "₽"'
    number_format = '#,##0.###'
    date_format = "dd.mm.yyyy"
    datetime_format = "dd.mm.yyyy hh:mm"

    def text_cell(cell: object, value: str) -> None:
        cell.value = value
        cell.data_type = "s"  # Номера с ведущими нулями и пользовательский текст — не формулы.

    def date_value(raw: str) -> date | datetime | None:
        if not raw:
            return None
        for pattern in ("%d.%m.%Y %H:%M", "%d.%m.%Y", "%Y-%m-%d %H:%M:%S"):
            try:
                parsed = datetime.strptime(raw, pattern)
                return parsed if "%H" in pattern else parsed.date()
            except ValueError:
                continue
        return None

    def setup(sheet: object, title: str, subtitle: str, widths: list[int]) -> None:
        sheet.sheet_view.showGridLines = False
        sheet.sheet_properties.pageSetUpPr.fitToPage = True
        sheet.sheet_properties.tabColor = accent
        text_cell(sheet["A1"], title)
        sheet["A1"].font = Font(name="Arial", size=16, bold=True, color=ink)
        sheet.row_dimensions[1].height = 30
        text_cell(sheet["A2"], subtitle)
        sheet["A2"].font = Font(name="Arial", size=10, italic=True, color=muted)
        sheet.row_dimensions[2].height = 22
        for index, width in enumerate(widths, 1):
            sheet.column_dimensions[get_column_letter(index)].width = width
        sheet.freeze_panes = "C5" if sheet == journal else "A5"

    def header(sheet: object, labels: tuple[str, ...]) -> None:
        for col, label in enumerate(labels, 1):
            cell = sheet.cell(4, col, label)
            cell.fill = PatternFill("solid", fgColor=accent)
            cell.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = Border(right=Side(style="hair", color="FFFFFF"))
        sheet.row_dimensions[4].height = 34

    def body_row(sheet: object, row_index: int, count: int) -> None:
        fill = PatternFill("solid", fgColor=stripe if row_index % 2 else "FFFFFF")
        for col in range(1, count + 1):
            cell = sheet.cell(row_index, col)
            cell.fill = fill
            cell.font = Font(name="Arial", size=10, color=ink)
            cell.alignment = Alignment(vertical="center")
            cell.border = Border(bottom=Side(style="hair", color=line))
        sheet.row_dimensions[row_index].height = 23

    setup(journal, "Все записи", "Полная история показаний и оплат", [8, 9, 9, 23, 17, 20, 17, 20, 17, 17, 17, 20, 17, 35, 21])
    header(journal, (
        "ID", "Год", "Месяц", "Услуга", "Дата показаний / период",
        "Счётчик 1", "Показание 1", "Счётчик 2", "Показание 2",
        "Общее воды", "Сумма", "Дата и время оплаты", "Статус",
        "Заметка", "Создано",
    ))
    for row_index, item in enumerate(entries, 5):
        body_row(journal, row_index, 15)
        period = datetime.strptime(item.period, "%d.%m.%Y").date()
        payment = date_value(item.payment_date)
        values = (
            item.id, period.year, period.month, item.service, period,
            item.meter_1, float(item.value_1) if item.value_1 is not None else None,
            item.meter_2, float(item.value_2) if item.value_2 is not None else None,
            float(item.total_water) if item.total_water is not None else None,
            float(item.amount) if item.amount is not None else None,
            payment or item.payment_date,
            "Оплачено" if item.paid else "Не оплачено", item.note,
            date_value(created.get(item.id, "")) or created.get(item.id, ""),
        )
        for col, value in enumerate(values, 1):
            cell = journal.cell(row_index, col)
            if col in (4, 6, 8, 13, 14) and isinstance(value, str):
                text_cell(cell, value)
            else:
                cell.value = value
        for col in (6, 8):
            journal.cell(row_index, col).number_format = "@"
        for col in (7, 9, 10):
            journal.cell(row_index, col).number_format = number_format
        journal.cell(row_index, 11).number_format = currency_format
        journal.cell(row_index, 5).number_format = date_format
        journal.cell(row_index, 12).number_format = datetime_format
        journal.cell(row_index, 15).number_format = datetime_format
        status_cell = journal.cell(row_index, 13)
        status_cell.fill = PatternFill("solid", fgColor=green_fill if item.paid else red_fill)
        status_cell.font = Font(name="Arial", size=10, bold=True, color=green if item.paid else red)
    last_journal_row = max(5, len(entries) + 4)
    journal.auto_filter.ref = f"A4:O{last_journal_row}"
    journal.print_options.horizontalCentered = True
    journal.sheet_properties.pageSetUpPr.fitToPage = True

    setup(summary, "Сводка по месяцам", "Оплаченные суммы по услугам и неоплаченный остаток", [12, 17, 17, 17, 20, 17, 15, 17, 25, 19, 19])
    header(summary, ("Год", "Месяц", *SERVICES, "Оплачено", "Не оплачено"))
    years = sorted({2025, date.today().year} | {
        datetime.strptime(item.period, "%d.%m.%Y").year for item in entries
    })
    month_names = (
        "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
        "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь",
    )
    data_end = max(5, len(entries) + 4)
    ref = "'Все записи'!"
    year_range = f"{ref}$B$5:$B${data_end}"
    month_range = f"{ref}$C$5:$C${data_end}"
    service_range = f"{ref}$D$5:$D${data_end}"
    amount_range = f"{ref}$K$5:$K${data_end}"
    status_range = f"{ref}$M$5:$M${data_end}"
    row_index = 5
    year_total_rows = []
    for year in years:
        first_month_row = row_index
        for month, name in enumerate(month_names, 1):
            body_row(summary, row_index, 11)
            summary.cell(row_index, 1, year)
            text_cell(summary.cell(row_index, 2), name)
            for col in range(3, 10):
                letter = get_column_letter(col)
                summary.cell(row_index, col, (
                    f'=SUMIFS({amount_range},{year_range},$A{row_index},'
                    f'{month_range},{month},{service_range},{letter}$4,'
                    f'{status_range},"Оплачено")'
                ))
            summary.cell(row_index, 10, f"=SUM(C{row_index}:I{row_index})")
            summary.cell(row_index, 11, (
                f'=SUMIFS({amount_range},{year_range},$A{row_index},'
                f'{month_range},{month},{status_range},"Не оплачено")'
            ))
            for col in range(3, 12):
                summary.cell(row_index, col).number_format = currency_format
            row_index += 1
        body_row(summary, row_index, 11)
        text_cell(summary.cell(row_index, 2), f"Итого {year}")
        for col in range(3, 12):
            letter = get_column_letter(col)
            summary.cell(row_index, col, f"=SUM({letter}{first_month_row}:{letter}{row_index - 1})")
            summary.cell(row_index, col).number_format = currency_format
        for cell in summary[row_index][:11]:
            cell.fill = PatternFill("solid", fgColor=pale)
            cell.font = Font(name="Arial", size=10, bold=True, color=ink)
        year_total_rows.append(row_index)
        row_index += 1
    summary["H2"] = "Всего оплачено"
    summary["H2"].font = Font(name="Arial", size=10, bold=True, color=ink)
    summary["J2"] = "=" + "+".join(f"J{row}" for row in year_total_rows)
    summary["J2"].number_format = currency_format
    summary["J2"].font = Font(name="Arial", size=12, bold=True, color=accent)
    summary.auto_filter.ref = f"A4:K{row_index - 1}"

    setup(meters, "Счётчики", "Номера и даты из настроек программы", [25, 22, 26, 22])
    header(meters, ("Услуга", "Счётчик", "Номер", "Дата"))
    meter_rows = (
        ("Вода", "ХВС", "water_cold_meter", "water_cold_date"),
        ("Вода", "ГВС", "water_hot_meter", "water_hot_date"),
        ("Электроэнергия", "День / ночь", "electric_meter", "electric_date"),
        ("Газ", "Газ", "gas_meter", "gas_date"),
    )
    known_keys = set()
    for row_index, (service, label, number_key, date_key) in enumerate(meter_rows, 5):
        body_row(meters, row_index, 4)
        text_cell(meters.cell(row_index, 1), service)
        text_cell(meters.cell(row_index, 2), label)
        text_cell(meters.cell(row_index, 3), settings.get(number_key, ""))
        meters.cell(row_index, 3).number_format = "@"
        date_text = settings.get(date_key, "")
        meters.cell(row_index, 4).value = date_value(date_text) or date_text
        meters.cell(row_index, 4).number_format = date_format
        known_keys.update((number_key, date_key))
    for row_index, (key, value) in enumerate(
        sorted((key, value) for key, value in settings.items() if key not in known_keys),
        9,
    ):
        body_row(meters, row_index, 4)
        text_cell(meters.cell(row_index, 1), "Дополнительно")
        text_cell(meters.cell(row_index, 2), key)
        text_cell(meters.cell(row_index, 3), value)
    book.save(path)


def suggested_reading(
    entries: list["Entry"],
    service: str,
    field: str,
    target_date: str,
    *,
    meter_number: str = "",
    exclude_id: int | None = None,
) -> Decimal | None:
    """Прогноз по среднему месячному расходу предыдущих показаний."""
    try:
        target = datetime.strptime(target_date, "%d.%m.%Y")
    except ValueError:
        return None
    target_month = target.year * 12 + target.month
    monthly: dict[int, tuple[datetime, Decimal]] = {}
    meter_field = "meter_2" if service == "Вода" and field == "value_2" else "meter_1"
    for item in entries:
        if item.service != service or item.id == exclude_id:
            continue
        value = getattr(item, field)
        if value is None:
            continue
        item_meter = getattr(item, meter_field)
        if meter_number and item_meter and item_meter != meter_number:
            continue
        item_date = datetime.strptime(item.period, "%d.%m.%Y")
        month = item_date.year * 12 + item_date.month
        if item_date >= target or month >= target_month:
            continue
        if month not in monthly or item_date >= monthly[month][0]:
            monthly[month] = (item_date, value)
    points = sorted((month, value) for month, (_day, value) in monthly.items())
    if len(points) < 2:
        return None
    rates = []
    for (first_month, first_value), (second_month, second_value) in zip(
        points, points[1:]
    ):
        delta = second_value - first_value
        if delta >= 0:
            rates.append(delta / Decimal(second_month - first_month))
    if not rates:
        return None
    average = sum(rates, Decimal("0")) / Decimal(len(rates))
    forecast = points[-1][1] + average * Decimal(target_month - points[-1][0])
    return forecast.quantize(Decimal("0.001"), rounding=ROUND_HALF_UP)


def center_window(window: tk.Misc, width: int | None = None, height: int | None = None) -> None:
    window.update_idletasks()
    width = width or window.winfo_reqwidth()
    height = height or window.winfo_reqheight()
    x = max(0, (window.winfo_screenwidth() - width) // 2)
    y = max(0, (window.winfo_screenheight() - height) // 2)
    window.geometry(f"{width}x{height}+{x}+{y}")


class ModernScrollbar(tk.Canvas):
    """Тонкая прокрутка без кнопок со стрелками."""

    def __init__(self, parent: tk.Misc, command):
        super().__init__(
            parent, width=12, height=60, highlightthickness=0, borderwidth=0,
            background="#e9ebe8", cursor="hand2",
        )
        self.command = command
        self.first = 0.0
        self.last = 1.0
        self.drag_offset: float | None = None
        self.bind("<Configure>", lambda _event: self._draw())
        self.bind("<Button-1>", self._press)
        self.bind("<B1-Motion>", self._drag)
        self.bind("<ButtonRelease-1>", lambda _event: setattr(self, "drag_offset", None))
        self.bind("<Enter>", lambda _event: self._draw(hover=True))
        self.bind("<Leave>", lambda _event: self._draw())

    def set(self, first: str, last: str) -> None:
        self.first = float(first)
        self.last = float(last)
        self._draw()

    def _thumb(self) -> tuple[float, float]:
        height = max(1, self.winfo_height())
        top = height * self.first
        bottom = height * self.last
        if bottom - top < 28:
            bottom = min(height, top + 28)
            top = max(0, bottom - 28)
        return top, bottom

    def _draw(self, hover: bool = False) -> None:
        self.delete("all")
        height = self.winfo_height()
        self.create_rectangle(5, 0, 7, height, fill="#dce2dd", outline="")
        if self.first <= 0 and self.last >= 1:
            return
        top, bottom = self._thumb()
        color = "#778d82" if hover or self.drag_offset is not None else "#96a99e"
        self.create_rectangle(3, top + 4, 9, bottom - 4, fill=color, outline="")
        self.create_oval(3, top, 9, top + 8, fill=color, outline="")
        self.create_oval(3, bottom - 8, 9, bottom, fill=color, outline="")

    def _press(self, event: tk.Event) -> None:
        top, bottom = self._thumb()
        if top <= event.y <= bottom:
            self.drag_offset = event.y - top
        else:
            self.drag_offset = (bottom - top) / 2
            self._move_to(event.y)
        self._draw(hover=True)

    def _drag(self, event: tk.Event) -> None:
        if self.drag_offset is not None:
            self._move_to(event.y)

    def _move_to(self, y: float) -> None:
        height = max(1, self.winfo_height())
        visible_fraction = max(0.0, min(1.0, self.last - self.first))
        offset = self.drag_offset or 0.0
        position = max(0.0, min(1.0 - visible_fraction, (y - offset) / height))
        self.command("moveto", position)


@dataclass
class Entry:
    service: str
    period: str
    amount: Decimal | None
    payment_date: str = ""
    paid: bool = False
    meter_1: str = ""
    meter_2: str = ""
    value_1: Decimal | None = None
    value_2: Decimal | None = None
    note: str = ""
    id: int | None = None

    @property
    def total_water(self) -> Decimal | None:
        if self.service == "Вода" and self.value_1 is not None and self.value_2 is not None:
            return self.value_1 + self.value_2
        return None


class DuplicateEntry(ValueError):
    """A record that would duplicate an existing history row."""


class Database:
    def __init__(self, path: Path = DB_PATH):
        self.connection = sqlite3.connect(path)
        self.connection.row_factory = sqlite3.Row
        self.connection.execute(
            """CREATE TABLE IF NOT EXISTS entries (
                id INTEGER PRIMARY KEY,
                service TEXT NOT NULL,
                period TEXT NOT NULL,
                amount TEXT,
                payment_date TEXT NOT NULL DEFAULT '',
                paid INTEGER NOT NULL DEFAULT 0,
                meter_1 TEXT NOT NULL DEFAULT '',
                meter_2 TEXT NOT NULL DEFAULT '',
                value_1 TEXT,
                value_2 TEXT,
                note TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            )"""
        )
        self.connection.execute(
            "CREATE INDEX IF NOT EXISTS idx_entries_service_period ON entries(service, period)"
        )
        self.connection.execute(
            "CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT NOT NULL)"
        )
        self._migrate_legacy()
        for key, service, field in (
            ("water_cold_meter", "Вода", "meter_1"),
            ("water_hot_meter", "Вода", "meter_2"),
            ("electric_meter", "Электроэнергия", "meter_1"),
            ("gas_meter", "Газ", "meter_1"),
        ):
            if not self.setting(key):
                row = self.connection.execute(
                    f"SELECT {field} FROM entries WHERE service=? AND {field}<>'' "
                    "ORDER BY id DESC LIMIT 1", (service,)
                ).fetchone()
                if row:
                    self.set_setting(key, row[field], commit=False)
        self.connection.commit()

    def setting(self, key: str) -> str:
        row = self.connection.execute(
            "SELECT value FROM settings WHERE key=?", (key,)
        ).fetchone()
        return row["value"] if row else ""

    def set_setting(self, key: str, value: str, *, commit: bool = True) -> None:
        self.connection.execute(
            "INSERT INTO settings(key,value) VALUES (?,?) "
            "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
            (key, value.strip()),
        )
        if commit:
            self.connection.commit()

    def _migrate_legacy(self) -> None:
        old = self.connection.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name='readings'"
        ).fetchone()
        if not old:
            return
        migrated = self.connection.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name='legacy_imports'"
        ).fetchone()
        if not migrated:
            self.connection.execute(
                "CREATE TABLE legacy_imports (old_id INTEGER PRIMARY KEY)"
            )
        columns = {
            row["name"] for row in self.connection.execute("PRAGMA table_info(readings)")
        }
        for row in self.connection.execute("SELECT * FROM readings ORDER BY id").fetchall():
            if self.connection.execute(
                "SELECT 1 FROM legacy_imports WHERE old_id=?", (row["id"],)
            ).fetchone():
                continue
            paid = bool(row["is_paid"]) if "is_paid" in columns else False
            old_date = row["reading_date"]
            amounts = (
                ("Вода", "payment_water", row["cold_water"], row["hot_water"]),
                ("Электроэнергия", "payment_electricity", row["electricity_t1"], row["electricity_t2"]),
                ("Газ", "payment_gas", None, None),
                ("Прочее (архив)", "payment_other", None, None),
            )
            for service, column, first, second in amounts:
                amount = Decimal(row[column])
                if service in METERS or amount != 0:
                    self.add(Entry(
                        service=service, period=old_date, amount=amount,
                        payment_date=old_date if paid and amount else "",
                        paid=paid and amount != 0,
                        value_1=Decimal(first) if first is not None else None,
                        value_2=Decimal(second) if second is not None else None,
                        note=row["notes"] if service == "Вода" else "",
                    ), commit=False)
            self.connection.execute(
                "INSERT INTO legacy_imports(old_id) VALUES (?)", (row["id"],)
            )

    @staticmethod
    def _data(entry: Entry) -> tuple:
        return (
            entry.service, entry.period,
            str(entry.amount) if entry.amount is not None else None,
            entry.payment_date, int(entry.paid), entry.meter_1.strip(),
            entry.meter_2.strip(),
            str(entry.value_1) if entry.value_1 is not None else None,
            str(entry.value_2) if entry.value_2 is not None else None,
            entry.note.strip(),
        )

    def add(self, entry: Entry, *, commit: bool = True) -> int:
        self._ensure_not_duplicate(entry)
        cursor = self.connection.execute(
            """INSERT INTO entries
            (service, period, amount, payment_date, paid, meter_1, meter_2, value_1, value_2, note)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            self._data(entry),
        )
        if commit:
            self.connection.commit()
        return int(cursor.lastrowid)

    def update(self, entry: Entry, *, commit: bool = True) -> None:
        if entry.id is None:
            raise ValueError("Не выбрана запись")
        self._ensure_not_duplicate(entry, exclude_id=entry.id)
        self.connection.execute(
            """UPDATE entries SET service=?, period=?, amount=?, payment_date=?,
            paid=?, meter_1=?, meter_2=?, value_1=?, value_2=?, note=? WHERE id=?""",
            (*self._data(entry), entry.id),
        )
        if commit:
            self.connection.commit()

    def _ensure_not_duplicate(self, entry: Entry, *, exclude_id: int | None = None) -> None:
        params: list[object] = [entry.service, entry.period[3:]]
        query = "SELECT id FROM entries WHERE service=? AND substr(period, 4)=?"
        if exclude_id is not None:
            query += " AND id<>?"
            params.append(exclude_id)
        same_month = self.connection.execute(query + " LIMIT 1", params).fetchone()
        if entry.service == "Газ" and same_month:
            raise DuplicateEntry(
                "Запись «Газ» за этот месяц уже существует. "
                "Выберите её в истории и нажмите «Изменить»."
            )

        data = self._data(entry)
        query = """SELECT id FROM entries WHERE
            service IS ? AND period IS ? AND amount IS ? AND payment_date IS ? AND
            paid IS ? AND meter_1 IS ? AND meter_2 IS ? AND value_1 IS ? AND
            value_2 IS ? AND note IS ?"""
        params = list(data)
        if exclude_id is not None:
            query += " AND id<>?"
            params.append(exclude_id)
        if self.connection.execute(query + " LIMIT 1", params).fetchone():
            raise DuplicateEntry("Такая запись уже сохранена.")

    def delete(self, entry_id: int) -> None:
        self.connection.execute("DELETE FROM entries WHERE id=?", (entry_id,))
        self.connection.commit()

    def all(self, service: str | None = None) -> list[Entry]:
        query = "SELECT * FROM entries"
        params: tuple = ()
        if service:
            query += " WHERE service=?"
            params = (service,)
        rows = self.connection.execute(query, params).fetchall()
        result = [self._from_row(row) for row in rows]
        return sorted(
            result,
            key=lambda item: (datetime.strptime(item.period, "%d.%m.%Y"), item.id or 0),
            reverse=True,
        )

    def get(self, entry_id: int) -> Entry | None:
        row = self.connection.execute("SELECT * FROM entries WHERE id=?", (entry_id,)).fetchone()
        return self._from_row(row) if row else None

    @staticmethod
    def _from_row(row: sqlite3.Row) -> Entry:
        return Entry(
            id=row["id"], service=row["service"], period=row["period"],
            amount=Decimal(row["amount"]) if row["amount"] is not None else None,
            payment_date=row["payment_date"], paid=bool(row["paid"]),
            meter_1=row["meter_1"], meter_2=row["meter_2"],
            value_1=Decimal(row["value_1"]) if row["value_1"] is not None else None,
            value_2=Decimal(row["value_2"]) if row["value_2"] is not None else None,
            note=row["note"],
        )

    def close(self) -> None:
        self.connection.close()


class ServicePage(ttk.Frame):
    def __init__(self, master: ttk.Notebook, app: "App", service: str):
        super().__init__(master, padding=18)
        self.app = app
        self.service = service
        self.edit_id: int | None = None
        self.fields: dict[str, tk.StringVar] = {}
        self.paid = tk.BooleanVar()
        self.columnconfigure(0, weight=1)
        form = ttk.Frame(self)
        form.grid(row=0, column=0, sticky="ew")
        form.columnconfigure(1, weight=1)

        row = 0
        self._field(form, row, "period", "Дата показаний" if service in METERS else "Расчётный период")
        row += 1
        if service in METERS:
            self._field(
                form, row, "meter_1",
                f"Номер счётчика {METERS[service][0]}" if service == "Вода" else "Номер счётчика",
            )
            row += 1
            self._field(form, row, "value_1", f"Показание {METERS[service][0]}")
            row += 1
            if service == "Вода":
                self._field(form, row, "meter_2", "Номер счётчика ГВС")
                row += 1
            self._field(form, row, "value_2", f"Показание {METERS[service][1]}")
            row += 1
            if service == "Вода":
                self.total_label = ttk.Label(form, text="Общее воды: —")
                self.total_label.grid(row=row, column=0, columnspan=2, sticky="w", pady=(4, 10))
                for name in ("value_1", "value_2"):
                    self.fields[name].trace_add("write", self._update_total)
                row += 1
        self._field(form, row, "amount", "Сумма, ₽")
        row += 1
        self._field(form, row, "payment_date", "Дата оплаты")
        row += 1
        ttk.Checkbutton(form, text="Оплачено", variable=self.paid, command=self._on_paid).grid(
            row=row, column=0, columnspan=2, sticky="w", pady=8
        )
        row += 1
        self._field(form, row, "note", "Заметка")
        row += 1
        actions = ttk.Frame(form)
        actions.grid(row=row, column=0, columnspan=2, sticky="w", pady=(12, 0))
        self.save_button = ttk.Button(actions, text="Сохранить", command=self.save)
        self.save_button.pack(side="left")
        ttk.Button(actions, text="Очистить", command=self.clear).pack(side="left", padx=8)

        ttk.Separator(self, orient="horizontal").grid(row=1, column=0, sticky="ew", pady=18)
        history_header = ttk.Frame(self)
        history_header.grid(row=2, column=0, sticky="ew", pady=(0, 8))
        ttk.Label(history_header, text="Год:").pack(side="left")
        self.year_var = tk.StringVar()
        self.year_box = ttk.Combobox(
            history_header, textvariable=self.year_var, state="readonly", width=9
        )
        self.year_box.pack(side="left", padx=(8, 18))
        self.year_box.bind("<<ComboboxSelected>>", lambda _event: self.refresh(self.app.db.all()))
        self.summary = ttk.Label(history_header, text="")
        self.summary.pack(side="left")
        table = ttk.Frame(self)
        table.grid(row=3, column=0, sticky="nsew")
        self.rowconfigure(3, weight=1)
        self.tree = ttk.Treeview(
            table, columns=("period", "reading", "change", "amount", "date", "status"),
            show="tree headings", selectmode="browse", height=8,
        )
        self.tree.heading("#0", text="Год / месяц")
        self.tree.column("#0", width=125, minwidth=105)
        for key, title, width in (
            ("period", "Дата", 105),
            ("reading", "Показания", 310),
            ("change", "Изменение", 185),
            ("amount", "Сумма", 120),
            ("date", "Дата оплаты", 110),
            ("status", "Статус", 110),
        ):
            self.tree.heading(key, text=title)
            self.tree.column(key, width=width, minwidth=90, anchor="w")
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar = ttk.Scrollbar(table, command=self.tree.yview)
        scrollbar.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.bind("<Double-1>", lambda _event: self.edit())
        history_actions = ttk.Frame(self)
        history_actions.grid(row=4, column=0, sticky="w", pady=(8, 0))
        ttk.Button(history_actions, text="Изменить", command=self.edit).pack(side="left")
        ttk.Button(history_actions, text="Удалить", command=self.delete).pack(side="left", padx=8)
        self.clear()

    def _field(self, parent: ttk.Frame, row: int, key: str, title: str) -> None:
        var = tk.StringVar()
        self.fields[key] = var
        ttk.Label(parent, text=title).grid(row=row, column=0, sticky="w", padx=(0, 16), pady=4)
        ttk.Entry(parent, textvariable=var, width=38).grid(
            row=row, column=1, sticky="ew", pady=4
        )

    def _update_total(self, *_args: object) -> None:
        try:
            first = number(self.fields["value_1"].get(), "ХВС")
            second = number(self.fields["value_2"].get(), "ГВС")
            self.total_label.configure(text=f"Общее воды: {reading(first + second)} м³")
        except ValueError:
            self.total_label.configure(text="Общее воды: —")

    def _on_paid(self) -> None:
        if self.paid.get() and not self.fields["payment_date"].get().strip():
            self.fields["payment_date"].set(date.today().strftime("%d.%m.%Y"))

    def _entry(self) -> Entry:
        amount = number(self.fields["amount"].get(), "Сумма", required=False)
        payment_date = valid_date(self.fields["payment_date"].get(), required=False)
        paid = self.paid.get()
        if paid and (amount is None or payment_date == ""):
            raise ValueError("Для оплаченной записи укажите сумму и дату оплаты")
        if payment_date and not paid:
            raise ValueError("При указанной дате оплаты отметьте «Оплачено»")
        meter_1 = meter_2 = ""
        value_1 = value_2 = None
        if self.service in METERS:
            meter_1 = self.fields["meter_1"].get().strip()
            if self.service == "Вода":
                meter_2 = self.fields["meter_2"].get().strip()
            value_1 = number(self.fields["value_1"].get(), METERS[self.service][0])
            value_2 = number(self.fields["value_2"].get(), METERS[self.service][1])
        return Entry(
            id=self.edit_id, service=self.service,
            period=valid_date(self.fields["period"].get()), amount=amount,
            payment_date=payment_date, paid=paid,
            meter_1=meter_1, meter_2=meter_2,
            value_1=value_1, value_2=value_2,
            note=self.fields["note"].get(),
        )

    def save(self) -> None:
        try:
            entry = self._entry()
            if self.edit_id is None:
                self.app.db.add(entry)
            else:
                self.app.db.update(entry)
        except (ValueError, sqlite3.Error) as exc:
            messagebox.showerror("Проверьте данные", str(exc), parent=self.app)
            return
        self.clear()
        self.app.refresh()
        if self.app.cloud is not None:
            self.app.cloud.saved()

    def clear(self) -> None:
        self.edit_id = None
        for var in self.fields.values():
            var.set("")
        self.fields["period"].set(date.today().strftime("%d.%m.%Y"))
        self.paid.set(False)
        self.save_button.configure(text="Сохранить")

    def _selected(self) -> Entry | None:
        selection = self.tree.selection()
        if not selection or not selection[0].startswith("entry:"):
            messagebox.showinfo("Выберите запись", "Выберите запись внутри месяца.", parent=self.app)
            return None
        return self.app.db.get(int(selection[0].split(":", 1)[1]))

    def edit(self) -> None:
        entry = self._selected()
        if entry is None:
            return
        self.edit_id = entry.id
        for key in ("period", "amount", "payment_date", "meter_1", "meter_2",
                    "value_1", "value_2", "note"):
            if key in self.fields:
                value = getattr(entry, key)
                self.fields[key].set(str(value) if value is not None else "")
        self.paid.set(entry.paid)
        self.save_button.configure(text="Сохранить изменения")

    def delete(self) -> None:
        entry = self._selected()
        if entry and messagebox.askyesno(
            "Удаление", "Удалить выбранную запись?", parent=self.app
        ):
            self.app.db.delete(entry.id)
            if self.edit_id == entry.id:
                self.clear()
            self.app.refresh()

    def refresh(self, entries: list[Entry]) -> None:
        self.tree.delete(*self.tree.get_children())
        selected = [item for item in entries if item.service == self.service]
        years = sorted(
            {datetime.strptime(item.period, "%d.%m.%Y").year for item in selected}
            | {date.today().year},
            reverse=True,
        )
        options = [str(year) for year in years]
        self.year_box.configure(values=options)
        if self.year_var.get() not in options:
            self.year_var.set(options[0])
        selected_year = int(self.year_var.get())
        visible = [
            item for item in selected
            if datetime.strptime(item.period, "%d.%m.%Y").year == selected_year
        ]
        paid_sum = sum(
            (item.amount for item in visible if item.paid and item.amount is not None),
            Decimal("0"),
        )
        self.summary.configure(text=f"Оплачено за {selected_year}: {money(paid_sum)}")
        previous_by_id: dict[int, Entry | None] = {}
        previous: Entry | None = None
        for item in reversed(selected):
            previous_by_id[item.id] = previous
            previous = item
        month_names = (
            "январь", "февраль", "март", "апрель", "май", "июнь",
            "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь",
        )
        year_key = f"year:{selected_year}"
        self.tree.insert("", "end", iid=year_key, text=str(selected_year), open=True)
        for month in range(1, 13):
            self.tree.insert(
                year_key, "end", iid=f"month:{selected_year}-{month:02d}",
                text=month_names[month - 1].capitalize(), open=True,
            )
        for item in visible:
            entry_date = datetime.strptime(item.period, "%d.%m.%Y")
            month_key = f"month:{entry_date.year}-{entry_date.month:02d}"
            details = ""
            change = ""
            if item.service in METERS:
                first, second = METERS[item.service]
                if item.service == "Вода":
                    details = (
                        f"{first} №{item.meter_1 or '—'}: {reading(item.value_1)}   "
                        f"{second} №{item.meter_2 or '—'}: {reading(item.value_2)}"
                    )
                else:
                    details = (
                        f"№{item.meter_1 or '—'}   "
                        f"{first}: {reading(item.value_1)}   "
                        f"{second}: {reading(item.value_2)}"
                    )
                if item.service == "Вода":
                    details += f"   Всего: {reading(item.total_water)}"
                prior = previous_by_id[item.id]
                if prior and prior.value_1 is not None and prior.value_2 is not None:
                    delta_1 = item.value_1 - prior.value_1
                    delta_2 = item.value_2 - prior.value_2
                    change = f"{first} {reading(delta_1)}; {second} {reading(delta_2)}"
                    if item.service == "Вода":
                        change += f"; всего {reading(delta_1 + delta_2)}"
                else:
                    change = "Первая запись"
            self.tree.insert(month_key, "end", iid=f"entry:{item.id}", values=(
                item.period, details, change,
                money(item.amount) if item.amount is not None else "—",
                item.payment_date or "—", "Оплачено" if item.paid else "Не оплачено",
            ))


class CompactForm(ttk.Frame):
    """Одна услуга: короткая форма, доступная прямо на главном экране."""

    def __init__(self, parent: ttk.Frame, app: "App", service: str):
        super().__init__(parent, style="Card.TFrame", padding=(9, 7))
        self.app = app
        self.service = service
        self.edit_id: int | None = None
        self.fields: dict[str, tk.StringVar] = {}
        self.suggestion_labels: dict[str, tk.Label] = {}
        self.suggestion_values: dict[str, Decimal | None] = {}
        self.meter_override: tuple[str, str] | None = None
        self.meter_labels: dict[str, ttk.Label] = {}
        self.paid = tk.BooleanVar()
        self.columnconfigure(1, weight=1)
        ttk.Label(self, text=service, style="CardTitle.TLabel").grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 4)
        )
        self.meter_meta = ttk.Label(self, text=" ", style="CardMeta.TLabel")
        self.meter_meta.grid(row=1, column=0, columnspan=2, sticky="w", pady=(0, 2))
        self._field("period", "Дата показаний" if service in METERS else "Период", 2)
        if service == "Вода":
            self._field("value_1", "ХВС", 3, meter_label=True)
            self._field("value_2", "ГВС", 4, meter_label=True)
            self.total_label = ttk.Label(self, text="Общее воды: —", style="Card.TLabel")
            self.total_label.grid(row=5, column=0, columnspan=2, sticky="w", pady=(2, 0))
            self.fields["value_1"].trace_add("write", self._water_total)
            self.fields["value_2"].trace_add("write", self._water_total)
        elif service == "Электроэнергия":
            self._field("value_1", "День", 3, meter_label=True)
            self._field("value_2", "Ночь", 4, meter_label=True)
        elif service == "Газ":
            self._field("value_1", "Показание", 3, meter_label=True)
        if service in METERS:
            self.rowconfigure(5, minsize=20)
            amount_row, receipt_row, action_row = 6, 7, 8
        else:
            self.rowconfigure(3, minsize=28)
            amount_row, receipt_row, action_row = 4, 5, 6
        self.rowconfigure(amount_row - 1, weight=1)
        self._field("amount", "Сумма, ₽", amount_row)
        self._payment_field(receipt_row)
        bottom = ttk.Frame(self, style="Card.TFrame")
        bottom.grid(
            row=action_row, column=0, columnspan=2, sticky="ew", pady=(4, 0),
        )
        ttk.Checkbutton(
            bottom, text="Оплачено", variable=self.paid, command=self._on_paid,
            style="Card.TCheckbutton",
        ).pack(side="left")
        self.clear()
        self.refresh_meter_labels()
        self.fields["period"].trace_add(
            "write", lambda *_args: self.refresh_suggestions()
        )
        self.refresh_suggestions()

    def focus_field(self) -> None:
        self.input_widgets["period"].focus_set()

    def _field(self, name: str, label: str, row: int, *, meter_label: bool = False) -> None:
        self.fields[name] = tk.StringVar()
        if not hasattr(self, "input_widgets"):
            self.input_widgets: dict[str, ttk.Entry] = {}
        label_widget = ttk.Label(self, text=label, style="Card.TLabel")
        label_widget.grid(
            row=row, column=0, sticky="w", padx=(0, 7), pady=1
        )
        if meter_label:
            self.meter_labels[name] = label_widget
        if name == "period":
            wrapper = ttk.Frame(self, style="Card.TFrame")
            wrapper.grid(row=row, column=1, sticky="ew", pady=1)
            wrapper.columnconfigure(0, weight=1)
            entry = ttk.Entry(wrapper, textvariable=self.fields[name], width=12)
            entry.grid(row=0, column=0, sticky="ew")
            ttk.Button(
                wrapper, text="…", width=2,
                command=lambda: self.app.choose_date_time(
                    self.fields["period"], with_time=False,
                ),
            ).grid(row=0, column=1, padx=(3, 0))
        else:
            entry = ttk.Entry(self, textvariable=self.fields[name], width=14)
            entry.grid(row=row, column=1, sticky="ew", pady=1)
        entry.bind("<Button-3>", self.app.show_edit_menu)
        entry.bind(
            "<FocusIn>",
            lambda _event: setattr(self.app, "active_service", self.service),
        )
        self.input_widgets[name] = entry
        if meter_label:
            suggestion = tk.Label(
                entry, text="", bg="#ffffff", fg="#98a59e",
                font=("Segoe UI", 9), borderwidth=0, cursor="xterm",
            )
            suggestion.bind("<Button-1>", lambda _event, target=entry: target.focus_set())
            self.suggestion_labels[name] = suggestion
            self.suggestion_values[name] = None
            self.fields[name].trace_add(
                "write",
                lambda *_args, field=name: self._show_suggestion(field),
            )

    def _show_suggestion(self, name: str) -> None:
        label = self.suggestion_labels[name]
        suggestion = self.suggestion_values.get(name)
        if suggestion is None or self.fields[name].get().strip():
            label.place_forget()
        else:
            label.configure(text=reading(suggestion))
            label.place(x=7, rely=0.5, anchor="w")

    def refresh_suggestions(self, entries: list[Entry] | None = None) -> None:
        if not self.suggestion_labels:
            return
        entries = entries if entries is not None else self.app.db.all()
        for name in self.suggestion_labels:
            meter_key = (
                "water_hot_meter"
                if self.service == "Вода" and name == "value_2"
                else "water_cold_meter"
                if self.service == "Вода"
                else "electric_meter"
                if self.service == "Электроэнергия"
                else "gas_meter"
            )
            meter_number = (
                self.meter_override[1 if meter_key == "water_hot_meter" else 0]
                if self.meter_override is not None else self.app.db.setting(meter_key)
            )
            self.suggestion_values[name] = suggested_reading(
                entries, self.service, name, self.fields["period"].get(),
                meter_number=meter_number, exclude_id=self.edit_id,
            )
            self._show_suggestion(name)

    def _payment_field(self, row: int) -> None:
        self.fields["payment_date"] = tk.StringVar()
        self.fields["payment_time"] = tk.StringVar()
        ttk.Label(self, text="Чек", style="Card.TLabel").grid(
            row=row, column=0, sticky="w", padx=(0, 7), pady=1
        )
        box = ttk.Frame(self, style="Card.TFrame")
        box.grid(row=row, column=1, sticky="ew", pady=1)
        box.columnconfigure(0, weight=1)
        date_entry = ttk.Entry(box, textvariable=self.fields["payment_date"], width=10)
        date_entry.grid(row=0, column=0, sticky="ew")
        ttk.Button(
            box, text="…", width=2,
            command=lambda: self.app.choose_date_time(
                self.fields["payment_date"], with_time=False
            ),
        ).grid(row=0, column=1, padx=(3, 5))
        time_entry = ttk.Entry(box, textvariable=self.fields["payment_time"], width=5)
        time_entry.grid(row=0, column=2)
        for name, widget in (
            ("payment_date", date_entry), ("payment_time", time_entry)
        ):
            widget.bind("<Button-3>", self.app.show_edit_menu)
            widget.bind(
                "<FocusIn>",
                lambda _event: setattr(self.app, "active_service", self.service),
            )
            self.input_widgets[name] = widget

    def refresh_meter_labels(self) -> None:
        def short(value: str) -> str:
            return value[-3:] if value else "—"

        def dated(number_value: str, date_key: str) -> str:
            date_value = self.app.db.setting(date_key)
            return f"№{short(number_value)} · {date_value or 'дата —'}"

        if self.service == "Вода":
            first, second = self.meter_override or (
                self.app.db.setting("water_cold_meter"),
                self.app.db.setting("water_hot_meter"),
            )
            self.meter_labels["value_1"].configure(text="ХВС")
            self.meter_labels["value_2"].configure(text="ГВС")
            self.meter_meta.configure(
                text=f"ХВС {dated(first, 'water_cold_date')}    "
                     f"ГВС {dated(second, 'water_hot_date')}"
            )
        elif self.service == "Электроэнергия":
            first = (
                self.meter_override[0] if self.meter_override
                else self.app.db.setting("electric_meter")
            )
            self.meter_labels["value_1"].configure(text="День")
            self.meter_labels["value_2"].configure(text="Ночь")
            self.meter_meta.configure(
                text=f"Счётчик {dated(first, 'electric_date')}"
            )
        elif self.service == "Газ":
            first = (
                self.meter_override[0] if self.meter_override
                else self.app.db.setting("gas_meter")
            )
            self.meter_labels["value_1"].configure(text="Показание")
            self.meter_meta.configure(
                text=f"Счётчик {dated(first, 'gas_date')}" if first else "Счётчик —"
            )
            gas_input = self.input_widgets["value_1"]
            gas_input.configure(state="normal" if first.strip() else "disabled")
            if not first.strip():
                self.fields["value_1"].set("")

    def _water_total(self, *_args: object) -> None:
        try:
            first = number(self.fields["value_1"].get(), "ХВС")
            second = number(self.fields["value_2"].get(), "ГВС")
            self.total_label.configure(text=f"Общее воды: {reading(first + second)} м³")
        except ValueError:
            self.total_label.configure(text="Общее воды: —")

    def _on_paid(self) -> None:
        if not self.paid.get():
            self.fields["payment_time"].set("")

    def _entry(self) -> Entry:
        amount = number(self.fields["amount"].get(), "Сумма", required=False)
        raw_payment_date = self.fields["payment_date"].get().strip()
        raw_payment_time = self.fields["payment_time"].get().strip()
        if self.edit_id is not None:
            old_entry = self.app.db.get(self.edit_id)
        else:
            old_entry = None
        if (
            old_entry is not None
            and raw_payment_date == old_entry.payment_date
            and not raw_payment_time
            and len(raw_payment_date) == 10
        ):
            payment_date = old_entry.payment_date
        elif not self.paid.get() and not raw_payment_time:
            valid_date(raw_payment_date)
            payment_date = ""
        else:
            payment_date = valid_payment_time(
                f"{raw_payment_date} {raw_payment_time}"
            )
        if self.paid.get() and (amount is None or not payment_date):
            raise ValueError("Для оплаты укажите сумму и дату")
        if payment_date and not self.paid.get():
            raise ValueError("Для даты оплаты установите отметку «Оплачено»")
        values = {}
        if self.service in METERS:
            if self.meter_override is not None:
                meter_1, meter_2 = self.meter_override
            elif self.service == "Вода":
                meter_1 = self.app.db.setting("water_cold_meter")
                meter_2 = self.app.db.setting("water_hot_meter")
            else:
                meter_1 = self.app.db.setting("electric_meter")
                meter_2 = ""
            values["meter_1"] = meter_1
            values["value_1"] = number(self.fields["value_1"].get(), METERS[self.service][0])
            values["value_2"] = number(self.fields["value_2"].get(), METERS[self.service][1])
            if self.service == "Вода":
                values["meter_2"] = meter_2
        elif self.service == "Газ":
            values["meter_1"] = (
                self.meter_override[0] if self.meter_override
                else self.app.db.setting("gas_meter")
            )
            values["value_1"] = number(
                self.fields["value_1"].get(), "Газ", required=False
            )
        return Entry(
            id=self.edit_id, service=self.service,
            period=valid_date(self.fields["period"].get()), amount=amount,
            payment_date=payment_date, paid=self.paid.get(), **values,
        )

    def save(self) -> None:
        self.app.save_all()

    def has_data(self) -> bool:
        return (
            self.edit_id is not None
            or bool(self.fields["amount"].get().strip())
            or bool(self.fields["payment_time"].get().strip())
            or self.paid.get()
            or any(
                self.fields[name].get().strip()
                for name in ("value_1", "value_2")
                if name in self.fields
            )
        )

    def clear(self) -> None:
        self.edit_id = None
        self.meter_override = None
        for variable in self.fields.values():
            variable.set("")
        self.fields["period"].set(date.today().strftime("%d.%m.%Y"))
        self.fields["payment_date"].set(date.today().strftime("%d.%m.%Y"))
        self.paid.set(False)
        if self.service in METERS or self.service == "Газ":
            self.refresh_meter_labels()
        self.refresh_suggestions()

    def load(self, entry: Entry) -> None:
        self.edit_id = entry.id
        for name, variable in self.fields.items():
            if name == "payment_date":
                value = entry.payment_date[:10] if entry.payment_date else ""
            elif name == "payment_time":
                value = entry.payment_date[11:] if len(entry.payment_date) > 10 else ""
            else:
                value = getattr(entry, name)
            variable.set(str(value) if value is not None else "")
        if self.service in METERS or self.service == "Газ":
            self.meter_override = (entry.meter_1, entry.meter_2)
            self.refresh_meter_labels()
        self.paid.set(entry.paid)
        self.refresh_suggestions()


class App(tk.Tk):
    def __init__(self, db: Database | None = None):
        if sys.platform == "win32":
            try:
                ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
                    "MyMeterReadings.Desktop.3"
                )
            except (AttributeError, OSError):
                pass
        super().__init__()
        # Build the complete interface while hidden so Tk never paints its small
        # default window before the final centered geometry is known.
        self.withdraw()
        center_window(self, 1080, 850)
        self._icons = [
            tk.PhotoImage(file=str(ICON_PATH.with_name(f"app_icon_{size}.png")))
            for size in ICON_SIZES
            if ICON_PATH.with_name(f"app_icon_{size}.png").exists()
        ]
        if self._icons:
            self.iconphoto(True, *self._icons)
        if sys.platform == "win32" and ICON_ICO_PATH.exists():
            self.iconbitmap(str(ICON_ICO_PATH))
        self.cloud = None
        self.update_controls = None
        self._preparing_update = False
        self._saving = False
        if db is None:
            prepare_database_location()
            if not prepare_database(self, DB_PATH):
                self.destroy()
                raise SystemExit(0)
        self.db = db or Database(DB_PATH)
        self.title("Мои показания")
        self.geometry("1080x850")
        self.minsize(1000, 810)
        self.protocol("WM_DELETE_WINDOW", self.close)
        style = ttk.Style(self)
        style.theme_use("clam")
        surface = "#e9ebe8"
        card = "#f7f7f4"
        input_bg = "#ffffff"
        ink = "#293431"
        line = "#d0d6d1"
        accent = "#49645d"
        style.configure(".", font=("Segoe UI", 9), foreground=ink)
        style.configure("TFrame", background=surface)
        style.configure("Card.TFrame", background=card)
        style.configure("TLabel", background=surface, foreground=ink)
        style.configure("Card.TLabel", background=card, foreground=ink)
        style.configure("CardTitle.TLabel", background=card, foreground="#2f5147",
                        font=("Segoe UI", 11, "bold"))
        style.configure("CardMeta.TLabel", background=card, foreground="#66756d",
                        font=("Segoe UI", 8))
        style.configure("TLabelframe", background=card, bordercolor=line,
                        relief="solid", borderwidth=1)
        style.configure("TLabelframe.Label", background=surface, foreground=ink,
                        font=("Segoe UI", 10, "bold"))
        style.configure("TEntry", fieldbackground=input_bg, foreground=ink,
                        bordercolor=line, lightcolor=line, darkcolor=line,
                        padding=(6, 4), relief="flat")
        style.map("TEntry", bordercolor=[("focus", accent)])
        style.configure("TCombobox", fieldbackground=input_bg, background=card,
                        foreground=ink, bordercolor=line, padding=3)
        style.configure("TCheckbutton", background=surface, foreground=ink)
        style.configure("Card.TCheckbutton", background=card, foreground=ink)
        style.map("Card.TCheckbutton", background=[("active", card)])
        style.configure("TButton", background="#e5e9e5", foreground=ink,
                        bordercolor="#d0d7d1", padding=(9, 4), relief="flat")
        style.map("TButton", background=[("active", "#d8e0da"), ("pressed", "#ccd8cf")])
        style.configure("Primary.TButton", background=accent, foreground="#ffffff",
                        bordercolor=accent, font=("Segoe UI", 9, "bold"))
        style.map("Primary.TButton", background=[("active", "#3d584f"),
                                                   ("pressed", "#344b43")])
        style.configure("Treeview", rowheight=25, font=("Segoe UI", 9),
                        background="#e8ece8", fieldbackground="#e8ece8",
                        foreground=ink, bordercolor=line)
        style.map("Treeview", background=[("selected", "#b9c9cc")],
                  foreground=[("selected", "#1b2a30")])
        style.configure("Treeview.Heading", background="#e2e7e2",
                        foreground=ink, font=("Segoe UI", 9, "bold"),
                        bordercolor=line, relief="flat")
        self.configure(background=surface)
        root = ttk.Frame(self, padding=12)
        root.pack(fill="both", expand=True)
        title_row = ttk.Frame(root)
        title_row.pack(fill="x", pady=(0, 8))
        header_icon = ICON_PATH.with_name("app_icon_40.png")
        if header_icon.exists():
            self._small_icon = tk.PhotoImage(file=str(header_icon))
            ttk.Label(title_row, image=self._small_icon).pack(side="left", padx=(0, 8))
        ttk.Label(title_row, text="Мои показания",
                  font=("Segoe UI", 17, "bold"),
                  foreground="#263b33").pack(side="left")
        self.update_controls = UpdateControls(self, title_row)
        if db is None:
            self.cloud = CloudControls(self, title_row, DB_PATH)
        forms = ttk.Frame(root)
        forms.pack(fill="x")
        for column in range(12):
            forms.columnconfigure(column, weight=1, uniform="services")
        self.forms: dict[str, CompactForm] = {}
        layout = {
            "Вода": (0, 0, 4),
            "Электроэнергия": (0, 4, 4),
            "Отопление": (1, 0, 3),
            "Газ": (1, 3, 3),
            "ТКО": (1, 6, 3),
            "Домофон": (1, 9, 3),
            "Домашний интернет": (0, 8, 4),
        }
        for service in SERVICES:
            row, column, span = layout[service]
            form = CompactForm(forms, self, service)
            self.forms[service] = form
            form.grid(
                row=row, column=column, columnspan=span,
                sticky="nsew", padx=4, pady=4,
            )
        history_header = ttk.Frame(root)
        history_header.pack(fill="x", pady=(8, 5))
        ttk.Label(history_header, text="История", font=("Segoe UI", 12, "bold")).pack(side="left")
        self.year_var = tk.StringVar()
        self.available_years: list[str] = []
        year_control = tk.Frame(
            history_header, bg="#f7f7f4", highlightthickness=1,
            highlightbackground="#cbd5cc", bd=0,
        )
        year_control.pack(side="left", padx=(16, 0))
        self.year_previous = tk.Label(
            year_control, text="‹", bg="#f7f7f4", fg="#49645d",
            font=("Segoe UI", 14), width=2, cursor="hand2",
        )
        self.year_previous.pack(side="left")
        self.year_previous.bind("<Button-1>", lambda _event: self._change_year(1))
        self.year_label = tk.Label(
            year_control, textvariable=self.year_var, bg="#f7f7f4", fg="#263b33",
            font=("Segoe UI", 10, "bold"), width=6, cursor="hand2",
        )
        self.year_label.pack(side="left")
        self.year_label.bind("<Button-1>", self._show_year_menu)
        self.year_next = tk.Label(
            year_control, text="›", bg="#f7f7f4", fg="#49645d",
            font=("Segoe UI", 14), width=2, cursor="hand2",
        )
        self.year_next.pack(side="left")
        self.year_next.bind("<Button-1>", lambda _event: self._change_year(-1))
        self.grand_total = ttk.Label(
            history_header, text="", font=("Segoe UI", 10, "bold")
        )
        self.grand_total.pack(side="right")
        table_frame = ttk.Frame(root)
        table_frame.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(
            table_frame,
            columns=("service", "values", "change", "amount", "payment", "status"),
            show="tree headings", selectmode="browse", height=5,
        )
        self.tree.heading("#0", text="Месяц", anchor="center")
        self.tree.column("#0", width=90, minwidth=80, anchor="center")
        for key, title, width in (
            ("service", "Услуга", 110),
            ("values", "Показания", 240),
            ("change", "Изменения в показаниях", 180),
            ("amount", "Сумма", 100),
            ("payment", "Дата и время оплаты", 150),
            ("status", "Статус", 100),
        ):
            self.tree.heading(key, text=title, anchor="center")
            self.tree.column(key, width=width, minwidth=60, anchor="center")
        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.tag_configure("month", background="#e8ece8", foreground="#40554d")
        self.tree.tag_configure(
            "attention", background="#f9e8e6", foreground="#a33d3d"
        )
        scrollbar = ModernScrollbar(table_frame, command=self.tree.yview)
        scrollbar.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.bind("<Double-1>", self._open_history_row)
        footer = ttk.Frame(root)
        footer.pack(fill="x", pady=(7, 0))
        ttk.Button(footer, text="Изменить", command=self.edit).pack(side="left")
        ttk.Button(footer, text="Удалить", command=self.delete).pack(side="left", padx=6)
        ttk.Button(footer, text="Клавиши", command=self.show_shortcuts).pack(
            side="left", padx=(8, 0)
        )
        ttk.Button(
            footer, text="Номера счётчиков", command=self.edit_meter_settings
        ).pack(side="left", padx=(8, 0))
        ttk.Button(footer, text="Экспорт CSV", command=self.export).pack(side="right")
        ttk.Button(footer, text="Экспорт Excel", command=self.export_excel).pack(
            side="right", padx=(0, 6)
        )
        self.save_button = ttk.Button(
            footer, text="Сохранить", command=self.save_all,
            style="Primary.TButton",
        )
        self.save_button.pack(side="right", padx=(0, 8))
        ttk.Button(
            footer, text="Очистить", command=self.clear_all,
        ).pack(side="right", padx=(0, 6))
        self._bind_shortcuts()
        self.refresh()
        target_height = min(
            max(850, self.winfo_reqheight() + 18),
            self.winfo_screenheight() - 70,
        )
        center_window(self, 1080, target_height)
        self.deiconify()

    def _bind_shortcuts(self) -> None:
        self.bind_class("UtilityHotkeys", "<KeyPress>", self._route_hotkey)
        for form in self.forms.values():
            for widget in form.input_widgets.values():
                self._enable_hotkeys(widget)
        self._enable_hotkeys(self.tree)
        for index, service in enumerate(SERVICES, 1):
            self.bind_all(
                f"<Control-Key-{index}>",
                lambda _event, name=service: self._focus_service(name),
            )
        self.bind_all("<Control-s>", lambda _event: self._save_active())
        self.bind_all("<Control-Return>", lambda _event: self._save_active())
        self.bind_all("<Control-f>", lambda _event: self._focus_history())
        self.bind_all("<Control-e>", lambda _event: self.export_excel())
        self.bind_all("<F2>", lambda _event: self.edit())
        self.bind_all("<Delete>", lambda _event: self._delete_if_history())
        self.bind_all("<Escape>", lambda _event: self._clear_active())
        self.bind_all("<Control-Left>", lambda _event: self._change_year(-1))
        self.bind_all("<Control-Right>", lambda _event: self._change_year(1))
        self.bind_all("<F1>", lambda _event: self.show_shortcuts())
        self.bind_all("<Control-m>", lambda _event: self.edit_meter_settings())
        for sequence, operation in (
            ("<Control-c>", "copy"), ("<Control-v>", "paste"),
            ("<Control-x>", "cut"), ("<Control-a>", "select_all"),
            ("<Control-Insert>", "copy"), ("<Shift-Insert>", "paste"),
            ("<Shift-Delete>", "cut"),
        ):
            self.bind_all(
                sequence,
                lambda event, action=operation: self._clipboard_action(action, event.widget),
            )

    def _enable_hotkeys(self, widget: tk.Misc) -> None:
        tags = widget.bindtags()
        if "UtilityHotkeys" not in tags:
            widget.bindtags(("UtilityHotkeys", *tags))
        if isinstance(widget, (ttk.Entry, tk.Entry, tk.Spinbox)):
            widget.bind("<Button-3>", self.show_edit_menu)

    def _route_hotkey(self, event: tk.Event) -> str | None:
        key = event.keysym.lower()
        control = bool(event.state & 0x4)
        if self.tk.call("tk", "windowingsystem") == "win32":
            physical_keys = {
                49: "1", 50: "2", 51: "3", 52: "4", 53: "5", 54: "6", 55: "7",
                65: "a", 67: "c", 69: "e", 70: "f", 77: "m",
                83: "s", 86: "v", 88: "x",
                13: "return", 37: "left", 39: "right",
            }
            key = physical_keys.get(event.keycode & 0xFF, key)
        dialog = event.widget.winfo_toplevel()
        if dialog is not self and not (control and key in ("c", "v", "x", "a")):
            if control and key in ("s", "return"):
                save_dialog = getattr(dialog, "shortcut_save", None)
                if callable(save_dialog):
                    save_dialog()
                    return "break"
            return None
        if control:
            if key in "1234567" and len(key) == 1:
                self._focus_service(SERVICES[int(key) - 1])
            elif key in ("s", "return"):
                self._save_active()
            elif key == "f":
                self._focus_history()
            elif key == "e":
                self.export_excel()
            elif key == "m":
                self.edit_meter_settings()
            elif key == "left":
                self._change_year(-1)
            elif key == "right":
                self._change_year(1)
            elif key in ("c", "v", "x", "a"):
                return self._clipboard_action(
                    {"c": "copy", "v": "paste", "x": "cut", "a": "select_all"}[key],
                    event.widget,
                )
            else:
                return None
            return "break"
        if key == "f1":
            self.show_shortcuts()
            return "break"
        if key == "f2":
            self.edit()
            return "break"
        if key == "delete" and event.widget == self.tree:
            self.delete()
            return "break"
        if key == "escape":
            self._clear_active()
            return "break"
        return None

    def _clipboard_action(self, action: str, widget: tk.Misc) -> str | None:
        if isinstance(widget, (ttk.Entry, tk.Entry, tk.Spinbox)):
            if action == "select_all":
                widget.selection_range(0, "end")
                widget.icursor("end")
            elif action == "copy":
                try:
                    start = widget.index("sel.first")
                    end = widget.index("sel.last")
                except tk.TclError:
                    return "break"
                self.clipboard_clear()
                self.clipboard_append(widget.get()[start:end])
            elif action == "cut":
                try:
                    start = widget.index("sel.first")
                    end = widget.index("sel.last")
                except tk.TclError:
                    return "break"
                self.clipboard_clear()
                self.clipboard_append(widget.get()[start:end])
                widget.delete(start, end)
            elif action == "paste":
                try:
                    value = self.clipboard_get()
                except tk.TclError:
                    return "break"
                try:
                    widget.delete("sel.first", "sel.last")
                except tk.TclError:
                    pass
                widget.insert("insert", value)
            return "break"
        if action == "copy" and widget == self.tree:
            selected = self.tree.selection()
            if selected:
                iid = selected[0]
                cells = [self.tree.item(iid, "text"), *self.tree.item(iid, "values")]
                self.clipboard_clear()
                self.clipboard_append("\t".join(str(cell) for cell in cells))
            return "break"
        return None

    def show_edit_menu(self, event: tk.Event) -> None:
        widget = event.widget
        widget.focus_set()
        menu = tk.Menu(self, tearoff=False)
        for label, action in (
            ("Вырезать", "cut"), ("Копировать", "copy"),
            ("Вставить", "paste"), ("Выделить всё", "select_all"),
        ):
            menu.add_command(
                label=label,
                command=lambda selected=action: self._clipboard_action(selected, widget),
            )
        menu.tk_popup(event.x_root, event.y_root)
        menu.grab_release()

    def edit_meter_settings(self) -> None:
        dialog = tk.Toplevel(self)
        dialog.withdraw()
        if self._icons:
            dialog.iconphoto(False, *self._icons)
        if sys.platform == "win32" and ICON_ICO_PATH.exists():
            dialog.iconbitmap(str(ICON_ICO_PATH))
        dialog.title("Номера счётчиков")
        dialog.transient(self)
        dialog.resizable(False, False)
        content = ttk.Frame(dialog, padding=18)
        content.pack(fill="both", expand=True)
        fields = (
            ("water_cold_meter", "water_cold_date", "ХВС"),
            ("water_hot_meter", "water_hot_date", "ГВС"),
            ("electric_meter", "electric_date", "Электроэнергия"),
            ("gas_meter", "gas_date", "Газ"),
        )
        variables: dict[str, tk.StringVar] = {}
        ttk.Label(content, text="Номер").grid(row=0, column=1, sticky="w")
        ttk.Label(content, text="Дата").grid(row=0, column=2, sticky="w")
        for row, (key, date_key, label) in enumerate(fields, start=1):
            ttk.Label(content, text=label).grid(row=row, column=0, sticky="w", padx=(0, 16), pady=6)
            variable = tk.StringVar(value=self.db.setting(key))
            variables[key] = variable
            number_entry = ttk.Entry(content, textvariable=variable, width=22)
            number_entry.grid(row=row, column=1, sticky="ew", pady=6)
            self._enable_hotkeys(number_entry)
            date_variable = tk.StringVar(value=self.db.setting(date_key))
            variables[date_key] = date_variable
            date_box = ttk.Frame(content)
            date_box.grid(row=row, column=2, sticky="ew", padx=(8, 0), pady=6)
            date_entry = ttk.Entry(date_box, textvariable=date_variable, width=13)
            date_entry.pack(side="left")
            self._enable_hotkeys(date_entry)
            ttk.Button(
                date_box, text="…", width=2,
                command=lambda target=date_variable: self.choose_date_time(
                    target, with_time=False
                ),
            ).pack(side="left", padx=(3, 0))

        def save_numbers() -> None:
            try:
                for _key, date_key, _label in fields:
                    if variables[date_key].get().strip():
                        valid_date(variables[date_key].get())
            except ValueError as exc:
                messagebox.showerror("Дата счётчика", str(exc), parent=dialog)
                return
            for key, variable in variables.items():
                self.db.set_setting(key, variable.get(), commit=False)
            self.db.connection.commit()
            if self.cloud is not None:
                self.cloud.saved()
            for service in (*METERS, "Газ"):
                self.forms[service].refresh_meter_labels()
            dialog.destroy()

        actions = ttk.Frame(content)
        actions.grid(row=len(fields) + 1, column=0, columnspan=3, sticky="e", pady=(12, 0))
        ttk.Button(actions, text="Отмена", command=dialog.destroy).pack(side="right")
        ttk.Button(actions, text="Сохранить", command=save_numbers).pack(
            side="right", padx=(0, 8)
        )
        dialog.bind("<Escape>", lambda _event: dialog.destroy())
        dialog.shortcut_save = save_numbers
        center_window(dialog)
        dialog.deiconify()
        dialog.grab_set()
        dialog.focus_set()

    def choose_date_time(
        self,
        target: tk.StringVar,
        *,
        with_time: bool,
        on_select: object = None,
    ) -> None:
        raw = target.get().strip()
        initial = None
        for pattern in ("%d.%m.%Y %H:%M", "%d.%m.%Y"):
            try:
                initial = datetime.strptime(raw, pattern)
                break
            except ValueError:
                pass
        initial = initial or datetime.now()
        dialog = tk.Toplevel(self)
        dialog.withdraw()
        if self._icons:
            dialog.iconphoto(False, *self._icons)
        if sys.platform == "win32" and ICON_ICO_PATH.exists():
            dialog.iconbitmap(str(ICON_ICO_PATH))
        dialog.title("Дата и время оплаты" if with_time else "Выбрать дату")
        dialog.transient(self)
        dialog.resizable(False, False)
        content = ttk.Frame(dialog, padding=12)
        content.pack(fill="both", expand=True)
        selected = tk.IntVar(value=initial.day)
        shown_year = initial.year
        shown_month = initial.month
        header = ttk.Frame(content)
        header.pack(fill="x", pady=(0, 8))
        calendar_frame = ttk.Frame(content)
        calendar_frame.pack()
        month_names = (
            "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
            "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь",
        )

        def render_calendar() -> None:
            for child in calendar_frame.winfo_children():
                child.destroy()
            month_title.configure(text=f"{month_names[shown_month - 1]} {shown_year}")
            for column, label in enumerate(("Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс")):
                ttk.Label(calendar_frame, text=label, anchor="center", width=5).grid(
                    row=0, column=column, pady=(0, 3)
                )
            weeks = calendar.monthcalendar(shown_year, shown_month)
            for row, week in enumerate(weeks, 1):
                for column, day in enumerate(week):
                    if day:
                        title = f"[{day}]" if day == selected.get() else str(day)
                        ttk.Button(
                            calendar_frame, text=title, width=4,
                            command=lambda value=day: choose_day(value),
                        ).grid(row=row, column=column, padx=1, pady=1)

        def choose_day(day: int) -> None:
            selected.set(day)
            render_calendar()

        def move_month(delta: int) -> None:
            nonlocal shown_year, shown_month
            index = shown_year * 12 + shown_month - 1 + delta
            shown_year, zero_month = divmod(index, 12)
            shown_month = zero_month + 1
            selected.set(1)
            render_calendar()

        ttk.Button(header, text="‹", width=3, command=lambda: move_month(-1)).pack(side="left")
        month_title = ttk.Label(header, text="", anchor="center")
        month_title.pack(side="left", fill="x", expand=True)
        ttk.Button(header, text="›", width=3, command=lambda: move_month(1)).pack(side="right")
        has_time = len(raw) > 10
        hour = tk.StringVar(value=f"{initial.hour:02d}" if has_time else "")
        minute = tk.StringVar(value=f"{initial.minute:02d}" if has_time else "")
        if with_time:
            clock = ttk.Frame(content)
            clock.pack(pady=(10, 2))
            ttk.Label(clock, text="Время с чека:").pack(side="left", padx=(0, 8))
            hour_entry = tk.Spinbox(
                clock, from_=0, to=23, format="%02.0f", wrap=True,
                width=3, textvariable=hour,
            )
            hour_entry.pack(side="left")
            self._enable_hotkeys(hour_entry)
            ttk.Label(clock, text=":").pack(side="left")
            minute_entry = tk.Spinbox(
                clock, from_=0, to=59, format="%02.0f", wrap=True,
                width=3, textvariable=minute,
            )
            minute_entry.pack(side="left")
            self._enable_hotkeys(minute_entry)

        def apply() -> None:
            try:
                chosen = date(shown_year, shown_month, selected.get())
                output = chosen.strftime("%d.%m.%Y")
                if with_time:
                    if not hour.get().strip() or not minute.get().strip():
                        raise ValueError
                    hours = int(hour.get())
                    minutes = int(minute.get())
                    if not 0 <= hours <= 23 or not 0 <= minutes <= 59:
                        raise ValueError
                    output += f" {hours:02d}:{minutes:02d}"
                target.set(output)
                if callable(on_select):
                    on_select()
            except ValueError:
                messagebox.showerror("Время", "Укажите часы от 00 до 23 и минуты от 00 до 59.", parent=dialog)
                return
            dialog.destroy()

        actions = ttk.Frame(content)
        actions.pack(fill="x", pady=(12, 0))
        ttk.Button(actions, text="Отмена", command=dialog.destroy).pack(side="right")
        ttk.Button(actions, text="Выбрать", command=apply).pack(side="right", padx=(0, 7))
        dialog.bind("<Return>", lambda _event: apply())
        dialog.bind("<Escape>", lambda _event: dialog.destroy())
        dialog.shortcut_save = apply
        render_calendar()
        center_window(dialog)
        dialog.deiconify()
        dialog.grab_set()
        dialog.focus_set()

    def _focus_service(self, service: str) -> None:
        form = self.forms[service]
        form.focus_field()
        self.active_service = service

    def _active_form(self) -> CompactForm:
        focus = self.focus_get()
        while focus is not None:
            if isinstance(focus, CompactForm):
                self.active_service = focus.service
                return focus
            focus = focus.master
        return self.forms[getattr(self, "active_service", "Вода")]

    def _save_active(self) -> None:
        self.save_all()

    def save_all(self) -> None:
        if self._saving:
            return
        self._saving = True
        self.save_button.state(["disabled"])
        try:
            prepared: list[tuple[CompactForm, Entry]] = []
            for service in SERVICES:
                form = self.forms[service]
                if not form.has_data():
                    continue
                try:
                    prepared.append((form, form._entry()))
                except ValueError as exc:
                    form.focus_field()
                    messagebox.showerror(service, str(exc), parent=self)
                    return
            if not prepared:
                messagebox.showinfo("Сохранение", "Заполните хотя бы один раздел.", parent=self)
                return
            try:
                with self.db.connection:
                    for _form, entry in prepared:
                        if entry.id is None:
                            self.db.add(entry, commit=False)
                        else:
                            self.db.update(entry, commit=False)
            except DuplicateEntry as exc:
                messagebox.showinfo("Запись уже существует", str(exc), parent=self)
                return
            except sqlite3.Error as exc:
                messagebox.showerror("Сохранение", str(exc), parent=self)
                return
            for form, _entry in prepared:
                form.clear()
            self.refresh()
            if self.cloud is not None:
                self.cloud.saved()
        finally:
            self._saving = False
            self.save_button.state(["!disabled"])

    def _clear_active(self) -> None:
        self._active_form().clear()

    def clear_all(self) -> None:
        if any(form.has_data() for form in self.forms.values()):
            if not messagebox.askyesno(
                "Очистить поля",
                "Очистить данные во всех разделах? Сохранённая история не изменится.",
                parent=self,
            ):
                return
        for form in self.forms.values():
            form.clear()

    def _focus_history(self) -> None:
        self.tree.focus_set()

    def _delete_if_history(self) -> None:
        if self.focus_get() == self.tree:
            self.delete()

    def _change_year(self, direction: int) -> None:
        options = self.available_years
        current = self.year_var.get()
        if current in options:
            next_index = options.index(current) + direction
            if 0 <= next_index < len(options):
                self.year_var.set(options[next_index])
                self.refresh()

    def _show_year_menu(self, event: tk.Event) -> None:
        menu = tk.Menu(
            self, tearoff=False, bg="#f7f7f4", fg="#263b33",
            activebackground="#dce6dd", activeforeground="#263b33",
        )
        for year in self.available_years:
            menu.add_command(
                label=year,
                command=lambda selected=year: self._select_year(selected),
            )
        menu.tk_popup(event.x_root, event.y_root)
        menu.grab_release()

    def _select_year(self, year: str) -> None:
        self.year_var.set(year)
        self.refresh()

    def show_shortcuts(self) -> None:
        messagebox.showinfo(
            "Горячие клавиши",
            "Ctrl+1…6 — перейти к разделу\n"
            "Ctrl+S или Ctrl+Enter — сохранить открытую форму\n"
            "Ctrl+F — перейти к истории\n"
            "F2 — изменить выбранную запись\n"
            "Delete — удалить выбранную запись в истории\n"
            "Esc — очистить открытую форму\n"
            "Ctrl+← / Ctrl+→ — сменить год\n"
            "Ctrl+E — экспорт Excel\n"
            "Ctrl+M — номера счётчиков\n"
            "Ctrl+C / V / X / A — копировать / вставить / вырезать / выделить всё\n"
            "Правая кнопка мыши — действия с текстом\n"
            "F1 — показать этот список",
            parent=self,
        )

    def refresh(self) -> None:
        entries = self.db.all()
        for form in self.forms.values():
            form.refresh_suggestions(entries)
        recorded_years = {
            datetime.strptime(item.period, "%d.%m.%Y").year for item in entries
        }
        first_year = min({2025, date.today().year} | recorded_years)
        last_year = max({date.today().year} | recorded_years)
        options = list(range(last_year, first_year - 1, -1))
        years = [str(year) for year in options]
        self.available_years = years
        if self.year_var.get() not in years:
            self.year_var.set(years[0])
        selected_index = years.index(self.year_var.get())
        self.year_previous.configure(
            fg="#49645d" if selected_index < len(years) - 1 else "#b4c0b7"
        )
        self.year_next.configure(
            fg="#49645d" if selected_index > 0 else "#b4c0b7"
        )
        selected_year = int(self.year_var.get())
        visible = [
            item for item in entries
            if datetime.strptime(item.period, "%d.%m.%Y").year == selected_year
        ]
        total = sum(
            (item.amount for item in visible if item.paid and item.amount is not None),
            Decimal("0"),
        )
        self.grand_total.configure(text=f"Оплачено за {selected_year}: {money(total)}")
        self.tree.delete(*self.tree.get_children())
        month_names = (
            "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
            "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь",
        )
        paid_by_month = {month: Decimal("0") for month in range(1, 13)}
        recorded_by_month = {month: set() for month in range(1, 13)}
        for item in visible:
            month = datetime.strptime(item.period, "%d.%m.%Y").month
            recorded_by_month[month].add(item.service)
            if item.paid and item.amount is not None:
                paid_by_month[month] += item.amount
        for month, name in enumerate(month_names, 1):
            self.tree.insert(
                "", "end", iid=f"month:{month}", text=name,
                open=(selected_year == date.today().year and month == date.today().month),
                tags=("month",),
                values=("", "", "", money(paid_by_month[month]), "", "—"),
            )
        prior_by_id: dict[int, Entry | None] = {}
        for service in (*METERS, "Газ"):
            previous = None
            service_entries = [item for item in entries if item.service == service]
            for item in reversed(service_entries):
                prior_by_id[item.id] = previous
                previous = item
        for item in visible:
            month = datetime.strptime(item.period, "%d.%m.%Y").month
            values = change = ""
            if item.service == "Вода":
                values = (
                    f"ХВС {reading(item.value_1)}; ГВС {reading(item.value_2)}; "
                    f"всего {reading(item.total_water)}"
                )
            elif item.service == "Электроэнергия":
                values = f"День {reading(item.value_1)}; ночь {reading(item.value_2)}"
            elif item.service == "Газ" and item.value_1 is not None:
                values = reading(item.value_1)
            if item.service in METERS:
                prior = prior_by_id[item.id]
                if prior and None not in (
                    item.value_1, item.value_2, prior.value_1, prior.value_2
                ):
                    first = item.value_1 - prior.value_1
                    second = item.value_2 - prior.value_2
                    if item.service == "Вода":
                        change = (
                            f"ХВС {reading(first)}; ГВС {reading(second)}; "
                            f"всего {reading(first + second)}"
                        )
                    else:
                        change = f"День {reading(first)}; ночь {reading(second)}"
                else:
                    change = "Первая запись"
            elif item.service == "Газ" and item.value_1 is not None:
                prior = prior_by_id[item.id]
                if prior and prior.value_1 is not None:
                    change = reading(item.value_1 - prior.value_1)
                else:
                    change = "Первая запись"
            self.tree.insert(
                f"month:{month}", "end", iid=f"entry:{item.id}",
                values=(
                    item.service, values, change,
                    money(item.amount) if item.amount is not None else "—",
                    item.payment_date or "—",
                    "Оплачено" if item.paid else "Не оплачено",
                ),
                tags=() if item.paid else ("attention",),
            )
        for month in range(1, 13):
            if not recorded_by_month[month]:
                continue
            for service_index, service in enumerate(SERVICES):
                if service in recorded_by_month[month]:
                    continue
                self.tree.insert(
                    f"month:{month}", "end",
                    iid=f"missing:{month}:{service_index}",
                    values=(service, "—", "—", "—", "—", "Не заполнено"),
                    tags=("attention",),
                )

    def _selected_entry(self) -> Entry | None:
        selection = self.tree.selection()
        if not selection or not selection[0].startswith("entry:"):
            messagebox.showinfo("История", "Выберите запись внутри месяца.", parent=self)
            return None
        return self.db.get(int(selection[0].split(":", 1)[1]))

    def _open_history_row(self, event: tk.Event) -> None:
        row = self.tree.identify_row(event.y)
        if not row:
            return
        self.tree.selection_set(row)
        if row.startswith("month:"):
            self._load_month(int(row.split(":", 1)[1]))
        else:
            self.edit()

    def _load_month(self, month: int) -> None:
        if any(form.has_data() for form in self.forms.values()):
            if not messagebox.askyesno(
                "Открыть месяц",
                "В полях есть несохранённые данные или открытые записи. "
                "Заменить их данными выбранного месяца?",
                parent=self,
            ):
                return
        selected_year = int(self.year_var.get())
        monthly_entries = [
            item for item in self.db.all()
            if datetime.strptime(item.period, "%d.%m.%Y").year == selected_year
            and datetime.strptime(item.period, "%d.%m.%Y").month == month
            and item.service in self.forms
        ]
        latest_by_service = {}
        for item in monthly_entries:
            previous = latest_by_service.get(item.service)
            if previous is None or (
                datetime.strptime(item.period, "%d.%m.%Y"), item.id
            ) > (
                datetime.strptime(previous.period, "%d.%m.%Y"), previous.id
            ):
                latest_by_service[item.service] = item
        for service, form in self.forms.items():
            form.clear()
            if service in latest_by_service:
                form.load(latest_by_service[service])
            else:
                form.fields["period"].set(f"01.{month:02d}.{selected_year}")

    def edit(self) -> None:
        selection = self.tree.selection()
        if selection and selection[0].startswith("missing:"):
            _, month_text, service_index_text = selection[0].split(":")
            service = SERVICES[int(service_index_text)]
            form = self.forms[service]
            if form.has_data():
                messagebox.showinfo(
                    service,
                    "Сначала сохраните или очистите уже заполненную форму.",
                    parent=self,
                )
                return
            form.fields["period"].set(
                f"01.{int(month_text):02d}.{self.year_var.get()}"
            )
            focus_name = "value_1" if "value_1" in form.input_widgets else "amount"
            form.input_widgets[focus_name].focus_set()
            return
        entry = self._selected_entry()
        if entry:
            if entry.service not in self.forms:
                messagebox.showinfo(
                    "Архивная запись",
                    "Эта запись перенесена из старого поля «Прочее».",
                    parent=self,
                )
                return
            self.forms[entry.service].load(entry)
            self.forms[entry.service].focus_field()

    def delete(self) -> None:
        entry = self._selected_entry()
        if entry and messagebox.askyesno("Удаление", "Удалить выбранную запись?", parent=self):
            self.db.delete(entry.id)
            if entry.service in self.forms and self.forms[entry.service].edit_id == entry.id:
                self.forms[entry.service].clear()
            self.refresh()

    def export(self) -> None:
        path = filedialog.asksaveasfilename(
            parent=self, defaultextension=".csv", initialfile="moi_pokazaniya.csv",
            filetypes=[("CSV", "*.csv")],
        )
        if not path:
            return
        with open(path, "w", encoding="utf-8-sig", newline="") as stream:
            writer = csv.writer(stream, delimiter=";")
            writer.writerow((
                "Раздел", "Период / дата показаний", "Счётчик 1", "Показание 1",
                "Счётчик 2", "Показание 2", "Общее воды", "Сумма",
                "Дата оплаты", "Статус", "Заметка",
            ))
            for item in self.db.all():
                writer.writerow((
                    item.service, item.period, item.meter_1, reading(item.value_1),
                    item.meter_2, reading(item.value_2), reading(item.total_water),
                    str(item.amount) if item.amount is not None else "",
                    item.payment_date, "Оплачено" if item.paid else "Не оплачено",
                    item.note,
                ))
        messagebox.showinfo("Экспорт", f"Сохранено: {path}", parent=self)

    def export_excel(self) -> None:
        path = filedialog.asksaveasfilename(
            parent=self, defaultextension=".xlsx",
            initialfile="moi_pokazaniya.xlsx",
            filetypes=[("Книга Excel", "*.xlsx")],
        )
        if not path:
            return
        try:
            export_excel_workbook(Path(path), self.db)
        except ImportError:
            messagebox.showerror(
                "Экспорт Excel",
                "Для экспорта установите пакет openpyxl в интерпретаторе PyCharm:\n"
                "python -m pip install openpyxl",
                parent=self,
            )
        except (OSError, PermissionError) as exc:
            messagebox.showerror(
                "Экспорт Excel", f"Не удалось сохранить файл:\n{exc}", parent=self,
            )
        else:
            messagebox.showinfo("Экспорт Excel", f"Сохранено: {path}", parent=self)

    def close(self) -> None:
        if self.cloud is not None and not self.cloud.close():
            return
        if self.update_controls is not None:
            self.update_controls.close()
        self.db.close()
        self.destroy()

    def prepare_for_update(self) -> bool:
        """Завершить облачную отправку перед передачей файлов обновлятору."""
        if self._preparing_update:
            return True
        if self.cloud is not None and not self.cloud.close():
            return False
        self.cloud = None
        self._preparing_update = True
        return True

    def finish_for_update(self) -> None:
        if self.update_controls is not None:
            self.update_controls.close()
        self.db.close()
        self.destroy()


if __name__ == "__main__":
    if updates.is_update_in_progress():
        if sys.platform == "win32":
            ctypes.windll.user32.MessageBoxW(
                None,
                "Установка новой версии ещё не завершена. Подождите немного.",
                "Мои показания обновляются",
                0x40,
            )
        raise SystemExit(0)
    App().mainloop()
