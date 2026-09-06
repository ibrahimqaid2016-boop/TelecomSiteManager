__version__ = "1.0.0"

import sqlite3
import os
import tempfile
from datetime import datetime

try:
    from plyer import filechooser
except ImportError:
    filechooser = None

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
except ImportError:
    Workbook = None
    load_workbook = None


from kivy.app import App
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.gridlayout import GridLayout
from kivy.uix.scrollview import ScrollView
from kivy.uix.label import Label
from kivy.uix.button import Button
from kivy.uix.checkbox import CheckBox
from kivy.uix.textinput import TextInput
from kivy.uix.spinner import Spinner
from kivy.uix.popup import Popup
from kivy.metrics import dp
from kivy.core.window import Window


# ============================================================
# GLOBAL VARIABLES
# ============================================================

CURRENT_SITE_ID = None
CURRENT_SITE_NAME = ""

CURRENT_2G_ID = None
CURRENT_2G_SECTOR_ID = None

CURRENT_LTE_ID = None
CURRENT_LTE_RRU_ID = None
CURRENT_LTE_CELL_ID = None

CURRENT_MICROWAVE_ID = None
CURRENT_ALARM_SITE_ID = None


# ============================================================
# DATABASE
# ============================================================

def get_database_path():
    """Use a writable app-private database directory on Android."""
    try:
        app = App.get_running_app()
        if app is not None:
            folder = app.user_data_dir
        else:
            folder = os.path.dirname(os.path.abspath(__file__))
    except Exception:
        folder = os.path.dirname(os.path.abspath(__file__))
    os.makedirs(folder, exist_ok=True)
    return os.path.join(folder, "telecom_database.db")


def get_connection():
    return sqlite3.connect(get_database_path())


# ============================================================
# DATABASE INITIALIZATION
# ============================================================

def initialize_database():

    connection = get_connection()
    cursor = connection.cursor()

    # ========================================================
    # SITES
    # ========================================================

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS sites (

        id INTEGER PRIMARY KEY AUTOINCREMENT,

        site_code TEXT,
        site_name TEXT,
        location TEXT,
        region TEXT,

        environment TEXT,
        structure_type TEXT,
        structure_height TEXT,

        number_of_sectors TEXT,

        latitude TEXT,
        longitude TEXT,

        notes TEXT
    )
    """)

    # ========================================================
    # 2G SYSTEM
    # ========================================================

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS system_2g (

        id INTEGER PRIMARY KEY AUTOINCREMENT,

        site_id INTEGER,

        bts_vendor TEXT,
        bts_type TEXT,
        system_band TEXT,

        esma TEXT,
        esea TEXT,

        number_of_sectors TEXT,

        sector_configuration TEXT,

        notes TEXT
    )
    """)

    # ========================================================
    # 2G CARDS / HARDWARE INVENTORY
    # ========================================================

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS system_2g_cards (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        system_2g_id INTEGER NOT NULL,
        card_type TEXT NOT NULL,
        slot_position TEXT DEFAULT '',
        status TEXT DEFAULT 'Working',
        serial_number TEXT DEFAULT '',
        notes TEXT DEFAULT ''
    )
    """)

    # ========================================================
    # 2G SECTORS
    # ========================================================

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS system_2g_sectors (

        id INTEGER PRIMARY KEY AUTOINCREMENT,

        system_2g_id INTEGER,

        sector_number INTEGER,
        sector_name TEXT,
        band TEXT,

        exga_number TEXT,
        erga_number TEXT,

        exda_number TEXT,
        erda_number TEXT,

        antenna_type TEXT,
        azimuth TEXT,

        notes TEXT
    )
    """)

    # ========================================================
    # 2G SECTOR CARDS
    # ========================================================
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS system_2g_sector_cards (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        sector_id INTEGER NOT NULL,
        card_type TEXT NOT NULL,
        notes TEXT DEFAULT ''
    )
    """)

    # ========================================================
    # LTE SYSTEM
    # ========================================================

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS lte_system (

        id INTEGER PRIMARY KEY AUTOINCREMENT,

        site_id INTEGER,

        bbu_type TEXT,
        number_of_rru TEXT,
        rru_information TEXT,
        frequency_band TEXT,

        notes TEXT
    )
    """)

    # LTE SYSTEM custom BBU/frequency fields
    for column_sql in [
        "ALTER TABLE lte_system ADD COLUMN custom_bbu_name TEXT DEFAULT ''",
        "ALTER TABLE lte_system ADD COLUMN custom_frequency_band TEXT DEFAULT ''",
    ]:
        try:
            cursor.execute(column_sql)
        except sqlite3.OperationalError:
            pass

    # LTE BBU hardware cards (one row per installed card)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS lte_bbu_cards (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        lte_system_id INTEGER NOT NULL,
        card_type TEXT NOT NULL,
        notes TEXT DEFAULT ''
    )
    """)

    # ========================================================
    # LTE RRU
    # ========================================================

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS lte_rru (

        id INTEGER PRIMARY KEY AUTOINCREMENT,

        lte_system_id INTEGER,

        rru_name TEXT,
        rru_model TEXT,
        vendor TEXT,

        frequency_band TEXT,
        technology TEXT,

        tx_rx TEXT,
        sector TEXT,

        cpri_port TEXT,

        status TEXT,

        notes TEXT
    )
    """)

    # Custom values for RRU fields when the user selects Other.
    for _column, _definition in (
        ("custom_vendor", "TEXT DEFAULT ''"),
        ("custom_technology", "TEXT DEFAULT ''"),
        ("custom_tx_rx", "TEXT DEFAULT ''"),
    ):
        try:
            cursor.execute(
                f"ALTER TABLE lte_rru ADD COLUMN {_column} {_definition}"
            )
        except sqlite3.OperationalError:
            pass

    # ========================================================
    # LTE CELLS / SECTORS
    # ========================================================

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS lte_cells (

        id INTEGER PRIMARY KEY AUTOINCREMENT,

        lte_system_id INTEGER,

        cell_number INTEGER,
        cell_name TEXT,

        sector TEXT,

        pci TEXT,
        tac TEXT,

        earfcn TEXT,
        bandwidth TEXT,

        frequency_band TEXT,

        mimo TEXT,

        antenna_type TEXT,
        azimuth TEXT,

        status TEXT,

        notes TEXT
    )
    """)

    # ========================================================
    # MICROWAVE
    # ========================================================

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS microwave (

        id INTEGER PRIMARY KEY AUTOINCREMENT,

        site_id INTEGER,

        link_name TEXT,
        equipment_vendor TEXT,

        idu_type TEXT,
        odu_type TEXT,

        frequency TEXT,
        capacity TEXT,

        dish_size TEXT,
        polarization TEXT,

        remote_site TEXT,

        notes TEXT
    )
    """)

    # Extra Microwave Link fields. Added at the end to preserve old
    # column positions and existing database compatibility.
    for column_sql in [
        "ALTER TABLE microwave ADD COLUMN link_id TEXT DEFAULT ''",
        "ALTER TABLE microwave ADD COLUMN custom_vendor TEXT DEFAULT ''",
        "ALTER TABLE microwave ADD COLUMN custom_idu_type TEXT DEFAULT ''",
        "ALTER TABLE microwave ADD COLUMN custom_polarization TEXT DEFAULT ''",
        "ALTER TABLE microwave ADD COLUMN custom_capacity TEXT DEFAULT ''",
        "ALTER TABLE microwave ADD COLUMN tx_frequency TEXT DEFAULT ''",
        "ALTER TABLE microwave ADD COLUMN rx_frequency TEXT DEFAULT ''",
        "ALTER TABLE microwave ADD COLUMN received_power TEXT DEFAULT ''",
        "ALTER TABLE microwave ADD COLUMN tx_power TEXT DEFAULT ''",
    ]:
        try:
            cursor.execute(column_sql)
        except sqlite3.OperationalError:
            pass

    # ========================================================
    # E1
    # ========================================================

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS e1_links (

        id INTEGER PRIMARY KEY AUTOINCREMENT,

        microwave_id INTEGER,

        e1_number INTEGER,

        service_name TEXT,
        destination TEXT,

        status TEXT,

        notes TEXT
    )
    """)

    # ========================================================
    # POWER
    # ========================================================

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS power_system (

        id INTEGER PRIMARY KEY AUTOINCREMENT,

        site_id INTEGER,

        electricity_type TEXT,

        generator_type TEXT,
        generator_capacity TEXT,

        battery_type TEXT,
        battery_number TEXT,
        battery_capacity TEXT,

        solar_system TEXT,

        notes TEXT
    )
    """)

    try:
        cursor.execute("ALTER TABLE system_2g ADD COLUMN cards_information TEXT DEFAULT ''")
    except sqlite3.OperationalError:
        pass

    try:
        cursor.execute("ALTER TABLE system_2g_sectors ADD COLUMN custom_band TEXT DEFAULT ''")
    except sqlite3.OperationalError:
        pass

    # Professional custom names for "Other" selections
    for column_sql in [
        "ALTER TABLE system_2g ADD COLUMN custom_vendor_name TEXT DEFAULT ''",
        "ALTER TABLE system_2g ADD COLUMN custom_bts_type TEXT DEFAULT ''",
        "ALTER TABLE system_2g ADD COLUMN custom_band TEXT DEFAULT ''",
    ]:
        try:
            cursor.execute(column_sql)
        except sqlite3.OperationalError:
            pass

    # Cards are assigned per sector, not per 2G system.
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS system_2g_sector_cards (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        sector_id INTEGER NOT NULL,
        card_type TEXT NOT NULL,
        notes TEXT DEFAULT ''
    )
    """)

    # Ensure alarms table exists before any screen or Excel operation.
    ensure_site_alarms_table(connection)

    connection.commit()
    connection.close()



# ============================================================
# SITE ALARMS
# ============================================================

SITE_ALARM_NAMES = [
    "Urgent",
    "Internal",
    "High temperature",
    "Fire alarm",
    "Mains failure",
    "Generator1",
    "Generator2"
]


def ensure_site_alarms_table(connection):
    cursor = connection.cursor()
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS site_alarms (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        site_id INTEGER NOT NULL,
        alarm_name TEXT NOT NULL,
        status TEXT NOT NULL DEFAULT 'Working',
        reason TEXT DEFAULT '',
        UNIQUE(site_id, alarm_name)
    )
    """)
    connection.commit()



# ============================================================
# EXCEL IMPORT / EXPORT - FULL DATABASE
# ============================================================

EXCEL_TABLES = [
    "sites",
    "system_2g",
    "system_2g_cards",
    "system_2g_sectors",
    "system_2g_sector_cards",
    "lte_system",
    "lte_bbu_cards",
    "lte_rru",
    "lte_cells",
    "microwave",
    "e1_links",
    "power_system",
    "site_alarms",
]

# For every table, return only records belonging to the selected Site IDs.
EXCEL_TABLE_QUERIES = {
    "sites": "SELECT * FROM sites WHERE id IN ({ids}) ORDER BY id",
    "system_2g": "SELECT * FROM system_2g WHERE site_id IN ({ids}) ORDER BY id",
    "system_2g_cards": "SELECT t.* FROM system_2g_cards t JOIN system_2g s ON s.id=t.system_2g_id WHERE s.site_id IN ({ids}) ORDER BY t.id",
    "system_2g_sectors": "SELECT t.* FROM system_2g_sectors t JOIN system_2g s ON s.id=t.system_2g_id WHERE s.site_id IN ({ids}) ORDER BY t.id",
    "system_2g_sector_cards": "SELECT t.* FROM system_2g_sector_cards t JOIN system_2g_sectors s ON s.id=t.sector_id JOIN system_2g g ON g.id=s.system_2g_id WHERE g.site_id IN ({ids}) ORDER BY t.id",
    "lte_system": "SELECT * FROM lte_system WHERE site_id IN ({ids}) ORDER BY id",
    "lte_bbu_cards": "SELECT t.* FROM lte_bbu_cards t JOIN lte_system s ON s.id=t.lte_system_id WHERE s.site_id IN ({ids}) ORDER BY t.id",
    "lte_rru": "SELECT t.* FROM lte_rru t JOIN lte_system s ON s.id=t.lte_system_id WHERE s.site_id IN ({ids}) ORDER BY t.id",
    "lte_cells": "SELECT t.* FROM lte_cells t JOIN lte_system s ON s.id=t.lte_system_id WHERE s.site_id IN ({ids}) ORDER BY t.id",
    "microwave": "SELECT * FROM microwave WHERE site_id IN ({ids}) ORDER BY id",
    "e1_links": "SELECT t.* FROM e1_links t JOIN microwave m ON m.id=t.microwave_id WHERE m.site_id IN ({ids}) ORDER BY t.id",
    "power_system": "SELECT * FROM power_system WHERE site_id IN ({ids}) ORDER BY id",
    "site_alarms": "SELECT * FROM site_alarms WHERE site_id IN ({ids}) ORDER BY id",
}


def _excel_ready():
    if Workbook is None or load_workbook is None:
        show_message("Excel Module Missing", "openpyxl is not installed. Install it with:\n\npip install openpyxl")
        return False
    return True


def _style_excel_sheet(sheet):
    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    sheet.freeze_panes = "A2"
    if sheet.max_row >= 1 and sheet.max_column >= 1:
        sheet.auto_filter.ref = sheet.dimensions
    sheet.row_dimensions[1].height = 24
    for col_index in range(1, sheet.max_column + 1):
        max_length = 12
        for row_index in range(1, min(sheet.max_row, 300) + 1):
            value = sheet.cell(row=row_index, column=col_index).value
            if value is not None:
                max_length = max(max_length, len(str(value)))
        sheet.column_dimensions[get_column_letter(col_index)].width = min(max(max_length + 3, 14), 40)


def _safe_excel_value(value):
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d %H:%M:%S")
    return value


def _table_columns(cursor, table):
    cursor.execute(f'PRAGMA table_info("{table}")')
    return [row[1] for row in cursor.fetchall()]


def _build_full_export_workbook(site_ids):
    if not site_ids:
        raise ValueError("No sites selected for export.")

    connection = get_connection()
    cursor = connection.cursor()
    try:
        ensure_site_alarms_table(connection)
        ids = [int(x) for x in site_ids]
        ids_sql = ",".join("?" for _ in ids)

        cursor.execute(f"SELECT id FROM sites WHERE id IN ({ids_sql}) ORDER BY id", ids)
        existing_ids = [row[0] for row in cursor.fetchall()]
        if not existing_ids:
            raise ValueError("The selected sites do not exist in the database.")

        workbook = Workbook()
        workbook.remove(workbook.active)

        info = workbook.create_sheet("Export_Info")
        info.append(["Field", "Value"])
        info.append(["Export Date", datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
        info.append(["Export Type", "Full Site Database"])
        info.append(["Selected Site IDs", ", ".join(map(str, existing_ids))])
        _style_excel_sheet(info)

        for table in EXCEL_TABLES:
            sheet = workbook.create_sheet(table[:31])
            query = EXCEL_TABLE_QUERIES[table].format(ids=ids_sql)
            cursor.execute(query, ids)
            rows = cursor.fetchall()
            columns = [d[0] for d in cursor.description]
            sheet.append(columns)
            for row in rows:
                sheet.append([_safe_excel_value(v) for v in row])
            _style_excel_sheet(sheet)

        return workbook
    finally:
        connection.close()


def _download_directory():
    candidates = [
        "/storage/emulated/0/Download",
        os.path.join(os.path.expanduser("~"), "Download"),
        os.getcwd(),
    ]
    for folder in candidates:
        try:
            os.makedirs(folder, exist_ok=True)
            if os.access(folder, os.W_OK):
                return folder
        except Exception:
            pass
    return os.getcwd()


def _write_workbook_to_download(workbook, filename):
    folder = _download_directory()
    path = os.path.join(folder, filename)
    workbook.save(path)
    return path


def export_sites_to_excel():
    """Export ALL sites or selected sites directly to the Android Download folder."""
    if not _excel_ready():
        return

    connection = get_connection()
    cursor = connection.cursor()
    try:
        ensure_site_alarms_table(connection)
        cursor.execute("SELECT id, site_code, site_name, location, region FROM sites ORDER BY site_name COLLATE NOCASE")
        sites = cursor.fetchall()
    finally:
        connection.close()

    if not sites:
        show_message("Export Excel", "There are no sites in the database.")
        return

    content = BoxLayout(orientation="vertical", padding=dp(10), spacing=dp(8))
    top = BoxLayout(size_hint_y=None, height=dp(52), spacing=dp(8))
    all_checkbox = CheckBox(active=False, size_hint_x=None, width=dp(50))
    top.add_widget(all_checkbox)
    top.add_widget(Label(text="EXPORT ALL SITES", font_size=dp(17)))
    content.add_widget(top)

    scroll = ScrollView(do_scroll_x=False, do_scroll_y=True)
    site_list = BoxLayout(orientation="vertical", spacing=dp(5), size_hint_y=None)
    site_list.bind(minimum_height=site_list.setter("height"))
    checks = []
    for site_id, site_code, site_name, location, region in sites:
        row = BoxLayout(size_hint_y=None, height=dp(52), spacing=dp(6))
        cb = CheckBox(active=False, size_hint_x=None, width=dp(50))
        label = Label(text=f"{site_name or 'Unnamed Site'} | ID: {site_code or site_id} | {location or '-'} | {region or '-'}", font_size=dp(14), halign="left", valign="middle")
        label.bind(size=lambda i, v: setattr(i, "text_size", v))
        row.add_widget(cb); row.add_widget(label); site_list.add_widget(row)
        checks.append((site_id, cb))
    scroll.add_widget(site_list); content.add_widget(scroll)

    buttons = BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(8))
    export_button = Button(text="EXPORT", font_size=dp(16))
    cancel_button = Button(text="CANCEL", font_size=dp(16))
    buttons.add_widget(export_button); buttons.add_widget(cancel_button); content.add_widget(buttons)
    popup = Popup(title="EXPORT FULL DATABASE", content=content, size_hint=(0.96, 0.90), auto_dismiss=False)

    def toggle_all(_instance, value):
        for _, cb in checks:
            cb.active = value

    def export_selected(_):
        selected = [site_id for site_id, cb in checks if cb.active] if not all_checkbox.active else [site_id for site_id, _ in checks]
        if not selected:
            show_message("Export Excel", "Please select at least one site, or choose EXPORT ALL SITES.")
            return
        try:
            workbook = _build_full_export_workbook(selected)
            filename = "Telecom_Full_Database_{}.xlsx".format(datetime.now().strftime("%Y%m%d_%H%M%S"))
            path = _write_workbook_to_download(workbook, filename)
            popup.dismiss()
            show_message("Export Complete", "Full database exported successfully.\n\nFile:\n" + path)
        except Exception as error:
            show_message("Export Error", "Could not export Excel.\n\n" + repr(error))

    all_checkbox.bind(active=toggle_all)
    export_button.bind(on_release=export_selected)
    cancel_button.bind(on_release=lambda x: popup.dismiss())
    popup.open()


def _normalize_import_path(path):
    """Return a normal filesystem path. Android content:// URIs are copied to a temp file."""
    if not path:
        raise ValueError("No Excel file was selected.")
    if isinstance(path, (list, tuple)):
        path = path[0] if path else ""
    path = str(path)
    if not path.startswith("content://"):
        return path, None

    try:
        from jnius import autoclass
        PythonActivity = autoclass("org.kivy.android.PythonActivity")
        ContentResolver = PythonActivity.mActivity.getContentResolver()
        Uri = autoclass("android.net.Uri")
        uri = Uri.parse(path)
        stream = ContentResolver.openInputStream(uri)
        if stream is None:
            raise IOError("Android could not open the selected file.")
        fd, temp_path = tempfile.mkstemp(suffix=".xlsx")
        os.close(fd)
        output = open(temp_path, "wb")
        try:
            # PyJNIus/Android has inconsistent support for InputStream.read(bytearray).
            # Use the Java InputStream one-byte read method here; it is slower but
            # reliable for the relatively small Excel database files used by this app.
            chunk = bytearray()
            while True:
                value = stream.read()
                if value == -1:
                    break
                chunk.append(value & 0xFF)
                if len(chunk) >= 65536:
                    output.write(chunk)
                    chunk.clear()
            if chunk:
                output.write(chunk)
        finally:
            output.close()
            stream.close()
        return temp_path, temp_path
    except Exception as error:
        raise IOError("Android returned a content URI and it could not be read. Please select the Excel file from Download again.\n\n" + str(error))


def _list_excel_files():
    """Find Excel workbooks without using Android/Plyer ActivityResult callbacks.
    This avoids the Android/PyJNIus ActivityResultListener visibility error.
    """
    folders = [
        "/storage/emulated/0/Download",
        "/storage/emulated/0/Documents",
    ]
    found = []
    seen = set()
    for folder in folders:
        try:
            if not os.path.isdir(folder):
                continue
            for name in os.listdir(folder):
                if not name.lower().endswith(".xlsx"):
                    continue
                path = os.path.join(folder, name)
                if os.path.isfile(path) and path not in seen:
                    found.append(path)
                    seen.add(path)
        except Exception:
            pass
    found.sort(key=lambda x: os.path.getmtime(x), reverse=True)
    return found


def _show_excel_file_list():
    """Show a native Kivy list of Excel files from Android shared storage.
    No Plyer filechooser and no Android ActivityResultListener are used.
    """
    files = _list_excel_files()
    if not files:
        show_message(
            "Import Excel",
            "No .xlsx files were found in:\n\n"
            "/storage/emulated/0/Download\n"
            "/storage/emulated/0/Documents\n\n"
            "Please copy the Excel file to Download and try again."
        )
        return

    content = BoxLayout(orientation="vertical", padding=dp(10), spacing=dp(8))
    title = Label(
        text="Select Excel file from Download / Documents",
        size_hint_y=None, height=dp(45), font_size=dp(16)
    )
    content.add_widget(title)

    scroll = ScrollView(do_scroll_x=False, do_scroll_y=True)
    file_list = BoxLayout(orientation="vertical", spacing=dp(7), size_hint_y=None)
    file_list.bind(minimum_height=file_list.setter("height"))

    popup = Popup(
        title="IMPORT FULL DATABASE",
        content=content,
        size_hint=(0.96, 0.90),
        auto_dismiss=False
    )

    for path in files:
        folder = os.path.basename(os.path.dirname(path))
        name = os.path.basename(path)
        try:
            size_mb = os.path.getsize(path) / (1024 * 1024)
            text = f"{name}\n{folder}  |  {size_mb:.2f} MB"
        except Exception:
            text = f"{name}\n{folder}"

        button = Button(
            text=text,
            size_hint_y=None,
            height=dp(68),
            font_size=dp(14),
            halign="left",
            valign="middle"
        )
        button.bind(size=lambda inst, value: setattr(inst, "text_size", (value[0] - dp(20), None)))
        button.bind(on_release=lambda _btn, selected_path=path: _import_selected_file_from_list(selected_path, popup))
        file_list.add_widget(button)

    scroll.add_widget(file_list)
    content.add_widget(scroll)

    cancel = Button(text="CANCEL", size_hint_y=None, height=dp(55), font_size=dp(16))
    cancel.bind(on_release=lambda *_: popup.dismiss())
    content.add_widget(cancel)
    popup.open()


def _import_selected_file_from_list(path, popup):
    try:
        popup.dismiss()
        result = import_sites_from_excel_file(path)
        show_message("Import Complete", result)
        app = App.get_running_app()
        if app and app.root:
            app.root.current = "sites"
    except Exception as error:
        show_message("Import Error", "Could not import Excel.\n\n" + repr(error))


def import_sites_from_excel():
    if not _excel_ready():
        return

    # IMPORTANT: Do not call Plyer filechooser on Android/Pydroid3 here.
    # Its ActivityResultListener integration can raise:
    # java.lang.IllegalArgumentException: interface
    # org.kivy.android.PythonActivity$ActivityResultListener is not visible...
    # Instead, display Excel files from shared storage using a Kivy popup.
    _show_excel_file_list()


def _import_excel_selection(selection):
    if not selection:
        return
    try:
        result = import_sites_from_excel_file(selection[0])
        show_message("Import Complete", result)
        app = App.get_running_app()
        if app and app.root:
            app.root.current = "sites"
    except Exception as error:
        show_message("Import Error", "Could not import Excel.\n\n" + repr(error))


def import_sites_from_excel_file(path):
    """Import all supported sheets while ignoring Excel columns not present in the current DB."""
    local_path, temp_path = _normalize_import_path(path)
    try:
        workbook = load_workbook(local_path, data_only=True, read_only=True)
        sheet_map = {str(name).strip().lower(): name for name in workbook.sheetnames}
        if "sites" not in sheet_map:
            raise ValueError("Invalid Excel file: the Sites sheet is missing.")

        connection = get_connection()
        cursor = connection.cursor()
        imported = {}
        table_order = ["sites", "system_2g", "system_2g_cards", "system_2g_sectors", "system_2g_sector_cards", "lte_system", "lte_bbu_cards", "lte_rru", "lte_cells", "microwave", "e1_links", "power_system", "site_alarms"]
        try:
            ensure_site_alarms_table(connection)
            connection.execute("BEGIN")
            for table in table_order:
                sheet_key = table.lower()
                if sheet_key not in sheet_map:
                    continue
                sheet = workbook[sheet_map[sheet_key]]
                rows = sheet.iter_rows(values_only=True)
                try:
                    raw_headers = next(rows)
                except StopIteration:
                    continue
                headers = [str(h).strip() if h is not None else "" for h in raw_headers]
                db_columns = _table_columns(cursor, table)
                db_map = {c.lower(): c for c in db_columns}
                usable = [(idx, db_map[h.lower()]) for idx, h in enumerate(headers) if h and h.lower() in db_map]
                if not usable:
                    continue
                if "id" not in [col.lower() for _, col in usable]:
                    raise ValueError(f"Sheet '{sheet.title}' is missing the id column.")
                columns = [col for _, col in usable]
                quoted = ",".join('"' + c.replace('"','""') + '"' for c in columns)
                placeholders = ",".join("?" for _ in columns)
                sql = f'INSERT OR REPLACE INTO "{table}" ({quoted}) VALUES ({placeholders})'
                count = 0
                for excel_row_number, excel_row in enumerate(rows, start=2):
                    if not any(v is not None and str(v).strip() != "" for v in excel_row):
                        continue
                    values = [_safe_excel_value(excel_row[idx]) if idx < len(excel_row) else "" for idx, _ in usable]
                    try:
                        cursor.execute(sql, values)
                    except Exception as row_error:
                        raise ValueError(
                            f"Import failed in sheet '{sheet.title}', row {excel_row_number}.\n"
                            f"SQLite error: {row_error}"
                        ) from row_error
                    count += 1
                if count:
                    imported[table] = count
            connection.commit()
        except Exception:
            connection.rollback()
            raise
        finally:
            connection.close()
            workbook.close()

        summary = "Excel imported successfully.\n\n" + "\n".join(f"{table}: {count} rows" for table, count in imported.items())
        return summary
    finally:
        if temp_path:
            try: os.remove(temp_path)
            except Exception: pass

# ============================================================
# UI HELPERS
# ============================================================

def create_input(multiline=False, height=50):

    return TextInput(
        multiline=multiline,
        size_hint_y=None,
        height=height,
        font_size=18,
        padding=[10, 10]
    )


def create_spinner(values):

    return Spinner(
        text="Select",
        values=values,
        size_hint_y=None,
        height=52,
        font_size=17,
    )


def add_field(layout, title, widget):

    label = Label(
        text=title,
        size_hint_y=None,
        height=32,
        halign="left",
        valign="middle",
        font_size=17
    )

    label.bind(
        size=lambda instance, value:
        setattr(instance, "text_size", value)
    )

    layout.add_widget(label)
    layout.add_widget(widget)


def show_message(title, message):

    layout = BoxLayout(
        orientation="vertical",
        padding=20,
        spacing=15
    )

    label = Label(
        text=message,
        font_size=17
    )

    close_button = Button(
        text="OK",
        size_hint_y=None,
        height=50
    )

    layout.add_widget(label)
    layout.add_widget(close_button)

    popup = Popup(
        title=title,
        content=layout,
        size_hint=(0.88, 0.42),
        auto_dismiss=False
    )

    close_button.bind(
        on_release=popup.dismiss
    )

    popup.open()


def confirm_delete(title, message, callback):

    layout = BoxLayout(
        orientation="vertical",
        padding=20,
        spacing=15
    )

    layout.add_widget(
        Label(
            text=message,
            font_size=17
        )
    )

    buttons = BoxLayout(
        size_hint_y=None,
        height=55,
        spacing=10
    )

    yes = Button(text="YES")
    no = Button(text="NO")

    buttons.add_widget(yes)
    buttons.add_widget(no)

    layout.add_widget(buttons)

    popup = Popup(
        title=title,
        content=layout,
        size_hint=(0.88, 0.42),
        auto_dismiss=False
    )

    yes.bind(
        on_release=lambda x:
        (
            popup.dismiss(),
            callback()
        )
    )

    no.bind(
        on_release=popup.dismiss
    )

    popup.open()


# ============================================================
# DASHBOARD CARD
# ============================================================

def dashboard_card(title, value, status, callback, subtitle=""):
    """Professional, fixed-height dashboard card designed for mobile screens."""
    card = BoxLayout(
        orientation="vertical",
        padding=[dp(12), dp(10)],
        spacing=dp(5),
        size_hint_y=None,
        height=dp(225)
    )

    title_label = Label(
        text=title,
        font_size=dp(18),
        bold=True,
        size_hint_y=None,
        height=dp(30),
        halign="left",
        valign="middle"
    )
    title_label.bind(size=lambda i, v: setattr(i, "text_size", v))

    value_label = Label(
        text=value,
        font_size=dp(30),
        bold=True,
        size_hint_y=None,
        height=dp(45),
        halign="left",
        valign="middle"
    )
    value_label.bind(size=lambda i, v: setattr(i, "text_size", v))

    status_label = Label(
        text=status,
        font_size=dp(14),
        size_hint_y=None,
        height=dp(42),
        halign="left",
        valign="middle"
    )
    status_label.bind(size=lambda i, v: setattr(i, "text_size", v))

    if subtitle:
        sub = Label(
            text=subtitle,
            font_size=dp(12),
            size_hint_y=None,
            height=dp(24),
            halign="left",
            valign="middle"
        )
        sub.bind(size=lambda i, v: setattr(i, "text_size", v))
        card.add_widget(sub)

    open_button = Button(
        text="OPEN",
        size_hint_y=None,
        height=dp(42),
        font_size=dp(15)
    )
    open_button.bind(on_release=callback)

    card.add_widget(title_label)
    card.add_widget(value_label)
    card.add_widget(status_label)
    card.add_widget(open_button)
    return card


# ============================================================
# PROFESSIONAL SITE SYSTEMS DASHBOARD
# ============================================================

class SystemsScreen(Screen):

    def on_pre_enter(self):
        self.build_screen()

    def build_screen(self):
        self.clear_widgets()

        if CURRENT_SITE_ID is None:
            show_message("No Site", "Please select a site first.")
            self.manager.current = "sites"
            return

        connection = get_connection()
        cursor = connection.cursor()
        cursor.execute(
            "SELECT site_code, site_name FROM sites WHERE id=?",
            (CURRENT_SITE_ID,)
        )
        site = cursor.fetchone()
        connection.close()

        if not site:
            show_message("Site Error", "Selected site was not found in database.")
            self.manager.current = "sites"
            return

        main = BoxLayout(orientation="vertical", padding=dp(12), spacing=dp(10))

        header = BoxLayout(orientation="vertical", size_hint_y=None, height=dp(82), spacing=dp(2))
        title = Label(text="SITE SYSTEMS DASHBOARD", font_size=dp(22), bold=True,
                      size_hint_y=None, height=dp(38), halign="center", valign="middle")
        title.bind(size=lambda i, v: setattr(i, "text_size", v))
        site_label = Label(text=f"{site[1] or 'Unnamed Site'}  |  Site ID: {site[0] or '-'}",
                           font_size=dp(16), size_hint_y=None, height=dp(38),
                           halign="center", valign="middle")
        site_label.bind(size=lambda i, v: setattr(i, "text_size", v))
        header.add_widget(title)
        header.add_widget(site_label)

        scroll = ScrollView(do_scroll_x=False, do_scroll_y=True)
        grid = GridLayout(cols=2, spacing=dp(12), padding=dp(6), size_hint_y=None)
        grid.bind(minimum_height=grid.setter("height"))

        def system_card(text, screen_name):
            button = Button(text=text, size_hint_y=None, height=dp(105), font_size=dp(18), bold=True)
            button.bind(on_release=lambda x, name=screen_name: setattr(self.manager, "current", name))
            return button

        grid.add_widget(system_card("2G SYSTEM", "2g_list"))
        grid.add_widget(system_card("LTE SYSTEM", "lte_list"))
        grid.add_widget(system_card("MICROWAVE SYSTEM", "microwave_list"))
        grid.add_widget(system_card("POWER SYSTEM", "power"))
        grid.add_widget(system_card("ALARM SYSTEM", "site_alarms"))

        scroll.add_widget(grid)

        back = Button(text="BACK TO SITE", size_hint_y=None, height=dp(58), font_size=dp(17))
        back.bind(on_release=lambda x: setattr(self.manager, "current", "site_form"))

        main.add_widget(header)
        main.add_widget(scroll)
        main.add_widget(back)
        self.add_widget(main)


# ============================================================
# SITE ALARMS SCREEN
# ============================================================

class SiteAlarmsScreen(Screen):

    def on_pre_enter(self):
        self.build_screen()

    def build_screen(self):
        self.clear_widgets()

        if CURRENT_SITE_ID is None:
            show_message("No Site", "Please select a site first.")
            self.manager.current = "sites"
            return

        connection = get_connection()
        ensure_site_alarms_table(connection)
        cursor = connection.cursor()

        cursor.execute(
            "SELECT site_name, site_code FROM sites WHERE id=?",
            (CURRENT_SITE_ID,)
        )
        site = cursor.fetchone()

        if not site:
            connection.close()
            show_message("Site Error", "Selected site was not found.")
            self.manager.current = "sites"
            return

        for alarm_name in SITE_ALARM_NAMES:
            cursor.execute("""
                INSERT OR IGNORE INTO site_alarms
                (site_id, alarm_name, status, reason)
                VALUES (?, ?, 'Working', '')
            """, (CURRENT_SITE_ID, alarm_name))

        connection.commit()

        cursor.execute("""
            SELECT alarm_name, status, reason
            FROM site_alarms
            WHERE site_id=?
            ORDER BY id
        """, (CURRENT_SITE_ID,))
        rows = cursor.fetchall()
        connection.close()

        main = BoxLayout(
            orientation="vertical",
            padding=dp(10),
            spacing=dp(8)
        )

        title = Label(
            text=f"SITE ALARMS\n{site[0] or 'Unnamed Site'}  |  {site[1] or '-'}",
            font_size=dp(21),
            bold=True,
            size_hint_y=None,
            height=dp(75),
            halign="left",
            valign="middle"
        )
        title.bind(size=lambda i, v: setattr(i, "text_size", v))

        scroll = ScrollView(do_scroll_x=False, do_scroll_y=True)
        content = BoxLayout(
            orientation="vertical",
            spacing=dp(10),
            padding=[dp(4), dp(4)],
            size_hint_y=None
        )
        content.bind(minimum_height=content.setter("height"))

        self.alarm_widgets = {}

        for alarm_name, status, reason in rows:
            card = BoxLayout(
                orientation="vertical",
                size_hint_y=None,
                height=dp(205),
                padding=[dp(10), dp(8)],
                spacing=dp(5)
            )

            alarm_label = Label(
                text=alarm_name,
                font_size=dp(18),
                bold=True,
                size_hint_y=None,
                height=dp(30),
                halign="left",
                valign="middle"
            )
            alarm_label.bind(
                size=lambda i, v: setattr(i, "text_size", v)
            )

            status_spinner = Spinner(
                text=status if status in ("Working", "Not Working") else "Working",
                values=("Working", "Not Working"),
                size_hint_y=None,
                height=dp(45),
                font_size=dp(17)
            )

            reason_input = TextInput(
                text=reason or "",
                hint_text="Reason / note if not working",
                multiline=False,
                size_hint_y=None,
                height=dp(48),
                font_size=dp(16),
                padding=[dp(10), dp(10)]
            )

            self.alarm_widgets[alarm_name] = (
                status_spinner,
                reason_input
            )

            delete_button = Button(
                text="DELETE ALARM",
                size_hint_y=None,
                height=dp(42),
                font_size=dp(15)
            )
            delete_button.bind(
                on_release=lambda x, name=alarm_name: self.confirm_delete_alarm(name)
            )

            card.add_widget(alarm_label)
            card.add_widget(status_spinner)
            card.add_widget(reason_input)
            card.add_widget(delete_button)
            content.add_widget(card)

        scroll.add_widget(content)

        add_alarm_button = Button(
            text="+ ADD CUSTOM ALARM",
            size_hint_y=None,
            height=dp(55),
            font_size=dp(17)
        )
        add_alarm_button.bind(on_release=self.add_custom_alarm)

        save_button = Button(
            text="SAVE ALARMS",
            size_hint_y=None,
            height=dp(58),
            font_size=dp(18)
        )
        save_button.bind(on_release=self.save_alarms)

        back_button = Button(
            text="BACK TO SITE DASHBOARD",
            size_hint_y=None,
            height=dp(55),
            font_size=dp(16)
        )
        back_button.bind(
            on_release=lambda x:
            setattr(self.manager, "current", "systems")
        )

        main.add_widget(title)
        main.add_widget(scroll)
        main.add_widget(add_alarm_button)
        main.add_widget(save_button)
        main.add_widget(back_button)

        self.add_widget(main)

    def add_custom_alarm(self, instance):
        content = BoxLayout(orientation="vertical", padding=dp(12), spacing=dp(10))

        alarm_input = TextInput(
            hint_text="Enter alarm name",
            multiline=False,
            size_hint_y=None,
            height=dp(50),
            font_size=dp(17)
        )

        buttons = BoxLayout(size_hint_y=None, height=dp(50), spacing=dp(8))
        add_button = Button(text="ADD", font_size=dp(16))
        cancel_button = Button(text="CANCEL", font_size=dp(16))
        buttons.add_widget(add_button)
        buttons.add_widget(cancel_button)
        content.add_widget(alarm_input)
        content.add_widget(buttons)

        popup = Popup(
            title="ADD CUSTOM ALARM",
            content=content,
            size_hint=(0.92, None),
            height=dp(190),
            auto_dismiss=False
        )

        def add_alarm(_):
            alarm_name = alarm_input.text.strip()
            if not alarm_name:
                show_message("Invalid Alarm", "Please enter an alarm name.")
                return

            connection = get_connection()
            ensure_site_alarms_table(connection)
            cursor = connection.cursor()
            cursor.execute("""
                INSERT OR IGNORE INTO site_alarms (site_id, alarm_name, status, reason)
                VALUES (?, ?, 'Working', '')
            """, (CURRENT_SITE_ID, alarm_name))
            connection.commit()
            connection.close()
            popup.dismiss()
            self.build_screen()

        add_button.bind(on_release=add_alarm)
        cancel_button.bind(on_release=lambda x: popup.dismiss())
        popup.open()


    def confirm_delete_alarm(self, alarm_name):
        content = BoxLayout(
            orientation="vertical",
            padding=dp(12),
            spacing=dp(10)
        )

        message = Label(
            text=f"Are you sure you want to delete this alarm?\n\n{alarm_name}",
            font_size=dp(16),
            halign="center",
            valign="middle"
        )
        message.bind(size=lambda i, v: setattr(i, "text_size", v))

        buttons = BoxLayout(
            size_hint_y=None,
            height=dp(50),
            spacing=dp(8)
        )
        delete_button = Button(text="DELETE", font_size=dp(16))
        cancel_button = Button(text="CANCEL", font_size=dp(16))
        buttons.add_widget(delete_button)
        buttons.add_widget(cancel_button)
        content.add_widget(message)
        content.add_widget(buttons)

        popup = Popup(
            title="DELETE ALARM",
            content=content,
            size_hint=(0.92, None),
            height=dp(210),
            auto_dismiss=False
        )

        def delete_alarm(_):
            connection = get_connection()
            ensure_site_alarms_table(connection)
            cursor = connection.cursor()
            cursor.execute(
                "DELETE FROM site_alarms WHERE site_id=? AND alarm_name=?",
                (CURRENT_SITE_ID, alarm_name)
            )
            connection.commit()
            connection.close()
            popup.dismiss()
            self.build_screen()

        delete_button.bind(on_release=delete_alarm)
        cancel_button.bind(on_release=lambda x: popup.dismiss())
        popup.open()


    def save_alarms(self, instance):
        if CURRENT_SITE_ID is None:
            return

        connection = get_connection()
        ensure_site_alarms_table(connection)
        cursor = connection.cursor()

        for alarm_name, widgets in self.alarm_widgets.items():
            status_spinner, reason_input = widgets
            cursor.execute("""
                UPDATE site_alarms
                SET status=?, reason=?
                WHERE site_id=? AND alarm_name=?
            """, (
                status_spinner.text,
                reason_input.text.strip(),
                CURRENT_SITE_ID,
                alarm_name
            ))

        connection.commit()
        connection.close()

        show_message(
            "Saved",
            "Site alarm status and notes have been saved successfully."
        )



# ============================================================
# HOME
# ============================================================

class HomeScreen(Screen):

    def on_pre_enter(self):
        self.build_screen()

    def build_screen(self):

        self.clear_widgets()

        layout = BoxLayout(
            orientation="vertical",
            padding=30,
            spacing=20
        )

        title = Label(
            text="TELECOM SITE DATABASE",
            font_size=28,
            size_hint_y=None,
            height=70
        )

        subtitle = Label(
            text="Professional Telecommunication Site Management System",
            font_size=18,
            size_hint_y=None,
            height=50
        )

        site_button = Button(
            text="SITE DATABASE",
            size_hint_y=None,
            height=65,
            font_size=19
        )

        site_button.bind(
            on_release=lambda x:
            setattr(
                self.manager,
                "current",
                "sites"
            )
        )

        layout.add_widget(title)
        layout.add_widget(subtitle)
        layout.add_widget(site_button)

        self.add_widget(layout)


# ============================================================
# SITES LIST
# ============================================================

class SitesScreen(Screen):

    def on_pre_enter(self):
        self.build_screen()

    def build_screen(self):
        self.clear_widgets()

        main = BoxLayout(orientation="vertical", padding=dp(10), spacing=dp(10))

        title = Label(text="SITE DATABASE", font_size=dp(25), bold=True,
                      size_hint_y=None, height=dp(55))

        # Database summary
        summary = GridLayout(cols=2, size_hint_y=None, height=dp(70), spacing=dp(8))
        self.sites_count_label = Label(text="SITES\n0", font_size=dp(18), bold=True,
                                       halign="center", valign="middle")
        self.regions_count_label = Label(text="REGIONS\n0", font_size=dp(18), bold=True,
                                         halign="center", valign="middle")
        self.sites_count_label.bind(size=lambda i, v: setattr(i, "text_size", v))
        self.regions_count_label.bind(size=lambda i, v: setattr(i, "text_size", v))
        summary.add_widget(self.sites_count_label)
        summary.add_widget(self.regions_count_label)

        self.search_input = create_input()
        self.search_input.hint_text = "Search Site ID, Name, Location or Region"
        self.search_input.bind(text=lambda instance, value: self.display_sites())

        self.region_filter = Spinner(
            text="ALL REGIONS",
            values=("ALL REGIONS",),
            size_hint_y=None,
            height=dp(52),
            font_size=dp(16)
        )
        self.region_filter.bind(text=lambda instance, value: self.display_sites())

        add_button = Button(text="ADD NEW SITE", size_hint_y=None, height=dp(58), font_size=dp(18))
        add_button.bind(on_release=self.add_new_site)

        excel_buttons = BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(8))
        export_button = Button(text="EXPORT EXCEL", font_size=dp(16))
        export_button.bind(on_release=lambda x: export_sites_to_excel())
        import_button = Button(text="IMPORT EXCEL", font_size=dp(16))
        import_button.bind(on_release=lambda x: import_sites_from_excel())
        excel_buttons.add_widget(export_button)
        excel_buttons.add_widget(import_button)

        scroll = ScrollView(do_scroll_x=False, do_scroll_y=True)
        self.sites_layout = BoxLayout(orientation="vertical", spacing=dp(15), padding=[dp(5), dp(5)], size_hint_y=None)
        self.sites_layout.bind(minimum_height=self.sites_layout.setter("height"))
        scroll.add_widget(self.sites_layout)

        back_button = Button(text="BACK TO HOME", size_hint_y=None, height=dp(58))
        back_button.bind(on_release=lambda x: setattr(self.manager, "current", "home"))

        main.add_widget(title)
        main.add_widget(summary)
        main.add_widget(self.search_input)
        main.add_widget(self.region_filter)
        main.add_widget(add_button)
        main.add_widget(excel_buttons)
        main.add_widget(scroll)
        main.add_widget(back_button)
        self.add_widget(main)

        self.refresh_filters_and_stats()

    def refresh_filters_and_stats(self):
        connection = get_connection()
        cursor = connection.cursor()
        cursor.execute("SELECT COUNT(*) FROM sites")
        site_count = cursor.fetchone()[0]
        cursor.execute("""
            SELECT COUNT(DISTINCT TRIM(region))
            FROM sites
            WHERE TRIM(COALESCE(region, '')) <> ''
        """)
        region_count = cursor.fetchone()[0]
        cursor.execute("""
            SELECT DISTINCT TRIM(region)
            FROM sites
            WHERE TRIM(COALESCE(region, '')) <> ''
            ORDER BY TRIM(region) COLLATE NOCASE
        """)
        regions = [row[0] for row in cursor.fetchall()]
        connection.close()

        self.sites_count_label.text = f"SITES\n{site_count}"
        self.regions_count_label.text = f"REGIONS\n{region_count}"

        current = self.region_filter.text
        values = ["ALL REGIONS"] + regions
        self.region_filter.values = values
        if current in values:
            self.region_filter.text = current
        else:
            self.region_filter.text = "ALL REGIONS"

        self.display_sites()

    def display_sites(self, *args):
        self.sites_layout.clear_widgets()

        search_text = self.search_input.text.strip()
        search = f"%{search_text}%"
        region = self.region_filter.text

        connection = get_connection()
        cursor = connection.cursor()

        base_query = """
        SELECT id, site_code, site_name, location, region, environment, structure_type
        FROM sites
        WHERE (site_code LIKE ? OR site_name LIKE ? OR location LIKE ? OR region LIKE ?)
        """
        params = [search, search, search, search]
        if region and region != "ALL REGIONS":
            base_query += " AND TRIM(COALESCE(region, '')) = TRIM(?)"
            params.append(region)
        base_query += " ORDER BY site_name COLLATE NOCASE"

        cursor.execute(base_query, params)
        rows = cursor.fetchall()
        connection.close()

        if not rows:
            self.sites_layout.add_widget(Label(text="No Sites Found", size_hint_y=None, height=dp(70), font_size=dp(18)))
            return

        for row in rows:
            site_id = row[0]
            card = BoxLayout(orientation="vertical", size_hint_y=None, height=dp(185), padding=dp(12), spacing=dp(8))
            info = Label(text=(f"[b]{row[2] or 'Unnamed Site'}[/b]\n"
                               f"Site ID: {row[1] or '-'}\n"
                               f"Location: {row[3] or '-'}\n"
                               f"Region: {row[4] or '-'}\n"
                               f"{row[5] or '-'} / {row[6] or '-'}"),
                         markup=True, halign="left", valign="middle", font_size=dp(17))
            info.bind(size=lambda instance, value: setattr(instance, "text_size", value))

            buttons = BoxLayout(size_hint_y=None, height=dp(48), spacing=dp(8))
            details_button = Button(text="VIEW DETAILS", font_size=dp(15))
            details_button.bind(on_release=lambda x, value=site_id: self.view_details(value))
            open_button = Button(text="OPEN / EDIT", font_size=dp(15))
            open_button.bind(on_release=lambda x, value=site_id: self.select_site(value))
            buttons.add_widget(details_button)
            buttons.add_widget(open_button)
            card.add_widget(info)
            card.add_widget(buttons)
            self.sites_layout.add_widget(card)

    def add_new_site(self, instance):
        global CURRENT_SITE_ID, CURRENT_SITE_NAME
        CURRENT_SITE_ID = None
        CURRENT_SITE_NAME = ""
        global CURRENT_2G_ID, CURRENT_2G_SECTOR_ID, CURRENT_LTE_ID, CURRENT_LTE_RRU_ID, CURRENT_LTE_CELL_ID, CURRENT_MICROWAVE_ID
        CURRENT_2G_ID = None
        CURRENT_2G_SECTOR_ID = None
        CURRENT_LTE_ID = None
        CURRENT_LTE_RRU_ID = None
        CURRENT_LTE_CELL_ID = None
        CURRENT_MICROWAVE_ID = None
        self.manager.current = "site_form"

    def select_site(self, site_id):
        global CURRENT_SITE_ID, CURRENT_SITE_NAME
        CURRENT_SITE_ID = site_id
        connection = get_connection()
        cursor = connection.cursor()
        cursor.execute("SELECT site_name FROM sites WHERE id=?", (site_id,))
        row = cursor.fetchone()
        connection.close()
        if row:
            CURRENT_SITE_NAME = row[0] or ""
        self.manager.current = "site_form"

    def view_details(self, site_id):
        global CURRENT_SITE_ID, CURRENT_SITE_NAME
        CURRENT_SITE_ID = site_id
        connection = get_connection()
        cursor = connection.cursor()
        cursor.execute("SELECT site_name FROM sites WHERE id=?", (site_id,))
        row = cursor.fetchone()
        connection.close()
        if not row:
            show_message("Site Error", "Selected site was not found.")
            return
        CURRENT_SITE_NAME = row[0] or ""
        self.manager.current = "site_details"


# ============================================================
# SITE DETAILS
# ============================================================

class SiteDetailsScreen(Screen):
    def on_pre_enter(self):
        self.build_screen()

    def build_screen(self):
        self.clear_widgets()

        if CURRENT_SITE_ID is None:
            show_message("No Site", "Please select a site first.")
            self.manager.current = "sites"
            return

        connection = get_connection()
        cursor = connection.cursor()

        cursor.execute("SELECT * FROM sites WHERE id=?", (CURRENT_SITE_ID,))
        site = cursor.fetchone()

        if not site:
            connection.close()
            show_message("Site Error", "Selected site was not found in database.")
            self.manager.current = "sites"
            return

        counts = {}
        queries = {
            "2G Systems": "SELECT COUNT(*) FROM system_2g WHERE site_id=?",
            "LTE Systems": "SELECT COUNT(*) FROM lte_system WHERE site_id=?",
            "Microwave Links": "SELECT COUNT(*) FROM microwave WHERE site_id=?",
            "Power Records": "SELECT COUNT(*) FROM power_system WHERE site_id=?",
        }

        for key, query in queries.items():
            cursor.execute(query, (CURRENT_SITE_ID,))
            counts[key] = cursor.fetchone()[0]

        connection.close()

        main = BoxLayout(
            orientation="vertical",
            padding=dp(12),
            spacing=dp(10)
        )

        title = Label(
            text=f"SITE DETAILS\n{site[2] or 'Unnamed Site'}",
            font_size=dp(23),
            bold=True,
            size_hint_y=None,
            height=dp(75),
            halign="left",
            valign="middle"
        )
        title.bind(size=lambda i, v: setattr(i, "text_size", v))

        scroll = ScrollView(do_scroll_x=False)
        content = BoxLayout(
            orientation="vertical",
            spacing=dp(8),
            padding=[dp(5), dp(5)],
            size_hint_y=None
        )
        content.bind(minimum_height=content.setter("height"))

        details = [
            ("Site Code", site[1]),
            ("Site Name", site[2]),
            ("Location", site[3]),
            ("Region", site[4]),
            ("Environment", site[5]),
            ("Structure Type", site[6]),
            ("Structure Height", site[7]),
            ("Number of Sectors", site[8]),
            ("Latitude", site[9]),
            ("Longitude", site[10]),
            ("Notes", site[11]),
        ]

        for label_text, value in details:
            card = BoxLayout(
                orientation="vertical",
                size_hint_y=None,
                height=dp(72),
                padding=[dp(10), dp(6)],
                spacing=dp(2)
            )
            label = Label(
                text=label_text,
                font_size=dp(14),
                bold=True,
                size_hint_y=None,
                height=dp(25),
                halign="left"
            )
            value_label = Label(
                text=str(value or "-"),
                font_size=dp(17),
                size_hint_y=None,
                height=dp(38),
                halign="left",
                valign="middle"
            )
            value_label.bind(size=lambda i, v: setattr(i, "text_size", v))
            card.add_widget(label)
            card.add_widget(value_label)
            content.add_widget(card)

        summary = Label(
            text=(
                f"SYSTEM SUMMARY\n"
                f"2G Systems: {counts['2G Systems']}\n"
                f"LTE Systems: {counts['LTE Systems']}\n"
                f"Microwave Links: {counts['Microwave Links']}\n"
                f"Power Records: {counts['Power Records']}"
            ),
            font_size=dp(17),
            size_hint_y=None,
            height=dp(150),
            halign="left",
            valign="middle"
        )
        summary.bind(size=lambda i, v: setattr(i, "text_size", v))
        content.add_widget(summary)

        scroll.add_widget(content)

        edit = Button(
            text="EDIT SITE",
            size_hint_y=None,
            height=dp(58),
            font_size=dp(18)
        )
        edit.bind(on_release=lambda x: setattr(self.manager, "current", "site_form"))

        systems = Button(
            text="SITE SYSTEMS DASHBOARD",
            size_hint_y=None,
            height=dp(58),
            font_size=dp(18)
        )
        systems.bind(on_release=lambda x: setattr(self.manager, "current", "systems"))

        back = Button(
            text="BACK TO SITE DATABASE",
            size_hint_y=None,
            height=dp(55)
        )
        back.bind(on_release=lambda x: setattr(self.manager, "current", "sites"))

        main.add_widget(title)
        main.add_widget(scroll)
        main.add_widget(edit)
        main.add_widget(systems)
        main.add_widget(back)

        self.add_widget(main)


# ============================================================
# SITE FORM
# ============================================================

class SiteFormScreen(Screen):

    def on_pre_enter(self):
        self.build_form()

    def build_form(self):

        self.clear_widgets()

        scroll = ScrollView()

        form = BoxLayout(
            orientation="vertical",
            padding=20,
            spacing=10,
            size_hint_y=None
        )
        self._form = form

        form.bind(
            minimum_height=
            form.setter("height")
        )

        title = Label(
            text="SITE INFORMATION",
            font_size=25,
            size_hint_y=None,
            height=65
        )

        self.site_code = create_input()
        self.site_name = create_input()
        self.location = create_input()
        self.region = create_input()

        self.environment = create_spinner([
            "Indoor",
            "Outdoor"
        ])

        self.structure = create_spinner([
            "Tower",
            "Pole",
            "Rooftop",
            "Building",
            "Other"
        ])

        self.structure_height = create_input()
        self.sectors = create_input()
        self.latitude = create_input()
        self.longitude = create_input()

        self.notes = create_input(
            multiline=True,
            height=110
        )

        add_field(form, "Site ID", self.site_code)
        add_field(form, "Site Name *", self.site_name)
        add_field(form, "Location", self.location)
        add_field(form, "Region", self.region)
        add_field(form, "Environment", self.environment)
        add_field(form, "Structure Type", self.structure)
        add_field(form, "Structure Height", self.structure_height)
        add_field(form, "Number of Sectors", self.sectors)
        add_field(form, "Latitude", self.latitude)
        add_field(form, "Longitude", self.longitude)
        add_field(form, "Notes", self.notes)

        save = Button(
            text="SAVE / UPDATE SITE",
            size_hint_y=None,
            height=60,
            font_size=18
        )

        systems = Button(
            text="SITE SYSTEMS DASHBOARD",
            size_hint_y=None,
            height=60,
            font_size=18
        )

        delete = Button(
            text="DELETE SITE",
            size_hint_y=None,
            height=60
        )

        back = Button(
            text="BACK TO SITE LIST",
            size_hint_y=None,
            height=60
        )

        save.bind(on_release=self.save_site)
        systems.bind(on_release=self.open_systems)
        delete.bind(on_release=self.ask_delete)

        back.bind(
            on_release=lambda x:
            setattr(
                self.manager,
                "current",
                "sites"
            )
        )

        form.add_widget(title)
        form.add_widget(save)
        form.add_widget(systems)
        form.add_widget(delete)
        form.add_widget(back)

        scroll.add_widget(form)
        self.add_widget(scroll)

        self.load_site()

    def load_site(self):

        if CURRENT_SITE_ID is None:
            return

        connection = get_connection()
        cursor = connection.cursor()

        cursor.execute(
            "SELECT * FROM sites WHERE id=?",
            (CURRENT_SITE_ID,)
        )

        row = cursor.fetchone()
        connection.close()

        if not row:
            return

        self.site_code.text = row[1] or ""
        self.site_name.text = row[2] or ""
        self.location.text = row[3] or ""
        self.region.text = row[4] or ""

        self.environment.text = row[5] or "Select"
        self.structure.text = row[6] or "Select"

        self.structure_height.text = row[7] or ""
        self.sectors.text = row[8] or ""

        self.latitude.text = row[9] or ""
        self.longitude.text = row[10] or ""

        self.notes.text = row[11] or ""

    def save_site(self, instance):

        global CURRENT_SITE_ID
        global CURRENT_SITE_NAME

        site_name = self.site_name.text.strip()

        if not site_name:

            show_message(
                "Error",
                "Site Name is required."
            )

            return

        values = (
            self.site_code.text.strip(),
            site_name,
            self.location.text.strip(),
            self.region.text.strip(),
            self.environment.text,
            self.structure.text,
            self.structure_height.text.strip(),
            self.sectors.text.strip(),
            self.latitude.text.strip(),
            self.longitude.text.strip(),
            self.notes.text.strip(),
            self.custom_band.text.strip()
        )

        connection = get_connection()
        cursor = connection.cursor()

        try:

            if CURRENT_SITE_ID is None:

                cursor.execute("""
                INSERT INTO sites (
                    site_code,
                    site_name,
                    location,
                    region,
                    environment,
                    structure_type,
                    structure_height,
                    number_of_sectors,
                    latitude,
                    longitude,
                    notes
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, values)

                CURRENT_SITE_ID = cursor.lastrowid

            else:

                cursor.execute("""
                UPDATE sites SET
                    site_code=?,
                    site_name=?,
                    location=?,
                    region=?,
                    environment=?,
                    structure_type=?,
                    structure_height=?,
                    number_of_sectors=?,
                    latitude=?,
                    longitude=?,
                    notes=?
                WHERE id=?
                """,
                values + (CURRENT_SITE_ID,))

            connection.commit()

        except Exception as error:

            connection.rollback()
            connection.close()

            show_message(
                "Database Error",
                str(error)
            )

            return

        connection.close()

        CURRENT_SITE_NAME = site_name

        show_message(
            "Success",
            "Site saved successfully."
        )

    def open_systems(self, instance):

        if CURRENT_SITE_ID is None:

            show_message(
                "Save Site First",
                "Please save the site before opening systems."
            )

            return

        self.manager.current = "systems"

    def ask_delete(self, instance):

        if CURRENT_SITE_ID is None:

            show_message(
                "Error",
                "No site selected."
            )

            return

        confirm_delete(
            "Delete Site",
            "Delete this site and ALL associated systems?",
            self.delete_site
        )

    def delete_site(self):

        global CURRENT_SITE_ID
        global CURRENT_SITE_NAME

        connection = get_connection()
        cursor = connection.cursor()

        # 2G

        cursor.execute(
            "SELECT id FROM system_2g WHERE site_id=?",
            (CURRENT_SITE_ID,)
        )

        systems = cursor.fetchall()

        for row in systems:

            cursor.execute(
                "DELETE FROM system_2g_sectors WHERE system_2g_id=?",
                (row[0],)
            )
            cursor.execute(
                "DELETE FROM system_2g_cards WHERE system_2g_id=?",
                (row[0],)
            )

        cursor.execute(
            "DELETE FROM system_2g WHERE site_id=?",
            (CURRENT_SITE_ID,)
        )

        # LTE

        cursor.execute(
            "SELECT id FROM lte_system WHERE site_id=?",
            (CURRENT_SITE_ID,)
        )

        lte_systems = cursor.fetchall()

        for row in lte_systems:

            cursor.execute(
                "DELETE FROM lte_rru WHERE lte_system_id=?",
                (row[0],)
            )

            cursor.execute(
                "DELETE FROM lte_cells WHERE lte_system_id=?",
                (row[0],)
            )

        cursor.execute(
            "DELETE FROM lte_system WHERE site_id=?",
            (CURRENT_SITE_ID,)
        )

        # Microwave

        cursor.execute(
            "SELECT id FROM microwave WHERE site_id=?",
            (CURRENT_SITE_ID,)
        )

        links = cursor.fetchall()

        for row in links:

            cursor.execute(
                "DELETE FROM e1_links WHERE microwave_id=?",
                (row[0],)
            )

        cursor.execute(
            "DELETE FROM microwave WHERE site_id=?",
            (CURRENT_SITE_ID,)
        )

        # Power

        cursor.execute(
            "DELETE FROM power_system WHERE site_id=?",
            (CURRENT_SITE_ID,)
        )

        # Site

        cursor.execute(
            "DELETE FROM sites WHERE id=?",
            (CURRENT_SITE_ID,)
        )

        connection.commit()
        connection.close()

        CURRENT_SITE_ID = None
        CURRENT_SITE_NAME = ""

        self.manager.current = "sites"


# ============================================================
# 2G LIST
# ============================================================

class TwoGListScreen(Screen):

    def on_pre_enter(self):
        self.build_screen()

    def build_screen(self):

        self.clear_widgets()

        main = BoxLayout(
            orientation="vertical",
            padding=10,
            spacing=10
        )

        title = Label(
            text=f"2G SYSTEMS\n{CURRENT_SITE_NAME}",
            font_size=23,
            size_hint_y=None,
            height=80
        )

        add_button = Button(
            text="ADD NEW 2G SYSTEM",
            size_hint_y=None,
            height=60,
            font_size=18
        )

        add_button.bind(
            on_release=self.add_2g
        )

        scroll = ScrollView()

        self.systems_layout = BoxLayout(
            orientation="vertical",
            spacing=15,
            padding=[5, 5],
            size_hint_y=None
        )

        self.systems_layout.bind(
            minimum_height=
            self.systems_layout.setter("height")
        )

        scroll.add_widget(self.systems_layout)

        back = Button(
            text="BACK TO SITE SYSTEMS DASHBOARD",
            size_hint_y=None,
            height=60
        )

        back.bind(
            on_release=lambda x:
            setattr(
                self.manager,
                "current",
                "systems"
            )
        )

        main.add_widget(title)
        main.add_widget(add_button)
        main.add_widget(scroll)
        main.add_widget(back)

        self.add_widget(main)

        self.display_2g()

    def display_2g(self):

        self.systems_layout.clear_widgets()

        connection = get_connection()
        cursor = connection.cursor()

        cursor.execute("""
        SELECT
            id,
            bts_vendor,
            bts_type,
            system_band,
            number_of_sectors,
            sector_configuration,
            custom_vendor_name,
            custom_bts_type,
            custom_band
        FROM system_2g
        WHERE site_id=?
        ORDER BY id DESC
        """, (CURRENT_SITE_ID,))

        rows = cursor.fetchall()

        connection.close()

        if not rows:

            self.systems_layout.add_widget(
                Label(
                    text="No 2G Systems Added",
                    size_hint_y=None,
                    height=70,
                    font_size=18
                )
            )

            return

        for row in rows:

            system_id = row[0]

            card = BoxLayout(
                orientation="vertical",
                size_hint_y=None,
                height=205,
                padding=12,
                spacing=8
            )

            vendor_display = row[6] if row[1] == "Other" and row[6] else (row[1] or "-")
            type_display = row[7] if row[2] == "Other" and row[7] else (row[2] or "-")
            band_display = row[8] if row[3] == "Other" and row[8] else (row[3] or "-")

            connection2 = get_connection()
            cursor2 = connection2.cursor()
            cursor2.execute("""
                SELECT COUNT(*) FROM system_2g_sectors
                WHERE system_2g_id=?
            """, (system_id,))
            sector_count = cursor2.fetchone()[0]
            cursor2.execute("""
                SELECT COUNT(*) FROM system_2g_sector_cards c
                JOIN system_2g_sectors s ON s.id=c.sector_id
                WHERE s.system_2g_id=?
            """, (system_id,))
            card_count = cursor2.fetchone()[0]
            connection2.close()

            info = Label(
                text=(
                    f"[b]{type_display}[/b]\n"
                    f"Vendor: {vendor_display}\n"
                    f"System Band: {band_display}\n"
                    f"Sectors: {row[4] or '-'} (Configured: {sector_count})\n"
                    f"Cards Assigned to Sectors: {card_count}\n"
                    f"General Sector Configuration: {row[5] or '-'}"
                ),
                markup=True,
                halign="left",
                valign="middle",
                font_size=17
            )

            info.bind(
                size=lambda instance, value:
                setattr(instance, "text_size", value)
            )

            button = Button(
                text="OPEN / EDIT 2G SYSTEM",
                size_hint_y=None,
                height=50
            )

            button.bind(
                on_release=lambda x,
                value=system_id:
                self.open_2g(value)
            )

            card.add_widget(info)
            card.add_widget(button)

            self.systems_layout.add_widget(card)

    def add_2g(self, instance):

        global CURRENT_2G_ID

        CURRENT_2G_ID = None

        self.manager.current = "2g_form"

    def open_2g(self, system_id):

        global CURRENT_2G_ID

        CURRENT_2G_ID = system_id

        self.manager.current = "2g_form"


# ============================================================
# 2G FORM
# ============================================================

class TwoGFormScreen(Screen):

    def on_pre_enter(self):
        self.build_form()

    def build_form(self):

        self.clear_widgets()

        scroll = ScrollView()

        form = BoxLayout(
            orientation="vertical",
            padding=20,
            spacing=10,
            size_hint_y=None
        )

        form.bind(
            minimum_height=
            form.setter("height")
        )

        title = Label(
            text="2G SYSTEM INFORMATION",
            font_size=24,
            size_hint_y=None,
            height=65
        )

        self.vendor = create_spinner([
            "NOKIA",
            "HUAWEI",
            "Other"
        ])
        self.vendor.bind(text=self.update_bts_options)

        self.custom_vendor = create_input()
        self.custom_vendor.hint_text = "Enter Vendor Name"

        self.bts_type = create_spinner(["Select"])
        self.bts_type.bind(text=self.on_bts_type_changed)

        self.custom_bts_type = create_input()
        self.custom_bts_type.hint_text = "Enter BTS Type"

        self.band = create_spinner([
            "900 MHz",
            "1800 MHz",
            "900 + 1800 MHz",
            "Other"
        ])
        self.band.bind(text=lambda instance, value: self._update_custom_fields())

        self.custom_band = create_input()
        self.custom_band.hint_text = "Enter System Band"

        self.number_of_sectors = create_input()
        self.number_of_sectors.input_filter = "int"

        self.sector_configuration = create_input(multiline=True, height=100)
        self.notes = create_input(multiline=True, height=110)

        form.add_widget(title)

        def simple_field_container(widget, height=82):
            box = BoxLayout(orientation="vertical", size_hint_y=None, height=height, spacing=4)
            box.add_widget(widget)
            return box

        self.custom_vendor_field = simple_field_container(self.custom_vendor)
        self.custom_bts_type_field = simple_field_container(self.custom_bts_type)
        self.custom_band_field = simple_field_container(self.custom_band)

        add_field(form, "BTS Vendor", self.vendor)
        form.add_widget(self.custom_vendor_field)
        add_field(form, "BTS Type", self.bts_type)
        form.add_widget(self.custom_bts_type_field)
        add_field(form, "System Band", self.band)
        form.add_widget(self.custom_band_field)
        add_field(form, "Number of Sectors", self.number_of_sectors)
        add_field(form, "General Sector Configuration", self.sector_configuration)
        add_field(form, "Notes", self.notes)

        save = Button(
            text="SAVE / UPDATE 2G SYSTEM",
            size_hint_y=None,
            height=60,
            font_size=18
        )

        sectors_button = Button(
            text="MANAGE 2G SECTORS / CARDS",
            size_hint_y=None,
            height=60,
            font_size=18
        )

        delete = Button(
            text="DELETE 2G SYSTEM",
            size_hint_y=None,
            height=60
        )

        back = Button(
            text="BACK TO 2G SYSTEM LIST",
            size_hint_y=None,
            height=60
        )

        save.bind(on_release=self.save_2g)
        sectors_button.bind(on_release=self.open_sectors)
        delete.bind(on_release=self.ask_delete)

        back.bind(
            on_release=lambda x:
            setattr(
                self.manager,
                "current",
                "2g_list"
            )
        )

        form.add_widget(save)
        form.add_widget(sectors_button)
        form.add_widget(delete)
        form.add_widget(back)

        scroll.add_widget(form)
        self.add_widget(scroll)

        self.update_bts_options(self.vendor, self.vendor.text)
        self.load_2g()

    def _set_field_visible(self, field, visible):
        if visible:
            if field.parent is None:
                parent = self.bts_type.parent
                if parent is not None:
                    parent.add_widget(field, index=parent.children.index(self.bts_type))
        else:
            if field.parent is not None:
                field.parent.remove_widget(field)

    def update_bts_options(self, instance, vendor):
        if vendor == "HUAWEI":
            self.bts_type.values = (
                "Huawei BTS3900", "Huawei BTS3900A", "Huawei BTS3900L",
                "Huawei BTS5900", "Other"
            )
        elif vendor == "NOKIA":
            self.bts_type.values = (
                "Nokia Flexi EDGE BTS", "Nokia Flexi BTS", "Nokia UltraSite BTS", "Other"
            )
        else:
            self.bts_type.values = ("Other",)

        if self.bts_type.text not in self.bts_type.values:
            self.bts_type.text = self.bts_type.values[0]
        self._update_custom_fields()

    def on_bts_type_changed(self, instance, value):
        self._update_custom_fields()

    def _update_custom_fields(self):
        self._set_field_visible(self.custom_vendor_field, self.vendor.text == "Other")
        self._set_field_visible(self.custom_bts_type_field, self.bts_type.text == "Other")
        self._set_field_visible(self.custom_band_field, self.band.text == "Other")
        if self.vendor.text != "Other": self.custom_vendor.text = ""
        if self.bts_type.text != "Other": self.custom_bts_type.text = ""
        if self.band.text != "Other": self.custom_band.text = ""

    def load_2g(self):
        if CURRENT_2G_ID is None:
            self._update_custom_fields()
            return

        connection = get_connection()
        cursor = connection.cursor()
        cursor.execute("SELECT * FROM system_2g WHERE id=?", (CURRENT_2G_ID,))
        row = cursor.fetchone()
        connection.close()
        if not row:
            return

        self.vendor.text = row[2] or "Select"
        self.update_bts_options(self.vendor, self.vendor.text)
        self.bts_type.text = row[3] or (self.bts_type.values[0] if self.bts_type.values else "Select")
        self.band.text = row[4] or "Select"
        self.number_of_sectors.text = row[7] or ""
        self.sector_configuration.text = row[8] or ""
        self.notes.text = row[9] or ""

        # Extended columns are optional for backward compatibility.
        self.custom_vendor.text = row[11] if len(row) > 11 and row[11] else ""
        self.custom_bts_type.text = row[12] if len(row) > 12 and row[12] else ""
        self.custom_band.text = row[13] if len(row) > 13 and row[13] else ""
        self._update_custom_fields()

    def save_2g(self, instance):
        global CURRENT_2G_ID

        if self.vendor.text == "Select":
            show_message("Error", "Please select BTS Vendor.")
            return
        if self.vendor.text == "Other" and not self.custom_vendor.text.strip():
            show_message("Error", "Please enter the Vendor Name.")
            return
        if self.bts_type.text in ("Select", ""):
            show_message("Error", "Please select BTS Type.")
            return
        if self.bts_type.text == "Other" and not self.custom_bts_type.text.strip():
            show_message("Error", "Please enter the BTS Type.")
            return
        if self.band.text == "Select":
            show_message("Error", "Please select System Band.")
            return
        if self.band.text == "Other" and not self.custom_band.text.strip():
            show_message("Error", "Please enter the System Band.")
            return
        if not self.number_of_sectors.text.strip():
            show_message("Error", "Please enter Number of Sectors.")
            return

        # The main 2G form contains only system-level information.
        values = (
            CURRENT_SITE_ID, self.vendor.text, self.bts_type.text, self.band.text,
            "", "", self.number_of_sectors.text.strip(),
            self.sector_configuration.text.strip(), self.notes.text.strip(),
            "", self.custom_vendor.text.strip(), self.custom_bts_type.text.strip(),
            self.custom_band.text.strip()
        )

        connection = get_connection()
        cursor = connection.cursor()
        try:
            if CURRENT_2G_ID is None:
                cursor.execute("""
                    INSERT INTO system_2g (
                        site_id,bts_vendor,bts_type,system_band,esma,esea,
                        number_of_sectors,sector_configuration,notes,cards_information,
                        custom_vendor_name,custom_bts_type,custom_band
                    ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)
                """, values)
                CURRENT_2G_ID = cursor.lastrowid
            else:
                cursor.execute("""
                    UPDATE system_2g SET
                        bts_vendor=?, bts_type=?, system_band=?, esma=?, esea=?,
                        number_of_sectors=?, sector_configuration=?, notes=?,
                        cards_information=?, custom_vendor_name=?, custom_bts_type=?, custom_band=?
                    WHERE id=?
                """, values[1:] + (CURRENT_2G_ID,))
            connection.commit()
        except Exception as error:
            connection.rollback()
            connection.close()
            show_message("Database Error", str(error))
            return
        connection.close()
        show_message("Success", "2G System saved successfully.")

    def open_sectors(self, instance):

        if CURRENT_2G_ID is None:

            show_message(
                "Save First",
                "Please save the 2G System first."
            )

            return

        self.manager.current = "2g_sectors"

    def ask_delete(self, instance):

        if CURRENT_2G_ID is None:

            show_message(
                "Error",
                "No 2G System selected."
            )

            return

        confirm_delete(
            "Delete 2G System",
            "Delete this 2G System and all its sectors?",
            self.delete_2g
        )

    def delete_2g(self):

        global CURRENT_2G_ID

        connection = get_connection()
        cursor = connection.cursor()

        cursor.execute(
            "DELETE FROM system_2g_sectors WHERE system_2g_id=?",
            (CURRENT_2G_ID,)
        )

        cursor.execute("""
            DELETE FROM system_2g_sector_cards
            WHERE sector_id IN (
                SELECT id FROM system_2g_sectors WHERE system_2g_id=?
            )
        """, (CURRENT_2G_ID,))

        cursor.execute(
            "DELETE FROM system_2g WHERE id=?",
            (CURRENT_2G_ID,)
        )

        connection.commit()
        connection.close()

        CURRENT_2G_ID = None

        self.manager.current = "2g_list"


# ============================================================
# 2G SECTORS LIST
# ============================================================

class TwoGSectorsScreen(Screen):

    def on_pre_enter(self):
        self.build_screen()

    def build_screen(self):

        self.clear_widgets()

        main = BoxLayout(
            orientation="vertical",
            padding=10,
            spacing=10
        )

        title = Label(
            text="2G SECTORS",
            font_size=24,
            size_hint_y=None,
            height=65
        )

        add = Button(
            text="ADD NEW SECTOR",
            size_hint_y=None,
            height=60,
            font_size=18
        )

        add.bind(
            on_release=self.add_sector
        )

        scroll = ScrollView()

        self.sectors_layout = BoxLayout(
            orientation="vertical",
            spacing=15,
            padding=[5, 5],
            size_hint_y=None
        )

        self.sectors_layout.bind(
            minimum_height=
            self.sectors_layout.setter("height")
        )

        scroll.add_widget(self.sectors_layout)

        back = Button(
            text="BACK TO 2G SYSTEM",
            size_hint_y=None,
            height=60
        )

        back.bind(
            on_release=lambda x:
            setattr(
                self.manager,
                "current",
                "2g_form"
            )
        )

        main.add_widget(title)
        main.add_widget(add)
        main.add_widget(scroll)
        main.add_widget(back)

        self.add_widget(main)

        self.display_sectors()

    def display_sectors(self):

        self.sectors_layout.clear_widgets()

        connection = get_connection()
        cursor = connection.cursor()

        cursor.execute("""
        SELECT
            s.id, s.sector_number, s.sector_name, s.band,
            s.antenna_type, s.azimuth, s.custom_band,
            (SELECT GROUP_CONCAT(c.card_type, ', ')
             FROM system_2g_sector_cards c
             WHERE c.sector_id=s.id) AS cards
        FROM system_2g_sectors s
        WHERE s.system_2g_id=?
        ORDER BY s.sector_number
        """, (CURRENT_2G_ID,))

        rows = cursor.fetchall()

        connection.close()

        if not rows:

            self.sectors_layout.add_widget(
                Label(
                    text="No Sectors Added",
                    size_hint_y=None,
                    height=70,
                    font_size=18
                )
            )

            return

        for row in rows:

            sector_id = row[0]

            card = BoxLayout(
                orientation="vertical",
                size_hint_y=None,
                height=225,
                padding=12,
                spacing=7
            )

            band_display = row[6] if row[3] == "Other" and row[6] else (row[3] or "-")
            info = Label(
                text=(
                    f"[b]SECTOR {row[1]} - {row[2] or ''}[/b]\n"
                    f"Band: {band_display}\n"
                    f"Antenna: {row[4] or '-'}\n"
                    f"Azimuth: {row[5] or '-'}\n"
                    f"Cards: {row[7] or 'No cards assigned'}"
                ),
                markup=True,
                halign="left",
                valign="middle",
                font_size=16
            )

            info.bind(
                size=lambda instance, value:
                setattr(instance, "text_size", value)
            )

            button = Button(
                text="OPEN / EDIT SECTOR",
                size_hint_y=None,
                height=50
            )

            button.bind(
                on_release=lambda x,
                value=sector_id:
                self.edit_sector(value)
            )

            card.add_widget(info)
            card.add_widget(button)

            self.sectors_layout.add_widget(card)

    def add_sector(self, instance):

        global CURRENT_2G_SECTOR_ID

        CURRENT_2G_SECTOR_ID = None

        self.manager.current = "2g_sector_form"

    def edit_sector(self, sector_id):

        global CURRENT_2G_SECTOR_ID

        CURRENT_2G_SECTOR_ID = sector_id

        self.manager.current = "2g_sector_form"


# ============================================================
# 2G SECTOR FORM
# ============================================================

class TwoGSectorFormScreen(Screen):

    def on_pre_enter(self):
        self.build_form()

    def get_2g_vendor_type(self):
        if CURRENT_2G_ID is None:
            return "", ""
        connection = get_connection()
        cursor = connection.cursor()
        cursor.execute("SELECT bts_vendor, bts_type FROM system_2g WHERE id=?", (CURRENT_2G_ID,))
        row = cursor.fetchone()
        connection.close()
        return ((row[0] or "") if row else "", (row[1] or "") if row else "")

    def card_catalogue(self, vendor, bts_type):
        vendor_u = (vendor or "").upper()
        bts_u = (bts_type or "").upper()
        if "HUAWEI" in vendor_u:
            return ["GTMU", "GTMUe", "UBFA", "UPEU", "UPEA", "UBRI", "UTRP", "Other"]
        if "NOKIA" in vendor_u:
            if "FLEXI EDGE" in bts_u:
                return ["ESMA", "ESEA", "EXGA", "ERGA", "EXDA", "ERDA", "Other"]
            if "FLEXI" in bts_u:
                return ["ESMB", "ESMA", "ESEA", "EXGA", "ERGA", "EXDA", "ERDA", "Other"]
            return ["ESMA", "ESEA", "ESMB", "EXGA", "ERGA", "EXDA", "ERDA", "Other"]
        return ["Other"]

    def section_title(self, text):
        lbl = Label(text=text, font_size=18, bold=True, size_hint_y=None, height=42,
                    halign="left", valign="middle")
        lbl.bind(size=lambda i,v: setattr(i, "text_size", v))
        return lbl

    def build_form(self):
        # Rebuilt 2G Sector form: custom fields are dynamically added/removed.
        # They are NEVER kept as zero-height children, which avoids Kivy drawing
        # them over the widgets below.
        self.clear_widgets()
        vendor, bts_type = self.get_2g_vendor_type()

        scroll = ScrollView(do_scroll_x=False)
        form = BoxLayout(orientation="vertical", padding=[16,14], spacing=10, size_hint_y=None)
        form.bind(minimum_height=form.setter("height"))

        title = Label(text=f"2G SECTOR\n{CURRENT_SITE_NAME}", font_size=23, bold=True,
                      size_hint_y=None, height=78, halign="center", valign="middle")
        title.bind(size=lambda i,v: setattr(i, "text_size", v))
        form.add_widget(title)

        summary = Label(text=f"BTS VENDOR: {vendor or '-'}    |    BTS TYPE: {bts_type or '-'}",
                        font_size=15, size_hint_y=None, height=42, halign="left", valign="middle")
        summary.bind(size=lambda i,v: setattr(i, "text_size", v))
        form.add_widget(summary)
        form.add_widget(self.section_title("SECTOR INFORMATION"))

        self.sector_number = create_input()
        self.sector_name = create_input()
        self.band = create_spinner(["900 MHz", "1800 MHz", "900 + 1800 MHz", "Other"])
        self.band.bind(text=self.on_band_changed)
        self.other_band = create_input()
        self.other_band_container = self._custom_field_block("Custom Sector Band", self.other_band)
        self.other_band_container_added = False

        self.antenna_type = create_input()
        self.azimuth = create_input()
        self.notes = create_input(multiline=True, height=100)

        add_field(form, "Sector Number", self.sector_number)
        add_field(form, "Sector Name", self.sector_name)

        # Sector Band is a fixed block. The custom block is inserted only when needed.
        self.band_block = BoxLayout(orientation="vertical", size_hint_y=None, height=92, spacing=6)
        self.band_block.add_widget(self._field_label("Sector Band"))
        self.band_block.add_widget(self.band)
        form.add_widget(self.band_block)

        add_field(form, "Antenna Type", self.antenna_type)
        add_field(form, "Azimuth", self.azimuth)
        add_field(form, "Notes", self.notes)

        form.add_widget(self.section_title("CARDS INSTALLED ON THIS SECTOR"))
        hint = Label(text="Choose the hardware cards installed on this sector.\nCards are saved separately for each sector.",
                     font_size=14, size_hint_y=None, height=48, halign="left", valign="middle")
        hint.bind(size=lambda i,v: setattr(i, "text_size", v))
        form.add_widget(hint)

        self.card_section = BoxLayout(orientation="vertical", size_hint_y=None, height=60, spacing=6)
        toolbar = BoxLayout(size_hint_y=None, height=54, spacing=8)
        self.card_selector = Spinner(text="Select Card", values=self.card_catalogue(vendor, bts_type),
                                     size_hint_x=0.70, size_hint_y=None, height=54, font_size=16)
        add_card = Button(text="+ ADD CARD", size_hint_x=0.30, font_size=15)
        add_card.bind(on_release=self.add_sector_card)
        toolbar.add_widget(self.card_selector)
        toolbar.add_widget(add_card)
        self.card_section.add_widget(toolbar)

        self.other_card = create_input()
        self.other_card_container = self._custom_field_block("Custom Card Name", self.other_card)
        self.other_card_container_added = False
        form.add_widget(self.card_section)
        self.card_selector.bind(text=self.on_card_changed)

        self.cards_layout = BoxLayout(orientation="vertical", spacing=6, size_hint_y=None)
        self.cards_layout.bind(minimum_height=self.cards_layout.setter("height"))
        form.add_widget(self.cards_layout)

        form.add_widget(self.section_title("ACTIONS"))
        save = Button(text="SAVE / UPDATE SECTOR", size_hint_y=None, height=58, font_size=17)
        delete = Button(text="DELETE SECTOR", size_hint_y=None, height=52, font_size=16)
        back = Button(text="BACK TO SECTORS", size_hint_y=None, height=52, font_size=16)
        save.bind(on_release=self.save_sector)
        delete.bind(on_release=self.ask_delete)
        back.bind(on_release=lambda x: setattr(self.manager, "current", "2g_sectors"))
        form.add_widget(save); form.add_widget(delete); form.add_widget(back)

        scroll.add_widget(form)
        self.add_widget(scroll)
        self.load_sector()

    def _field_label(self, text):
        label = Label(text=text, size_hint_y=None, height=30, font_size=17,
                      halign="left", valign="middle")
        label.bind(size=lambda i,v: setattr(i, "text_size", v))
        return label

    def _custom_field_block(self, title, widget):
        block = BoxLayout(orientation="vertical", size_hint_y=None, height=86, spacing=6)
        block.add_widget(self._field_label(title))
        block.add_widget(widget)
        return block

    def on_band_changed(self, spinner, value):
        if not hasattr(self, "band_block"):
            return
        if value == "Other":
            if not self.other_band_container_added:
                self.band_block.add_widget(self.other_band_container)
                self.other_band_container_added = True
            self.band_block.height = 92 + 86 + 6
        else:
            if self.other_band_container_added:
                self.band_block.remove_widget(self.other_band_container)
                self.other_band_container_added = False
            self.band_block.height = 92
            self.other_band.text = ""

    def on_card_changed(self, spinner, value):
        if not hasattr(self, "card_section"):
            return
        if value == "Other":
            if not self.other_card_container_added:
                self.card_section.add_widget(self.other_card_container)
                self.other_card_container_added = True
            self.card_section.height = 54 + 86 + 6
        else:
            if self.other_card_container_added:
                self.card_section.remove_widget(self.other_card_container)
                self.other_card_container_added = False
            self.card_section.height = 60
            self.other_card.text = ""

    def add_sector_card(self, instance):
        card = self.other_card.text.strip() if self.card_selector.text == "Other" else self.card_selector.text
        if not card or card == "Select Card":
            show_message("Card Selection", "Please select a card."); return
        for child in self.cards_layout.children:
            if getattr(child, "card_name", "") == card:
                show_message("Card Selection", "This card is already added to this sector."); return
        row = BoxLayout(size_hint_y=None, height=48, spacing=7, padding=[5,3])
        row.card_name = card
        label = Label(text=card, font_size=16, halign="left", valign="middle")
        label.bind(size=lambda i,v: setattr(i,"text_size",v))
        remove = Button(text="REMOVE", size_hint_x=None, width=95, font_size=13)
        remove.bind(on_release=lambda x, r=row: self.remove_sector_card(r))
        row.add_widget(label); row.add_widget(remove)
        self.cards_layout.add_widget(row)
        self.card_selector.text = "Select Card"
        self.other_card.text = ""
        self.other_card_container.height = 0

    def remove_sector_card(self, row):
        self.cards_layout.remove_widget(row)

    def load_sector(self):
        if CURRENT_2G_SECTOR_ID is None: return
        connection = get_connection(); cursor = connection.cursor()
        cursor.execute("SELECT * FROM system_2g_sectors WHERE id=?", (CURRENT_2G_SECTOR_ID,))
        row = cursor.fetchone()
        cursor.execute("SELECT card_type FROM system_2g_sector_cards WHERE sector_id=? ORDER BY id", (CURRENT_2G_SECTOR_ID,))
        cards = cursor.fetchall(); connection.close()
        if not row: return
        self.sector_number.text = str(row[2] or "")
        self.sector_name.text = row[3] or ""
        band = row[4] or "Select"
        if band in ["900 MHz", "1800 MHz", "900 + 1800 MHz"]: self.band.text = band
        else:
            self.band.text = "Other"; self.other_band.text = band
        self.on_band_changed(self.band, self.band.text)
        self.antenna_type.text = row[9] or ""; self.azimuth.text = row[10] or ""; self.notes.text = row[11] or ""
        for c in cards:
            row_widget = BoxLayout(size_hint_y=None, height=48, spacing=7, padding=[5,3])
            row_widget.card_name = c[0]
            label = Label(text=c[0], font_size=16, halign="left", valign="middle")
            label.bind(size=lambda i,v: setattr(i,"text_size",v))
            remove = Button(text="REMOVE", size_hint_x=None, width=95, font_size=13)
            remove.bind(on_release=lambda x, r=row_widget: self.remove_sector_card(r))
            row_widget.add_widget(label); row_widget.add_widget(remove); self.cards_layout.add_widget(row_widget)

    def get_selected_cards(self):
        return [getattr(child, "card_name") for child in reversed(self.cards_layout.children) if getattr(child, "card_name", None)]

    def save_sector(self, instance):
        global CURRENT_2G_SECTOR_ID
        if not self.sector_number.text.strip(): show_message("Error", "Sector Number is required."); return
        if self.band.text == "Select": show_message("Error", "Please select the Sector Band."); return
        band = self.other_band.text.strip() if self.band.text == "Other" else self.band.text
        if self.band.text == "Other" and not band: show_message("Error", "Please enter the custom Sector Band."); return
        values = (CURRENT_2G_ID, self.sector_number.text.strip(), self.sector_name.text.strip(), band, "", "", "", "", self.antenna_type.text.strip(), self.azimuth.text.strip(), self.notes.text.strip())
        connection = get_connection(); cursor = connection.cursor()
        try:
            if CURRENT_2G_SECTOR_ID is None:
                cursor.execute("""INSERT INTO system_2g_sectors (system_2g_id,sector_number,sector_name,band,exga_number,erga_number,exda_number,erda_number,antenna_type,azimuth,notes) VALUES (?,?,?,?,?,?,?,?,?,?,?)""", values)
                CURRENT_2G_SECTOR_ID = cursor.lastrowid
            else:
                cursor.execute("""UPDATE system_2g_sectors SET sector_number=?,sector_name=?,band=?,exga_number=?,erga_number=?,exda_number=?,erda_number=?,antenna_type=?,azimuth=?,notes=? WHERE id=?""", values[1:] + (CURRENT_2G_SECTOR_ID,))
            cursor.execute("DELETE FROM system_2g_sector_cards WHERE sector_id=?", (CURRENT_2G_SECTOR_ID,))
            for card in self.get_selected_cards(): cursor.execute("INSERT INTO system_2g_sector_cards (sector_id, card_type) VALUES (?,?)", (CURRENT_2G_SECTOR_ID, card))
            connection.commit()
        except Exception as error:
            connection.rollback(); connection.close(); show_message("Database Error", str(error)); return
        connection.close(); show_message("Success", "Sector and sector cards saved successfully.")

    def ask_delete(self, instance):
        if CURRENT_2G_SECTOR_ID is None: return
        confirm_delete("Delete Sector", "Are you sure you want to delete this sector and its cards?", self.delete_sector)

    def delete_sector(self):
        global CURRENT_2G_SECTOR_ID
        connection = get_connection(); cursor = connection.cursor()
        cursor.execute("DELETE FROM system_2g_sector_cards WHERE sector_id=?", (CURRENT_2G_SECTOR_ID,))
        cursor.execute("DELETE FROM system_2g_sectors WHERE id=?", (CURRENT_2G_SECTOR_ID,)); connection.commit(); connection.close()
        CURRENT_2G_SECTOR_ID = None; self.manager.current = "2g_sectors"


# ============================================================
# LTE SYSTEM LIST
# ============================================================

class LTEListScreen(Screen):

    def on_pre_enter(self):
        self.build_screen()

    def build_screen(self):

        self.clear_widgets()

        main = BoxLayout(
            orientation="vertical",
            padding=10,
            spacing=10
        )

        title = Label(
            text=f"LTE SYSTEMS\n{CURRENT_SITE_NAME}",
            font_size=23,
            size_hint_y=None,
            height=80
        )

        add = Button(
            text="ADD NEW LTE SYSTEM",
            size_hint_y=None,
            height=60,
            font_size=18
        )

        add.bind(
            on_release=self.add_lte
        )

        scroll = ScrollView()

        self.lte_layout = BoxLayout(
            orientation="vertical",
            spacing=15,
            padding=[5, 5],
            size_hint_y=None
        )

        self.lte_layout.bind(
            minimum_height=
            self.lte_layout.setter("height")
        )

        scroll.add_widget(self.lte_layout)

        back = Button(
            text="BACK TO SITE SYSTEMS DASHBOARD",
            size_hint_y=None,
            height=60
        )

        back.bind(
            on_release=lambda x:
            setattr(
                self.manager,
                "current",
                "systems"
            )
        )

        main.add_widget(title)
        main.add_widget(add)
        main.add_widget(scroll)
        main.add_widget(back)

        self.add_widget(main)

        self.display_lte()

    def display_lte(self):

        self.lte_layout.clear_widgets()

        connection = get_connection()
        cursor = connection.cursor()

        cursor.execute("""
        SELECT
            id,
            bbu_type,
            number_of_rru,
            rru_information,
            frequency_band
        FROM lte_system
        WHERE site_id=?
        ORDER BY id DESC
        """, (CURRENT_SITE_ID,))

        rows = cursor.fetchall()

        connection.close()

        if not rows:

            self.lte_layout.add_widget(
                Label(
                    text="No LTE Systems Added",
                    size_hint_y=None,
                    height=70,
                    font_size=18
                )
            )

            return

        for row in rows:

            lte_id = row[0]

            connection = get_connection()
            cursor = connection.cursor()

            cursor.execute(
                "SELECT COUNT(*) FROM lte_rru WHERE lte_system_id=?",
                (lte_id,)
            )

            rru = cursor.fetchone()[0]

            cursor.execute(
                "SELECT COUNT(*) FROM lte_cells WHERE lte_system_id=?",
                (lte_id,)
            )

            cells = cursor.fetchone()[0]

            connection.close()

            card = BoxLayout(
                orientation="vertical",
                size_hint_y=None,
                height=220,
                padding=12,
                spacing=8
            )

            info = Label(
                text=(
                    f"[b]LTE SYSTEM[/b]\n"
                    f"BBU: {row[1] or '-'}\n"
                    f"RRU: {rru}\n"
                    f"LTE Cells: {cells}\n"
                    f"Configured RRU: {row[2] or '-'}\n"
                    f"Band: {row[4] or '-'}"
                ),
                markup=True,
                halign="left",
                valign="middle",
                font_size=17
            )

            info.bind(
                size=lambda instance, value:
                setattr(instance, "text_size", value)
            )

            button = Button(
                text="OPEN / EDIT LTE SYSTEM",
                size_hint_y=None,
                height=50
            )

            button.bind(
                on_release=lambda x,
                value=lte_id:
                self.open_lte(value)
            )

            card.add_widget(info)
            card.add_widget(button)

            self.lte_layout.add_widget(card)

    def add_lte(self, instance):

        global CURRENT_LTE_ID

        CURRENT_LTE_ID = None

        self.manager.current = "lte_form"

    def open_lte(self, lte_id):

        global CURRENT_LTE_ID

        CURRENT_LTE_ID = lte_id

        self.manager.current = "lte_form"


# ============================================================
# ============================================================
# LTE SYSTEM FORM
# ============================================================

LTE_BBU_CARD_CATALOGUE = {
    "Huawei BBU3900": ["UMPT", "UBBP", "UPEU", "UELP", "UBRI", "Other"],
    "Huawei BBU3910": ["UMPT", "UBBP", "UPEU", "UPEA", "UTRP", "UBRI", "Other"],
    "Huawei BBU5900": ["UMPT", "UBBP", "UPEU", "UPEA", "UBRI", "Other"],
    "Nokia AirScale": ["ABIA", "ASIA", "ABIP", "Other"],
    "Nokia Flexi": ["FSMF", "FBBA", "FIxx", "Other"],
    "Other": ["Other"],
}


class LTEFormScreen(Screen):

    def on_pre_enter(self):
        self.build_form()

    def _field_label(self, text):
        label = Label(
            text=text,
            size_hint_y=None,
            height=32,
            font_size=17,
            halign="left",
            valign="middle"
        )
        label.bind(size=lambda i, v: setattr(i, "text_size", v))
        return label

    def _make_dynamic_block(self):
        return BoxLayout(
            orientation="vertical",
            size_hint_y=None,
            height=0,
            spacing=6
        )

    def build_form(self):
        self.clear_widgets()

        scroll = ScrollView(do_scroll_x=False)
        form = BoxLayout(
            orientation="vertical",
            padding=[20, 16],
            spacing=10,
            size_hint_y=None
        )
        form.bind(minimum_height=form.setter("height"))

        title = Label(
            text="LTE SYSTEM INFORMATION",
            font_size=24,
            bold=True,
            size_hint_y=None,
            height=65,
            halign="center",
            valign="middle"
        )
        title.bind(size=lambda i, v: setattr(i, "text_size", v))
        form.add_widget(title)

        # --------------------------------------------------------
        # BBU
        # --------------------------------------------------------
        self.bbu = create_spinner([
            "Huawei BBU3900",
            "Huawei BBU3910",
            "Huawei BBU5900",
            "Nokia AirScale",
            "Nokia Flexi",
            "Other"
        ])
        self.bbu.bind(text=self.on_bbu_changed)

        form.add_widget(self._field_label("BBU Type"))
        form.add_widget(self.bbu)

        # Custom BBU is dynamically inserted/removed.
        self.custom_bbu_block = self._make_dynamic_block()
        self.custom_bbu_label = self._field_label("Custom BBU Name")
        self.custom_bbu = create_input()
        self.custom_bbu_block.add_widget(self.custom_bbu_label)
        self.custom_bbu_block.add_widget(self.custom_bbu)

        # --------------------------------------------------------
        # BBU CARDS
        # --------------------------------------------------------
        form.add_widget(self._field_label("BBU CARDS INSTALLED"))

        hint = Label(
            text="Select the cards installed inside this BBU and add them one by one.",
            font_size=14,
            size_hint_y=None,
            height=42,
            halign="left",
            valign="middle"
        )
        hint.bind(size=lambda i, v: setattr(i, "text_size", v))
        form.add_widget(hint)

        toolbar = BoxLayout(
            orientation="horizontal",
            size_hint_y=None,
            height=54,
            spacing=8
        )
        self.card_selector = create_spinner(["Other"])
        self.card_selector.size_hint_x = 0.70
        self.card_selector.bind(text=self.on_bbu_card_changed)

        add_card = Button(
            text="+ ADD CARD",
            size_hint_x=0.30,
            font_size=15
        )
        add_card.bind(on_release=self.add_bbu_card)
        toolbar.add_widget(self.card_selector)
        toolbar.add_widget(add_card)
        form.add_widget(toolbar)

        self.custom_card_block = self._make_dynamic_block()
        self.custom_card_label = self._field_label("Custom Card Name")
        self.custom_card = create_input()
        self.custom_card_block.add_widget(self.custom_card_label)
        self.custom_card_block.add_widget(self.custom_card)

        self.bbu_cards_layout = BoxLayout(
            orientation="vertical",
            spacing=6,
            size_hint_y=None
        )
        self.bbu_cards_layout.bind(
            minimum_height=self.bbu_cards_layout.setter("height")
        )
        form.add_widget(self.bbu_cards_layout)

        # --------------------------------------------------------
        # RRU / FREQUENCY
        # --------------------------------------------------------
        self.number_of_rru = create_input()
        add_field(form, "Number of RRU", self.number_of_rru)

        self.frequency_band = create_spinner([
            "LTE 700 MHz",
            "LTE 800 MHz",
            "LTE 900 MHz",
            "LTE 1800 MHz",
            "LTE 2100 MHz",
            "LTE 2600 MHz",
            "Other"
        ])
        self.frequency_band.bind(text=self.on_frequency_changed)
        add_field(form, "Frequency Band", self.frequency_band)

        # Custom frequency is dynamically inserted/removed.
        self.custom_frequency_block = self._make_dynamic_block()
        self.custom_frequency_label = self._field_label("Custom Frequency Band")
        self.custom_frequency = create_input()
        self.custom_frequency_block.add_widget(self.custom_frequency_label)
        self.custom_frequency_block.add_widget(self.custom_frequency)

        self.notes = create_input(multiline=True, height=110)
        add_field(form, "Notes", self.notes)

        # --------------------------------------------------------
        # ACTIONS
        # --------------------------------------------------------
        save = Button(
            text="SAVE / UPDATE LTE SYSTEM",
            size_hint_y=None,
            height=60,
            font_size=18
        )
        rru_button = Button(
            text="MANAGE LTE RRU",
            size_hint_y=None,
            height=60,
            font_size=18
        )
        cells_button = Button(
            text="MANAGE LTE CELLS / SECTORS",
            size_hint_y=None,
            height=60,
            font_size=18
        )
        delete = Button(
            text="DELETE LTE SYSTEM",
            size_hint_y=None,
            height=60
        )
        back = Button(
            text="BACK TO LTE SYSTEMS",
            size_hint_y=None,
            height=60
        )

        save.bind(on_release=self.save_lte)
        rru_button.bind(on_release=self.open_rru)
        cells_button.bind(on_release=self.open_cells)
        delete.bind(on_release=self.ask_delete)
        back.bind(on_release=lambda x: setattr(self.manager, "current", "lte_list"))

        form.add_widget(save)
        form.add_widget(rru_button)
        form.add_widget(cells_button)
        form.add_widget(delete)
        form.add_widget(back)

        scroll.add_widget(form)
        self.add_widget(scroll)

        # Set initial card list without firing layout changes prematurely.
        self.update_bbu_card_options(self.bbu.text)
        self.load_lte()
        self.refresh_dynamic_fields()

    def _insert_dynamic_block(self, block, after_widget):
        if block.parent is not None:
            return
        parent = after_widget.parent
        if parent is None:
            return
        index = parent.children.index(after_widget)
        parent.add_widget(block, index=index)

    def _remove_dynamic_block(self, block):
        if block.parent is not None:
            block.parent.remove_widget(block)

    def refresh_dynamic_fields(self):
        if self.bbu.text == "Other":
            self._insert_dynamic_block(self.custom_bbu_block, self.bbu)
            self.custom_bbu_block.height = 88
        else:
            self._remove_dynamic_block(self.custom_bbu_block)
            self.custom_bbu.text = ""

        if self.frequency_band.text == "Other":
            self._insert_dynamic_block(self.custom_frequency_block, self.frequency_band)
            self.custom_frequency_block.height = 88
        else:
            self._remove_dynamic_block(self.custom_frequency_block)
            self.custom_frequency.text = ""

        if self.card_selector.text == "Other":
            self._insert_dynamic_block(self.custom_card_block, self.card_selector)
            self.custom_card_block.height = 88
        else:
            self._remove_dynamic_block(self.custom_card_block)
            self.custom_card.text = ""

    def on_bbu_changed(self, spinner, value):
        self.update_bbu_card_options(value)
        self.refresh_dynamic_fields()

    def update_bbu_card_options(self, bbu_type):
        values = LTE_BBU_CARD_CATALOGUE.get(bbu_type, ["Other"])
        self.card_selector.values = values
        if self.card_selector.text not in values:
            self.card_selector.text = "Select Card" if "Select Card" in values else values[0]

    def on_bbu_card_changed(self, spinner, value):
        self.refresh_dynamic_fields()

    def on_frequency_changed(self, spinner, value):
        self.refresh_dynamic_fields()

    def add_bbu_card(self, instance):
        if self.card_selector.text in ("Select Card", ""):
            show_message("Card Selection", "Please select a BBU card.")
            return

        card = self.custom_card.text.strip() if self.card_selector.text == "Other" else self.card_selector.text
        if not card:
            show_message("Card Selection", "Please enter the custom card name.")
            return

        # A physical BBU may contain more than one card of the same type,
        # so duplicate card types are allowed.
        row = BoxLayout(
            orientation="horizontal",
            size_hint_y=None,
            height=48,
            spacing=7,
            padding=[5, 3]
        )
        row.card_name = card

        label = Label(
            text=card,
            font_size=16,
            halign="left",
            valign="middle"
        )
        label.bind(size=lambda i, v: setattr(i, "text_size", v))

        remove = Button(
            text="REMOVE",
            size_hint_x=None,
            width=95,
            font_size=13
        )
        remove.bind(on_release=lambda x, r=row: self.remove_bbu_card(r))

        row.add_widget(label)
        row.add_widget(remove)
        self.bbu_cards_layout.add_widget(row)

        self.card_selector.text = "Select Card" if "Select Card" in self.card_selector.values else self.card_selector.values[0]
        self.custom_card.text = ""
        self.refresh_dynamic_fields()

    def remove_bbu_card(self, row):
        self.bbu_cards_layout.remove_widget(row)

    def get_selected_bbu_cards(self):
        return [
            getattr(child, "card_name")
            for child in reversed(self.bbu_cards_layout.children)
            if getattr(child, "card_name", None)
        ]

    def load_lte(self):
        if CURRENT_LTE_ID is None:
            return

        connection = get_connection()
        cursor = connection.cursor()
        cursor.execute(
            "SELECT * FROM lte_system WHERE id=?",
            (CURRENT_LTE_ID,)
        )
        row = cursor.fetchone()

        cursor.execute(
            "SELECT card_type FROM lte_bbu_cards WHERE lte_system_id=? ORDER BY id",
            (CURRENT_LTE_ID,)
        )
        cards = cursor.fetchall()
        connection.close()

        if not row:
            return

        self.bbu.text = row[2] or "Select"
        self.number_of_rru.text = row[3] or ""

        # Legacy rru_information remains supported as a fallback.
        legacy_card_text = row[4] or ""

        self.frequency_band.text = row[5] or "Select"
        self.notes.text = row[6] or ""

        # Custom fields are stored in the new columns when available.
        custom_bbu = row[7] if len(row) > 7 else ""
        custom_frequency = row[8] if len(row) > 8 else ""

        if self.bbu.text == "Other":
            self.custom_bbu.text = custom_bbu or ""
        if self.frequency_band.text == "Other":
            self.custom_frequency.text = custom_frequency or ""

        self.update_bbu_card_options(self.bbu.text)

        loaded_cards = [c[0] for c in cards]
        if not loaded_cards and legacy_card_text:
            loaded_cards = [x.strip() for x in legacy_card_text.replace(";", ",").split(",") if x.strip()]

        for card_name in loaded_cards:
            self._add_loaded_card(card_name)

    def _add_loaded_card(self, card_name):
        row = BoxLayout(
            orientation="horizontal",
            size_hint_y=None,
            height=48,
            spacing=7,
            padding=[5, 3]
        )
        row.card_name = card_name

        label = Label(
            text=card_name,
            font_size=16,
            halign="left",
            valign="middle"
        )
        label.bind(size=lambda i, v: setattr(i, "text_size", v))

        remove = Button(
            text="REMOVE",
            size_hint_x=None,
            width=95,
            font_size=13
        )
        remove.bind(on_release=lambda x, r=row: self.remove_bbu_card(r))

        row.add_widget(label)
        row.add_widget(remove)
        self.bbu_cards_layout.add_widget(row)

    def save_lte(self, instance):
        global CURRENT_LTE_ID

        if self.bbu.text == "Select":
            show_message("Error", "Please select BBU.")
            return

        if self.bbu.text == "Other" and not self.custom_bbu.text.strip():
            show_message("Error", "Please enter the custom BBU name.")
            return

        if self.frequency_band.text == "Select":
            show_message("Error", "Please select Frequency Band.")
            return

        if self.frequency_band.text == "Other" and not self.custom_frequency.text.strip():
            show_message("Error", "Please enter the custom Frequency Band.")
            return

        selected_cards = self.get_selected_bbu_cards()
        cards_text = ", ".join(selected_cards)

        # Keep the legacy rru_information column synchronized for compatibility.
        values = (
            CURRENT_SITE_ID,
            self.bbu.text,
            self.number_of_rru.text.strip(),
            cards_text,
            self.frequency_band.text,
            self.notes.text.strip(),
            self.custom_bbu.text.strip() if self.bbu.text == "Other" else "",
            self.custom_frequency.text.strip() if self.frequency_band.text == "Other" else ""
        )

        connection = get_connection()
        cursor = connection.cursor()

        try:
            if CURRENT_LTE_ID is None:
                cursor.execute("""
                INSERT INTO lte_system (
                    site_id, bbu_type, number_of_rru, rru_information,
                    frequency_band, notes, custom_bbu_name, custom_frequency_band
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """, values)
                CURRENT_LTE_ID = cursor.lastrowid
            else:
                cursor.execute("""
                UPDATE lte_system SET
                    bbu_type=?, number_of_rru=?, rru_information=?,
                    frequency_band=?, notes=?, custom_bbu_name=?,
                    custom_frequency_band=?
                WHERE id=?
                """, values[1:] + (CURRENT_LTE_ID,))

            cursor.execute(
                "DELETE FROM lte_bbu_cards WHERE lte_system_id=?",
                (CURRENT_LTE_ID,)
            )

            for card_name in selected_cards:
                cursor.execute("""
                INSERT INTO lte_bbu_cards (lte_system_id, card_type, notes)
                VALUES (?, ?, '')
                """, (CURRENT_LTE_ID, card_name))

            connection.commit()

        except Exception as error:
            connection.rollback()
            connection.close()
            show_message("Database Error", str(error))
            return

        connection.close()
        show_message("Success", "LTE System saved successfully.")

    def open_rru(self, instance):
        if CURRENT_LTE_ID is None:
            show_message("Save First", "Please save the LTE System first.")
            return
        self.manager.current = "lte_rru_list"

    def open_cells(self, instance):
        if CURRENT_LTE_ID is None:
            show_message("Save First", "Please save the LTE System first.")
            return
        self.manager.current = "lte_cell_list"

    def ask_delete(self, instance):
        if CURRENT_LTE_ID is None:
            self.manager.current = "lte_list"
            return

        confirm_delete(
            "Delete LTE System",
            "Delete this LTE System and its RRU/Cell records?",
            self.delete_lte
        )

    def delete_lte(self):
        global CURRENT_LTE_ID

        if CURRENT_LTE_ID is None:
            return

        connection = get_connection()
        cursor = connection.cursor()
        try:
            cursor.execute("DELETE FROM lte_bbu_cards WHERE lte_system_id=?", (CURRENT_LTE_ID,))
            cursor.execute("DELETE FROM lte_rru WHERE lte_system_id=?", (CURRENT_LTE_ID,))
            cursor.execute("DELETE FROM lte_cells WHERE lte_system_id=?", (CURRENT_LTE_ID,))
            cursor.execute("DELETE FROM lte_system WHERE id=?", (CURRENT_LTE_ID,))
            connection.commit()
        except Exception as error:
            connection.rollback()
            connection.close()
            show_message("Delete Error", str(error))
            return

        connection.close()
        CURRENT_LTE_ID = None
        self.manager.current = "lte_list"


# LTE RRU LIST
# ============================================================

class LTERRUListScreen(Screen):

    def on_pre_enter(self):
        self.build_screen()

    def build_screen(self):

        self.clear_widgets()

        main = BoxLayout(
            orientation="vertical",
            padding=10,
            spacing=10
        )

        title = Label(
            text="LTE RRU",
            font_size=24,
            size_hint_y=None,
            height=65
        )

        add = Button(
            text="ADD NEW RRU",
            size_hint_y=None,
            height=60,
            font_size=18
        )

        add.bind(
            on_release=self.add_rru
        )

        scroll = ScrollView()

        self.layout = BoxLayout(
            orientation="vertical",
            spacing=15,
            padding=5,
            size_hint_y=None
        )

        self.layout.bind(
            minimum_height=
            self.layout.setter("height")
        )

        scroll.add_widget(self.layout)

        back = Button(
            text="BACK TO LTE SYSTEM",
            size_hint_y=None,
            height=60
        )

        back.bind(
            on_release=lambda x:
            setattr(
                self.manager,
                "current",
                "lte_form"
            )
        )

        main.add_widget(title)
        main.add_widget(add)
        main.add_widget(scroll)
        main.add_widget(back)

        self.add_widget(main)

        self.display_rru()

    def display_rru(self):

        self.layout.clear_widgets()

        connection = get_connection()
        cursor = connection.cursor()

        cursor.execute("""
        SELECT
            id,
            rru_name,
            rru_model,
            vendor,
            frequency_band,
            technology,
            tx_rx,
            sector,
            cpri_port,
            status
        FROM lte_rru
        WHERE lte_system_id=?
        ORDER BY id
        """, (CURRENT_LTE_ID,))

        rows = cursor.fetchall()

        connection.close()

        if not rows:

            self.layout.add_widget(
                Label(
                    text="No LTE RRU Added",
                    size_hint_y=None,
                    height=70
                )
            )

            return

        for row in rows:

            rru_id = row[0]

            card = BoxLayout(
                orientation="vertical",
                size_hint_y=None,
                height=250,
                padding=12,
                spacing=8
            )

            info = Label(
                text=(
                    f"[b]{row[1] or 'RRU'}[/b]\n"
                    f"Model: {row[2] or '-'}\n"
                    f"Vendor: {row[3] or '-'}\n"
                    f"Band: {row[4] or '-'}\n"
                    f"Technology: {row[5] or '-'}\n"
                    f"TX/RX: {row[6] or '-'}\n"
                    f"Sector: {row[7] or '-'}\n"
                    f"CPRI: {row[8] or '-'}\n"
                    f"Status: {row[9] or '-'}"
                ),
                markup=True,
                halign="left",
                valign="middle",
                font_size=16
            )

            info.bind(
                size=lambda instance, value:
                setattr(instance, "text_size", value)
            )

            button = Button(
                text="OPEN / EDIT RRU",
                size_hint_y=None,
                height=50
            )

            button.bind(
                on_release=lambda x,
                value=rru_id:
                self.edit_rru(value)
            )

            card.add_widget(info)
            card.add_widget(button)

            self.layout.add_widget(card)

    def add_rru(self, instance):

        global CURRENT_LTE_RRU_ID

        CURRENT_LTE_RRU_ID = None

        self.manager.current = "lte_rru_form"

    def edit_rru(self, rru_id):

        global CURRENT_LTE_RRU_ID

        CURRENT_LTE_RRU_ID = rru_id

        self.manager.current = "lte_rru_form"


# ============================================================
# LTE RRU FORM
# ============================================================

class LTERRUFormScreen(Screen):
    def on_pre_enter(self):
        self.build_form()

    def _dynamic_block(self, title, hint):
        block = BoxLayout(
            orientation="vertical",
            size_hint_y=None,
            height=84,
            spacing=4
        )
        label = Label(
            text=title,
            size_hint_y=None,
            height=30,
            halign="left",
            valign="middle",
            font_size=17
        )
        label.bind(size=lambda i, v: setattr(i, "text_size", v))
        field = create_input()
        field.hint_text = hint
        block.add_widget(label)
        block.add_widget(field)
        return block, field

    def _insert_after(self, block, reference_widget):
        if block.parent is not None:
            return
        parent = reference_widget.parent
        if parent is None:
            return
        index = parent.children.index(reference_widget)
        parent.add_widget(block, index=index)

    def _remove_dynamic(self, block):
        if block.parent is not None:
            block.parent.remove_widget(block)

    def _refresh_other_fields(self, *args):
        vendor_other = self.vendor.text == "Other"
        tech_other = self.technology.text == "Other"
        txrx_other = self.tx_rx.text == "Other"

        if vendor_other:
            self._insert_after(self.custom_vendor_block, self.vendor_field)
        else:
            self._remove_dynamic(self.custom_vendor_block)
            self.custom_vendor.text = ""

        if tech_other:
            self._insert_after(self.custom_technology_block, self.technology_field)
        else:
            self._remove_dynamic(self.custom_technology_block)
            self.custom_technology.text = ""

        if txrx_other:
            self._insert_after(self.custom_tx_rx_block, self.tx_rx_field)
        else:
            self._remove_dynamic(self.custom_tx_rx_block)
            self.custom_tx_rx.text = ""

    def build_form(self):
        self.clear_widgets()

        scroll = ScrollView()
        form = BoxLayout(
            orientation="vertical",
            padding=20,
            spacing=10,
            size_hint_y=None
        )
        form.bind(minimum_height=form.setter("height"))

        title = Label(
            text="LTE RRU INFORMATION",
            font_size=24,
            size_hint_y=None,
            height=65
        )

        self.rru_name = create_input()
        self.model = create_input()

        self.vendor = create_spinner([
            "HUAWEI",
            "NOKIA",
            "ERICSSON",
            "Other"
        ])
        self.vendor.bind(text=self._refresh_other_fields)

        self.band = create_input()

        self.technology = create_spinner([
            "LTE",
            "LTE + GSM",
            "LTE + NR",
            "Other"
        ])
        self.technology.bind(text=self._refresh_other_fields)

        self.tx_rx = create_spinner([
            "2T2R",
            "4T4R",
            "8T8R",
            "Other"
        ])
        self.tx_rx.bind(text=self._refresh_other_fields)

        self.sector = create_input()
        self.cpri = create_input()

        self.status = create_spinner([
            "Active",
            "Standby",
            "Faulty",
            "Decommissioned"
        ])

        self.notes = create_input(multiline=True, height=110)

        # Normal fields are kept in separate containers so each custom
        # Other field can be inserted immediately below its related field.
        def field_container(title_text, widget):
            box = BoxLayout(
                orientation="vertical",
                size_hint_y=None,
                height=84,
                spacing=4
            )
            label = Label(
                text=title_text,
                size_hint_y=None,
                height=30,
                halign="left",
                valign="middle",
                font_size=17
            )
            label.bind(size=lambda i, v: setattr(i, "text_size", v))
            box.add_widget(label)
            box.add_widget(widget)
            return box

        self.vendor_field = field_container("Vendor", self.vendor)
        self.technology_field = field_container("Technology", self.technology)
        self.tx_rx_field = field_container("TX / RX", self.tx_rx)

        self.custom_vendor_block, self.custom_vendor = self._dynamic_block(
            "Custom Vendor Name",
            "Enter the vendor name"
        )
        self.custom_technology_block, self.custom_technology = self._dynamic_block(
            "Custom Technology",
            "Enter the technology"
        )
        self.custom_tx_rx_block, self.custom_tx_rx = self._dynamic_block(
            "Custom TX / RX",
            "Example: 16T16R, 4T8R"
        )

        form.add_widget(title)
        add_field(form, "RRU Name", self.rru_name)
        add_field(form, "RRU Model", self.model)
        form.add_widget(self.vendor_field)
        add_field(form, "Frequency Band", self.band)
        form.add_widget(self.technology_field)
        form.add_widget(self.tx_rx_field)
        add_field(form, "Sector", self.sector)
        add_field(form, "CPRI Port", self.cpri)
        add_field(form, "Status", self.status)
        add_field(form, "Notes", self.notes)

        save = Button(
            text="SAVE / UPDATE RRU",
            size_hint_y=None,
            height=60
        )

        delete = Button(
            text="DELETE RRU",
            size_hint_y=None,
            height=60
        )

        back = Button(
            text="BACK TO RRU LIST",
            size_hint_y=None,
            height=60
        )

        save.bind(on_release=self.save_rru)
        delete.bind(on_release=self.ask_delete)
        back.bind(
            on_release=lambda x: setattr(
                self.manager,
                "current",
                "lte_rru_list"
            )
        )

        form.add_widget(save)
        form.add_widget(delete)
        form.add_widget(back)

        scroll.add_widget(form)
        self.add_widget(scroll)

        self._refresh_other_fields()
        self.load_rru()
        self._refresh_other_fields()

    def load_rru(self):
        if CURRENT_LTE_RRU_ID is None:
            return

        connection = get_connection()
        cursor = connection.cursor()

        cursor.execute(
            "SELECT * FROM lte_rru WHERE id=?",
            (CURRENT_LTE_RRU_ID,)
        )

        row = cursor.fetchone()
        connection.close()

        if not row:
            return

        self.rru_name.text = row[2] or ""
        self.model.text = row[3] or ""
        self.vendor.text = row[4] or "Select"
        self.band.text = row[5] or ""
        self.technology.text = row[6] or "Select"
        self.tx_rx.text = row[7] or "Select"
        self.sector.text = row[8] or ""
        self.cpri.text = row[9] or ""
        self.status.text = row[10] or "Active"
        self.notes.text = row[11] or ""

        # New custom columns are appended to the existing table, preserving
        # all old column positions and compatibility with existing databases.
        self.custom_vendor.text = row[12] if len(row) > 12 and row[12] else ""
        self.custom_technology.text = row[13] if len(row) > 13 and row[13] else ""
        self.custom_tx_rx.text = row[14] if len(row) > 14 and row[14] else ""

    def save_rru(self, instance):
        global CURRENT_LTE_RRU_ID

        if not self.rru_name.text.strip():
            show_message("Error", "RRU Name is required.")
            return

        if self.vendor.text == "Other" and not self.custom_vendor.text.strip():
            show_message("Error", "Please enter the Custom Vendor Name.")
            return

        if self.technology.text == "Other" and not self.custom_technology.text.strip():
            show_message("Error", "Please enter the Custom Technology.")
            return

        if self.tx_rx.text == "Other" and not self.custom_tx_rx.text.strip():
            show_message("Error", "Please enter the Custom TX / RX.")
            return

        values = (
            CURRENT_LTE_ID,
            self.rru_name.text.strip(),
            self.model.text.strip(),
            self.vendor.text,
            self.band.text.strip(),
            self.technology.text,
            self.tx_rx.text,
            self.sector.text.strip(),
            self.cpri.text.strip(),
            self.status.text,
            self.notes.text.strip(),
            self.custom_vendor.text.strip() if self.vendor.text == "Other" else "",
            self.custom_technology.text.strip() if self.technology.text == "Other" else "",
            self.custom_tx_rx.text.strip() if self.tx_rx.text == "Other" else ""
        )

        connection = get_connection()
        cursor = connection.cursor()

        try:
            if CURRENT_LTE_RRU_ID is None:
                cursor.execute("""
                INSERT INTO lte_rru (
                    lte_system_id,
                    rru_name,
                    rru_model,
                    vendor,
                    frequency_band,
                    technology,
                    tx_rx,
                    sector,
                    cpri_port,
                    status,
                    notes,
                    custom_vendor,
                    custom_technology,
                    custom_tx_rx
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, values)

                CURRENT_LTE_RRU_ID = cursor.lastrowid

            else:
                cursor.execute("""
                UPDATE lte_rru SET
                    rru_name=?,
                    rru_model=?,
                    vendor=?,
                    frequency_band=?,
                    technology=?,
                    tx_rx=?,
                    sector=?,
                    cpri_port=?,
                    status=?,
                    notes=?,
                    custom_vendor=?,
                    custom_technology=?,
                    custom_tx_rx=?
                WHERE id=?
                """,
                values[1:] + (CURRENT_LTE_RRU_ID,))

            connection.commit()
            show_message("Success", "LTE RRU saved successfully.")
            self.manager.current = "lte_rru_list"

        except Exception as e:
            connection.rollback()
            show_message("Database Error", str(e))

        finally:
            connection.close()

    def ask_delete(self, instance):
        if CURRENT_LTE_RRU_ID is None:
            return

        content = BoxLayout(
            orientation="vertical",
            padding=15,
            spacing=10
        )

        message = Label(
            text="Delete this LTE RRU?",
            font_size=18
        )

        buttons = BoxLayout(
            size_hint_y=None,
            height=55,
            spacing=10
        )

        yes = Button(text="DELETE")
        no = Button(text="CANCEL")

        buttons.add_widget(yes)
        buttons.add_widget(no)
        content.add_widget(message)
        content.add_widget(buttons)

        popup = Popup(
            title="CONFIRM DELETE",
            content=content,
            size_hint=(0.85, None),
            height=180,
            auto_dismiss=False
        )

        yes.bind(on_release=lambda x: self.delete_rru(popup))
        no.bind(on_release=popup.dismiss)
        popup.open()

    def delete_rru(self, popup):
        global CURRENT_LTE_RRU_ID

        connection = get_connection()
        cursor = connection.cursor()

        try:
            cursor.execute(
                "DELETE FROM lte_rru WHERE id=?",
                (CURRENT_LTE_RRU_ID,)
            )
            connection.commit()
            CURRENT_LTE_RRU_ID = None
            popup.dismiss()
            self.manager.current = "lte_rru_list"

        except Exception as e:
            connection.rollback()
            popup.dismiss()
            show_message("Database Error", str(e))

        finally:
            connection.close()


# ============================================================
# LTE CELLS LIST
# ============================================================

class LTECellsListScreen(Screen):

    def on_pre_enter(self):
        self.build_screen()

    def build_screen(self):

        self.clear_widgets()

        main = BoxLayout(
            orientation="vertical",
            padding=10,
            spacing=10
        )

        title = Label(
            text="LTE CELLS / SECTORS",
            font_size=24,
            size_hint_y=None,
            height=65
        )

        add = Button(
            text="ADD NEW LTE CELL / SECTOR",
            size_hint_y=None,
            height=60,
            font_size=18
        )

        add.bind(
            on_release=self.add_cell
        )

        scroll = ScrollView()

        self.layout = BoxLayout(
            orientation="vertical",
            spacing=15,
            padding=5,
            size_hint_y=None
        )

        self.layout.bind(
            minimum_height=
            self.layout.setter("height")
        )

        scroll.add_widget(self.layout)

        back = Button(
            text="BACK TO LTE SYSTEM",
            size_hint_y=None,
            height=60
        )

        back.bind(
            on_release=lambda x:
            setattr(
                self.manager,
                "current",
                "lte_form"
            )
        )

        main.add_widget(title)
        main.add_widget(add)
        main.add_widget(scroll)
        main.add_widget(back)

        self.add_widget(main)

        self.display_cells()

    def display_cells(self):

        self.layout.clear_widgets()

        connection = get_connection()
        cursor = connection.cursor()

        cursor.execute("""
        SELECT
            id,
            cell_number,
            cell_name,
            sector,
            pci,
            tac,
            earfcn,
            bandwidth,
            frequency_band,
            mimo,
            antenna_type,
            azimuth,
            status
        FROM lte_cells
        WHERE lte_system_id=?
        ORDER BY cell_number
        """, (CURRENT_LTE_ID,))

        rows = cursor.fetchall()

        connection.close()

        if not rows:

            self.layout.add_widget(
                Label(
                    text="No LTE Cells Added",
                    size_hint_y=None,
                    height=70
                )
            )

            return

        for row in rows:

            cell_id = row[0]

            card = BoxLayout(
                orientation="vertical",
                size_hint_y=None,
                height=300,
                padding=12,
                spacing=8
            )

            info = Label(
                text=(
                    f"[b]CELL {row[1]} - "
                    f"{row[2] or ''}[/b]\n"
                    f"Sector: {row[3] or '-'}\n"
                    f"PCI: {row[4] or '-'}\n"
                    f"TAC: {row[5] or '-'}\n"
                    f"EARFCN: {row[6] or '-'}\n"
                    f"Bandwidth: {row[7] or '-'}\n"
                    f"Band: {row[8] or '-'}\n"
                    f"MIMO: {row[9] or '-'}\n"
                    f"Antenna: {row[10] or '-'}\n"
                    f"Azimuth: {row[11] or '-'}\n"
                    f"Status: {row[12] or '-'}"
                ),
                markup=True,
                halign="left",
                valign="middle",
                font_size=15
            )

            info.bind(
                size=lambda instance, value:
                setattr(instance, "text_size", value)
            )

            button = Button(
                text="OPEN / EDIT CELL",
                size_hint_y=None,
                height=50
            )

            button.bind(
                on_release=lambda x,
                value=cell_id:
                self.edit_cell(value)
            )

            card.add_widget(info)
            card.add_widget(button)

            self.layout.add_widget(card)

    def add_cell(self, instance):

        global CURRENT_LTE_CELL_ID

        CURRENT_LTE_CELL_ID = None

        self.manager.current = "lte_cell_form"

    def edit_cell(self, cell_id):

        global CURRENT_LTE_CELL_ID

        CURRENT_LTE_CELL_ID = cell_id

        self.manager.current = "lte_cell_form"


# ============================================================
# LTE CELL FORM
# ============================================================

class LTECellFormScreen(Screen):

    def on_pre_enter(self):
        self.build_form()

    def build_form(self):

        self.clear_widgets()

        scroll = ScrollView()

        form = BoxLayout(
            orientation="vertical",
            padding=20,
            spacing=10,
            size_hint_y=None
        )

        form.bind(
            minimum_height=
            form.setter("height")
        )

        title = Label(
            text="LTE CELL / SECTOR INFORMATION",
            font_size=24,
            size_hint_y=None,
            height=65
        )

        self.cell_number = create_input()
        self.cell_name = create_input()
        self.sector = create_input()
        self.pci = create_input()
        self.tac = create_input()
        self.earfcn = create_input()

        self.bandwidth = create_spinner([
            "1.4 MHz",
            "3 MHz",
            "5 MHz",
            "10 MHz",
            "15 MHz",
            "20 MHz"
        ])

        self.frequency_band = create_spinner([
            "LTE 700 MHz",
            "LTE 800 MHz",
            "LTE 900 MHz",
            "LTE 1800 MHz",
            "LTE 2100 MHz",
            "LTE 2600 MHz",
            "Other"
        ])

        self.mimo = create_spinner([
            "2x2",
            "4x2",
            "4x4",
            "8x8",
            "8T8R",
            "Other"
        ])

        self.antenna = create_input()
        self.azimuth = create_input()

        self.status = create_spinner([
            "Active",
            "Locked",
            "Down",
            "Maintenance"
        ])

        self.notes = create_input(
            multiline=True,
            height=110
        )

        add_field(form, "Cell Number", self.cell_number)
        add_field(form, "Cell Name", self.cell_name)
        add_field(form, "Sector", self.sector)
        add_field(form, "PCI", self.pci)
        add_field(form, "TAC", self.tac)
        add_field(form, "EARFCN", self.earfcn)
        add_field(form, "Bandwidth", self.bandwidth)
        add_field(form, "Frequency Band", self.frequency_band)
        add_field(form, "MIMO", self.mimo)
        add_field(form, "Antenna Type", self.antenna)
        add_field(form, "Azimuth", self.azimuth)
        add_field(form, "Status", self.status)
        add_field(form, "Notes", self.notes)

        save = Button(
            text="SAVE / UPDATE CELL",
            size_hint_y=None,
            height=60
        )

        delete = Button(
            text="DELETE CELL",
            size_hint_y=None,
            height=60
        )

        back = Button(
            text="BACK TO LTE CELLS",
            size_hint_y=None,
            height=60
        )

        save.bind(on_release=self.save_cell)
        delete.bind(on_release=self.ask_delete)

        back.bind(
            on_release=lambda x:
            setattr(
                self.manager,
                "current",
                "lte_cells_list"
            )
        )

        form.add_widget(title)
        form.add_widget(save)
        form.add_widget(delete)
        form.add_widget(back)

        scroll.add_widget(form)
        self.add_widget(scroll)

        self.load_cell()

    def load_cell(self):

        if CURRENT_LTE_CELL_ID is None:
            return

        connection = get_connection()
        cursor = connection.cursor()

        cursor.execute(
            "SELECT * FROM lte_cells WHERE id=?",
            (CURRENT_LTE_CELL_ID,)
        )

        row = cursor.fetchone()
        connection.close()

        if not row:
            return

        self.cell_number.text = str(row[2] or "")
        self.cell_name.text = row[3] or ""
        self.sector.text = row[4] or ""

        self.pci.text = row[5] or ""
        self.tac.text = row[6] or ""
        self.earfcn.text = row[7] or ""

        self.bandwidth.text = row[8] or "Select"
        self.frequency_band.text = row[9] or "Select"

        self.mimo.text = row[10] or "Select"

        self.antenna.text = row[11] or ""
        self.azimuth.text = row[12] or ""

        self.status.text = row[13] or "Active"
        self.notes.text = row[14] or ""

    def save_cell(self, instance):

        global CURRENT_LTE_CELL_ID

        if not self.cell_number.text.strip():

            show_message(
                "Error",
                "Cell Number is required."
            )

            return

        values = (
            CURRENT_LTE_ID,
            self.cell_number.text.strip(),
            self.cell_name.text.strip(),
            self.sector.text.strip(),
            self.pci.text.strip(),
            self.tac.text.strip(),
            self.earfcn.text.strip(),
            self.bandwidth.text,
            self.frequency_band.text,
            self.mimo.text,
            self.antenna.text.strip(),
            self.azimuth.text.strip(),
            self.status.text,
            self.notes.text.strip()
        )

        connection = get_connection()
        cursor = connection.cursor()

        try:

            if CURRENT_LTE_CELL_ID is None:

                cursor.execute("""
                INSERT INTO lte_cells (
                    lte_system_id,
                    cell_number,
                    cell_name,
                    sector,
                    pci,
                    tac,
                    earfcn,
                    bandwidth,
                    frequency_band,
                    mimo,
                    antenna_type,
                    azimuth,
                    status,
                    notes
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, values)

                CURRENT_LTE_CELL_ID = cursor.lastrowid

            else:

                cursor.execute("""
                UPDATE lte_cells SET
                    cell_number=?,
                    cell_name=?,
                    sector=?,
                    pci=?,
                    tac=?,
                    earfcn=?,
                    bandwidth=?,
                    frequency_band=?,
                    mimo=?,
                    antenna_type=?,
                    azimuth=?,
                    status=?,
                    notes=?
                WHERE id=?
                """,
                values[1:] + (CURRENT_LTE_CELL_ID,))

            connection.commit()

        except Exception as error:

            connection.rollback()
            connection.close()

            show_message(
                "Database Error",
                str(error)
            )

            return

        connection.close()

        show_message(
            "Success",
            "LTE Cell saved successfully."
        )

    def ask_delete(self, instance):

        confirm_delete(
            "Delete LTE Cell",
            "Are you sure you want to delete this LTE Cell?",
            self.delete_cell
        )

    def delete_cell(self):

        global CURRENT_LTE_CELL_ID

        connection = get_connection()
        cursor = connection.cursor()

        cursor.execute(
            "DELETE FROM lte_cells WHERE id=?",
            (CURRENT_LTE_CELL_ID,)
        )

        connection.commit()
        connection.close()

        CURRENT_LTE_CELL_ID = None

        self.manager.current = "lte_cells_list"


# ============================================================
# MICROWAVE LIST
# ============================================================

class MicrowaveListScreen(Screen):

    def on_pre_enter(self):
        self.build_screen()

    def build_screen(self):

        self.clear_widgets()

        main = BoxLayout(
            orientation="vertical",
            padding=10,
            spacing=10
        )

        title = Label(
            text=f"MICROWAVE LINKS\n{CURRENT_SITE_NAME}",
            font_size=23,
            size_hint_y=None,
            height=80
        )

        add = Button(
            text="ADD NEW MICROWAVE LINK",
            size_hint_y=None,
            height=60,
            font_size=18
        )

        add.bind(
            on_release=self.add_link
        )

        scroll = ScrollView()

        self.links_layout = BoxLayout(
            orientation="vertical",
            spacing=15,
            padding=[5, 5],
            size_hint_y=None
        )

        self.links_layout.bind(
            minimum_height=
            self.links_layout.setter("height")
        )

        scroll.add_widget(self.links_layout)

        back = Button(
            text="BACK TO SYSTEMS DASHBOARD",
            size_hint_y=None,
            height=60
        )

        back.bind(
            on_release=lambda x:
            setattr(
                self.manager,
                "current",
                "systems"
            )
        )

        main.add_widget(title)
        main.add_widget(add)
        main.add_widget(scroll)
        main.add_widget(back)

        self.add_widget(main)

        self.display_links()

    def display_links(self):

        self.links_layout.clear_widgets()

        connection = get_connection()
        cursor = connection.cursor()

        cursor.execute("""
        SELECT
            id,
            link_name,
            equipment_vendor,
            idu_type,
            odu_type,
            tx_frequency,
            rx_frequency,
            capacity,
            received_power,
            tx_power,
            remote_site,
            link_id
        FROM microwave
        WHERE site_id=?
        ORDER BY id DESC
        """, (CURRENT_SITE_ID,))

        rows = cursor.fetchall()

        connection.close()

        if not rows:

            self.links_layout.add_widget(
                Label(
                    text="No Microwave Links Added",
                    size_hint_y=None,
                    height=70,
                    font_size=18
                )
            )

            return

        for row in rows:

            link_id = row[0]

            card = BoxLayout(
                orientation="vertical",
                size_hint_y=None,
                height=225,
                padding=12,
                spacing=8
            )

            info = Label(
                text=(
                    f"[b]LINK: {row[1] or '-'}[/b]\n"
                    f"LINK ID: {row[11] or '-'}\n"
                    f"Vendor: {row[2] or '-'}\n"
                    f"IDU: {row[3] or '-'}\n"
                    f"ODU: {row[4] or '-'}\n"
                    f"TX Frequency: {row[5] or '-'}\n"
                    f"RX Frequency: {row[6] or '-'}\n"
                    f"Capacity: {row[7] or '-'}\n"
                    f"Received Power: {row[8] or '-'}\n"
                    f"TX Power: {row[9] or '-'}\n"
                    f"Remote Site: {row[10] or '-'}"
                ),
                markup=True,
                halign="left",
                valign="middle",
                font_size=16
            )

            info.bind(
                size=lambda instance, value:
                setattr(instance, "text_size", value)
            )

            button = Button(
                text="OPEN / EDIT LINK",
                size_hint_y=None,
                height=50
            )

            button.bind(
                on_release=lambda x,
                value=link_id:
                self.open_link(value)
            )

            card.add_widget(info)
            card.add_widget(button)

            self.links_layout.add_widget(card)

    def add_link(self, instance):

        global CURRENT_MICROWAVE_ID

        CURRENT_MICROWAVE_ID = None

        self.manager.current = "microwave_form"

    def open_link(self, link_id):

        global CURRENT_MICROWAVE_ID

        CURRENT_MICROWAVE_ID = link_id

        self.manager.current = "microwave_form"


# ============================================================
# MICROWAVE FORM
# ============================================================

class MicrowaveFormScreen(Screen):

    def on_pre_enter(self):
        self.build_form()

    def _make_dynamic_block(self):
        return BoxLayout(
            orientation="vertical",
            size_hint_y=None,
            height=82,
            spacing=4
        )

    def _field_label(self, text):
        return Label(
            text=text,
            size_hint_y=None,
            height=32,
            halign="left",
            valign="middle",
            font_size=17
        )

    def _insert_dynamic_block(self, block, after_widget):
        if block.parent is not None:
            return
        parent = after_widget.parent
        if parent is None:
            return
        index = parent.children.index(after_widget)
        parent.add_widget(block, index=index)

    def _remove_dynamic_block(self, block):
        if block.parent is not None:
            block.parent.remove_widget(block)

    def build_form(self):

        self.clear_widgets()

        scroll = ScrollView()

        form = BoxLayout(
            orientation="vertical",
            padding=20,
            spacing=10,
            size_hint_y=None
        )

        form.bind(
            minimum_height=form.setter("height")
        )

        title = Label(
            text="MICROWAVE LINK INFORMATION",
            font_size=23,
            size_hint_y=None,
            height=65
        )

        self.link_id = create_input()
        self.link_name = create_input()

        self.vendor = create_spinner([
            "NOKIA",
            "HUAWEI",
            "Other"
        ])

        self.custom_vendor = create_input()
        self.custom_vendor.hint_text = "Enter exact Vendor"
        self.custom_vendor_block = self._make_dynamic_block()
        self.custom_vendor_block.add_widget(self._field_label("Custom Vendor"))
        self.custom_vendor_block.add_widget(self.custom_vendor)

        self.idu = create_spinner([
            "SRAL L",
            "SRAL XD",
            "RTN 905",
            "RTN 910",
            "RTN 950",
            "RTN 950A",
            "RTN 980",
            "Other"
        ])

        self.custom_idu = create_input()
        self.custom_idu.hint_text = "Enter exact IDU Type"
        self.custom_idu_block = self._make_dynamic_block()
        self.custom_idu_block.add_widget(self._field_label("Custom IDU Type"))
        self.custom_idu_block.add_widget(self.custom_idu)

        self.odu = create_input()
        self.odu.hint_text = "ODU Type / Model"

        self.tx_frequency = create_input()
        self.tx_frequency.hint_text = "TX Frequency"

        self.rx_frequency = create_input()
        self.rx_frequency.hint_text = "RX Frequency"

        self.capacity = create_spinner([
            "4 E1",
            "8 E1",
            "16 E1",
            "32 E1",
            "Other"
        ])

        self.custom_capacity = create_input()
        self.custom_capacity.hint_text = "Enter exact Capacity"
        self.custom_capacity_block = self._make_dynamic_block()
        self.custom_capacity_block.add_widget(self._field_label("Custom Capacity"))
        self.custom_capacity_block.add_widget(self.custom_capacity)

        self.dish_size = create_input()

        self.polarization = create_spinner([
            "Horizontal",
            "Vertical",
            "Dual",
            "Other"
        ])

        self.custom_polarization = create_input()
        self.custom_polarization.hint_text = "Enter exact Polarization"
        self.custom_polarization_block = self._make_dynamic_block()
        self.custom_polarization_block.add_widget(self._field_label("Custom Polarization"))
        self.custom_polarization_block.add_widget(self.custom_polarization)

        self.received_power = create_input()
        self.received_power.hint_text = "Example: -45 dBm"

        self.tx_power = create_input()
        self.tx_power.hint_text = "Example: 20 dBm"

        self.remote_site = create_input()

        self.notes = create_input(
            multiline=True,
            height=110
        )

        form.add_widget(title)
        add_field(form, "LINK ID", self.link_id)
        add_field(form, "Link Name", self.link_name)
        add_field(form, "Vendor", self.vendor)
        form.add_widget(self.custom_vendor_block)
        add_field(form, "IDU Type", self.idu)
        form.add_widget(self.custom_idu_block)
        add_field(form, "ODU Type / Model", self.odu)
        add_field(form, "TX Frequency", self.tx_frequency)
        add_field(form, "RX Frequency", self.rx_frequency)
        add_field(form, "Capacity", self.capacity)
        form.add_widget(self.custom_capacity_block)
        add_field(form, "Dish Size", self.dish_size)
        add_field(form, "Polarization", self.polarization)
        form.add_widget(self.custom_polarization_block)
        add_field(form, "Received Power", self.received_power)
        add_field(form, "TX Power", self.tx_power)
        add_field(form, "Remote Site", self.remote_site)
        add_field(form, "Notes", self.notes)

        save = Button(
            text="SAVE / UPDATE LINK",
            size_hint_y=None,
            height=60
        )

        e1 = Button(
            text="MANAGE E1s",
            size_hint_y=None,
            height=60
        )

        delete = Button(
            text="DELETE LINK",
            size_hint_y=None,
            height=60
        )

        back = Button(
            text="BACK TO MICROWAVE LIST",
            size_hint_y=None,
            height=60
        )

        self.vendor.bind(text=self._on_vendor_changed)
        self.idu.bind(text=self._on_idu_changed)
        self.capacity.bind(text=self._on_capacity_changed)
        self.polarization.bind(text=self._on_polarization_changed)

        save.bind(on_release=self.save_link)
        e1.bind(on_release=self.open_e1s)
        delete.bind(on_release=self.ask_delete)
        back.bind(
            on_release=lambda x: setattr(
                self.manager,
                "current",
                "microwave_list"
            )
        )

        form.add_widget(save)
        form.add_widget(e1)
        form.add_widget(delete)
        form.add_widget(back)

        scroll.add_widget(form)
        self.add_widget(scroll)

        self.refresh_dynamic_fields()
        self.load_link()

    def refresh_dynamic_fields(self):
        if self.vendor.text == "Other":
            self._insert_dynamic_block(self.custom_vendor_block, self.vendor)
        else:
            self._remove_dynamic_block(self.custom_vendor_block)
            self.custom_vendor.text = ""

        if self.idu.text == "Other":
            self._insert_dynamic_block(self.custom_idu_block, self.idu)
        else:
            self._remove_dynamic_block(self.custom_idu_block)
            self.custom_idu.text = ""

        if self.capacity.text == "Other":
            self._insert_dynamic_block(self.custom_capacity_block, self.capacity)
        else:
            self._remove_dynamic_block(self.custom_capacity_block)
            self.custom_capacity.text = ""

        if self.polarization.text == "Other":
            self._insert_dynamic_block(self.custom_polarization_block, self.polarization)
        else:
            self._remove_dynamic_block(self.custom_polarization_block)
            self.custom_polarization.text = ""

    def _on_vendor_changed(self, instance, value):
        self.refresh_dynamic_fields()

    def _on_idu_changed(self, instance, value):
        self.refresh_dynamic_fields()

    def _on_capacity_changed(self, instance, value):
        self.refresh_dynamic_fields()

    def _on_polarization_changed(self, instance, value):
        self.refresh_dynamic_fields()

    def load_link(self):

        if CURRENT_MICROWAVE_ID is None:
            return

        connection = get_connection()
        cursor = connection.cursor()

        cursor.execute(
            "SELECT * FROM microwave WHERE id=?",
            (CURRENT_MICROWAVE_ID,)
        )

        row = cursor.fetchone()
        connection.close()

        if not row:
            return

        # Original columns 0-11 remain unchanged.
        self.link_name.text = row[2] or ""
        self.vendor.text = row[3] or "Select"
        self.idu.text = row[4] or "Select"
        self.odu.text = row[5] or ""

        # Legacy frequency is retained as a fallback for TX frequency.
        legacy_frequency = row[6] or ""
        self.capacity.text = row[7] or "Select"
        self.dish_size.text = row[8] or ""
        self.polarization.text = row[9] or "Select"
        self.remote_site.text = row[10] or ""
        self.notes.text = row[11] or ""

        # New columns are appended after the legacy columns.
        self.link_id.text = row[12] or "" if len(row) > 12 else ""
        custom_vendor = row[13] or "" if len(row) > 13 else ""
        custom_idu = row[14] or "" if len(row) > 14 else ""
        custom_pol = row[15] or "" if len(row) > 15 else ""
        custom_capacity = row[16] or "" if len(row) > 16 else ""
        tx_frequency = row[17] or "" if len(row) > 17 else ""
        rx_frequency = row[18] or "" if len(row) > 18 else ""
        received_power = row[19] or "" if len(row) > 19 else ""
        tx_power = row[20] or "" if len(row) > 20 else ""

        self.custom_vendor.text = custom_vendor
        self.custom_idu.text = custom_idu
        self.custom_polarization.text = custom_pol
        self.custom_capacity.text = custom_capacity
        self.tx_frequency.text = tx_frequency or legacy_frequency
        self.rx_frequency.text = rx_frequency
        self.received_power.text = received_power
        self.tx_power.text = tx_power

        self.refresh_dynamic_fields()

    def save_link(self, instance):

        global CURRENT_MICROWAVE_ID

        if not self.link_id.text.strip():
            show_message("Error", "LINK ID is required.")
            return

        if not self.link_name.text.strip():
            show_message("Error", "Link Name is required.")
            return

        if self.vendor.text == "Select":
            show_message("Error", "Please select Vendor.")
            return

        if self.vendor.text == "Other" and not self.custom_vendor.text.strip():
            show_message("Error", "Please enter the Custom Vendor.")
            return

        if self.idu.text == "Select":
            show_message("Error", "Please select IDU.")
            return

        if self.idu.text == "Other" and not self.custom_idu.text.strip():
            show_message("Error", "Please enter the Custom IDU Type.")
            return

        if self.capacity.text == "Select":
            show_message("Error", "Please select Capacity.")
            return

        if self.capacity.text == "Other" and not self.custom_capacity.text.strip():
            show_message("Error", "Please enter the Custom Capacity.")
            return

        if self.polarization.text == "Select":
            show_message("Error", "Please select Polarization.")
            return

        if self.polarization.text == "Other" and not self.custom_polarization.text.strip():
            show_message("Error", "Please enter the Custom Polarization.")
            return

        values = (
            CURRENT_SITE_ID,
            self.link_name.text.strip(),
            self.vendor.text,
            self.idu.text,
            self.odu.text.strip(),
            self.tx_frequency.text.strip(),
            self.capacity.text,
            self.dish_size.text.strip(),
            self.polarization.text,
            self.remote_site.text.strip(),
            self.notes.text.strip(),
            self.link_id.text.strip(),
            self.custom_vendor.text.strip(),
            self.custom_idu.text.strip(),
            self.custom_polarization.text.strip(),
            self.custom_capacity.text.strip(),
            self.tx_frequency.text.strip(),
            self.rx_frequency.text.strip(),
            self.received_power.text.strip(),
            self.tx_power.text.strip()
        )

        connection = get_connection()
        cursor = connection.cursor()

        try:
            if CURRENT_MICROWAVE_ID is None:
                cursor.execute("""
                INSERT INTO microwave (
                    site_id,
                    link_name,
                    equipment_vendor,
                    idu_type,
                    odu_type,
                    frequency,
                    capacity,
                    dish_size,
                    polarization,
                    remote_site,
                    notes,
                    link_id,
                    custom_vendor,
                    custom_idu_type,
                    custom_polarization,
                    custom_capacity,
                    tx_frequency,
                    rx_frequency,
                    received_power,
                    tx_power
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, values)

                CURRENT_MICROWAVE_ID = cursor.lastrowid

            else:
                cursor.execute("""
                UPDATE microwave SET
                    link_name=?,
                    equipment_vendor=?,
                    idu_type=?,
                    odu_type=?,
                    frequency=?,
                    capacity=?,
                    dish_size=?,
                    polarization=?,
                    remote_site=?,
                    notes=?,
                    link_id=?,
                    custom_vendor=?,
                    custom_idu_type=?,
                    custom_polarization=?,
                    custom_capacity=?,
                    tx_frequency=?,
                    rx_frequency=?,
                    received_power=?,
                    tx_power=?
                WHERE id=?
                """, values[1:] + (CURRENT_MICROWAVE_ID,))

            connection.commit()

        except Exception as error:
            connection.rollback()
            connection.close()
            show_message("Database Error", str(error))
            return

        connection.close()

        self.create_e1s()

        show_message(
            "Success",
            "Microwave Link saved successfully."
        )

    def create_e1s(self):

        capacity_map = {
            "4 E1": 4,
            "8 E1": 8,
            "16 E1": 16,
            "32 E1": 32
        }

        total = capacity_map.get(self.capacity.text)

        if not total:
            return

        connection = get_connection()
        cursor = connection.cursor()

        cursor.execute("""
        SELECT e1_number
        FROM e1_links
        WHERE microwave_id=?
        """, (CURRENT_MICROWAVE_ID,))

        existing = [row[0] for row in cursor.fetchall()]

        for number in range(1, total + 1):
            if number not in existing:
                cursor.execute("""
                INSERT INTO e1_links (
                    microwave_id,
                    e1_number,
                    service_name,
                    destination,
                    status,
                    notes
                )
                VALUES (?, ?, ?, ?, ?, ?)
                """, (
                    CURRENT_MICROWAVE_ID,
                    number,
                    "",
                    "",
                    "Available",
                    ""
                ))

        connection.commit()
        connection.close()

    def open_e1s(self, instance):

        if CURRENT_MICROWAVE_ID is None:
            show_message(
                "Save Link First",
                "Please save the Microwave Link first."
            )
            return

        self.manager.current = "e1_list"

    def ask_delete(self, instance):

        if CURRENT_MICROWAVE_ID is None:
            return

        confirm_delete(
            "Delete Microwave Link",
            "Delete this link and all its E1s?",
            self.delete_link
        )

    def delete_link(self):

        global CURRENT_MICROWAVE_ID

        connection = get_connection()
        cursor = connection.cursor()

        cursor.execute(
            "DELETE FROM e1_links WHERE microwave_id=?",
            (CURRENT_MICROWAVE_ID,)
        )

        cursor.execute(
            "DELETE FROM microwave WHERE id=?",
            (CURRENT_MICROWAVE_ID,)
        )

        connection.commit()
        connection.close()

        CURRENT_MICROWAVE_ID = None

        self.manager.current = "microwave_list"


# ============================================================
# E1 LIST
# ============================================================

class E1ListScreen(Screen):

    def on_pre_enter(self):
        self.build_screen()

    def build_screen(self):

        self.clear_widgets()

        main = BoxLayout(
            orientation="vertical",
            padding=10,
            spacing=10
        )

        title = Label(
            text="E1 MANAGEMENT",
            font_size=24,
            size_hint_y=None,
            height=65
        )

        scroll = ScrollView()

        self.e1_layout = BoxLayout(
            orientation="vertical",
            spacing=15,
            padding=[5, 5],
            size_hint_y=None
        )

        self.e1_layout.bind(
            minimum_height=
            self.e1_layout.setter("height")
        )

        scroll.add_widget(self.e1_layout)

        back = Button(
            text="BACK TO MICROWAVE LINK",
            size_hint_y=None,
            height=60
        )

        back.bind(
            on_release=lambda x:
            setattr(
                self.manager,
                "current",
                "microwave_form"
            )
        )

        main.add_widget(title)
        main.add_widget(scroll)
        main.add_widget(back)

        self.add_widget(main)

        self.display_e1s()

    def display_e1s(self):

        self.e1_layout.clear_widgets()

        connection = get_connection()
        cursor = connection.cursor()

        cursor.execute("""
        SELECT
            id,
            e1_number,
            service_name,
            destination,
            status,
            notes
        FROM e1_links
        WHERE microwave_id=?
        ORDER BY e1_number
        """, (CURRENT_MICROWAVE_ID,))

        rows = cursor.fetchall()

        connection.close()

        if not rows:

            self.e1_layout.add_widget(
                Label(
                    text="No E1s Available",
                    size_hint_y=None,
                    height=70
                )
            )

            return

        for row in rows:

            e1_id = row[0]

            card = BoxLayout(
                orientation="vertical",
                size_hint_y=None,
                height=205,
                padding=12,
                spacing=8
            )

            label = Label(
                text=(
                    f"[b]E1 {row[1]}[/b]\n"
                    f"Service: {row[2] or 'Available'}\n"
                    f"Destination: {row[3] or '-'}\n"
                    f"Status: {row[4] or '-'}"
                ),
                markup=True,
                halign="left",
                valign="middle",
                font_size=17
            )

            label.bind(
                size=lambda instance, value:
                setattr(instance, "text_size", value)
            )

            edit = Button(
                text="EDIT E1",
                size_hint_y=None,
                height=50
            )

            edit.bind(
                on_release=lambda x,
                value=e1_id:
                self.edit_e1(value)
            )

            card.add_widget(label)
            card.add_widget(edit)

            self.e1_layout.add_widget(card)

    def edit_e1(self, e1_id):

        E1EditPopup(e1_id).open()


# ============================================================
# E1 POPUP
# ============================================================

class E1EditPopup(Popup):

    def __init__(self, e1_id, **kwargs):

        super().__init__(**kwargs)

        self.e1_id = e1_id

        self.title = "EDIT E1"

        self.size_hint = (0.92, 0.82)

        scroll = ScrollView()

        layout = BoxLayout(
            orientation="vertical",
            padding=15,
            spacing=10,
            size_hint_y=None
        )

        layout.bind(
            minimum_height=
            layout.setter("height")
        )

        self.e1_number = create_input()
        self.service = create_input()
        self.destination = create_input()

        self.status = create_spinner([
            "Available",
            "Used",
            "Faulty",
            "Reserved"
        ])

        self.notes = create_input(
            multiline=True,
            height=100
        )

        add_field(layout, "E1 Number", self.e1_number)
        add_field(layout, "Service Name", self.service)
        add_field(layout, "Destination", self.destination)
        add_field(layout, "Status", self.status)
        add_field(layout, "Notes", self.notes)

        save = Button(
            text="SAVE E1",
            size_hint_y=None,
            height=55
        )

        save.bind(
            on_release=self.save_e1
        )

        layout.add_widget(save)

        scroll.add_widget(layout)

        self.content = scroll

        self.load_e1()

    def load_e1(self):

        connection = get_connection()
        cursor = connection.cursor()

        cursor.execute("""
        SELECT
            e1_number,
            service_name,
            destination,
            status,
            notes
        FROM e1_links
        WHERE id=?
        """, (self.e1_id,))

        row = cursor.fetchone()

        connection.close()

        if row:

            self.e1_number.text = str(row[0])
            self.service.text = row[1] or ""
            self.destination.text = row[2] or ""
            self.status.text = row[3] or "Available"
            self.notes.text = row[4] or ""

    def save_e1(self, instance):

        connection = get_connection()
        cursor = connection.cursor()

        cursor.execute("""
        UPDATE e1_links SET
            e1_number=?,
            service_name=?,
            destination=?,
            status=?,
            notes=?
        WHERE id=?
        """, (
            self.e1_number.text.strip(),
            self.service.text.strip(),
            self.destination.text.strip(),
            self.status.text,
            self.notes.text.strip(),
            self.e1_id
        ))

        connection.commit()
        connection.close()

        self.dismiss()


# ============================================================
# POWER
# ============================================================

class PowerScreen(Screen):

    def on_pre_enter(self):
        self.build_screen()

    def build_screen(self):

        self.clear_widgets()

        scroll = ScrollView()

        form = BoxLayout(
            orientation="vertical",
            padding=20,
            spacing=10,
            size_hint_y=None
        )

        form.bind(
            minimum_height=
            form.setter("height")
        )

        title = Label(
            text="POWER SYSTEM",
            font_size=24,
            size_hint_y=None,
            height=65
        )

        self.electricity = create_spinner([
            "Generator",
            "Public Electricity",
            "Commercial Electricity",
            "Solar",
            "Hybrid"
        ])

        self.generator_type = create_input()
        self.generator_capacity = create_input()

        self.battery_type = create_spinner([
            "GEL",
            "Lithium",
            "Lead Acid",
            "Other"
        ])

        self.battery_number = create_input()
        self.battery_capacity = create_input()
        self.solar_system = create_input()

        self.notes = create_input(
            multiline=True,
            height=110
        )

        add_field(form, "Electricity Type", self.electricity)
        add_field(form, "Generator Type", self.generator_type)
        add_field(form, "Generator Capacity", self.generator_capacity)
        add_field(form, "Battery Type", self.battery_type)
        add_field(form, "Number of Batteries", self.battery_number)
        add_field(form, "Battery Capacity", self.battery_capacity)
        add_field(form, "Solar System", self.solar_system)
        add_field(form, "Notes", self.notes)

        save = Button(
            text="SAVE POWER SYSTEM",
            size_hint_y=None,
            height=60,
            font_size=18
        )

        back = Button(
            text="BACK TO SYSTEMS DASHBOARD",
            size_hint_y=None,
            height=60
        )

        save.bind(
            on_release=self.save_power
        )

        back.bind(
            on_release=lambda x:
            setattr(
                self.manager,
                "current",
                "systems"
            )
        )

        form.add_widget(title)
        form.add_widget(save)
        form.add_widget(back)

        scroll.add_widget(form)
        self.add_widget(scroll)

        self.load_power()

    def load_power(self):

        connection = get_connection()
        cursor = connection.cursor()

        cursor.execute(
            "SELECT * FROM power_system WHERE site_id=?",
            (CURRENT_SITE_ID,)
        )

        row = cursor.fetchone()

        connection.close()

        if not row:
            return

        self.electricity.text = row[2] or "Select"
        self.generator_type.text = row[3] or ""
        self.generator_capacity.text = row[4] or ""

        self.battery_type.text = row[5] or "Select"
        self.battery_number.text = row[6] or ""
        self.battery_capacity.text = row[7] or ""

        self.solar_system.text = row[8] or ""
        self.notes.text = row[9] or ""

    def save_power(self, instance):

        if CURRENT_SITE_ID is None:

            show_message(
                "Error",
                "No Site selected."
            )

            return

        connection = get_connection()
        cursor = connection.cursor()

        cursor.execute(
            "SELECT id FROM power_system WHERE site_id=?",
            (CURRENT_SITE_ID,)
        )

        existing = cursor.fetchone()

        values = (
            self.electricity.text,
            self.generator_type.text.strip(),
            self.generator_capacity.text.strip(),
            self.battery_type.text,
            self.battery_number.text.strip(),
            self.battery_capacity.text.strip(),
            self.solar_system.text.strip(),
            self.notes.text.strip()
        )

        try:

            if existing:

                cursor.execute("""
                UPDATE power_system SET
                    electricity_type=?,
                    generator_type=?,
                    generator_capacity=?,
                    battery_type=?,
                    battery_number=?,
                    battery_capacity=?,
                    solar_system=?,
                    notes=?
                WHERE site_id=?
                """,
                values + (CURRENT_SITE_ID,))

            else:

                cursor.execute("""
                INSERT INTO power_system (
                    site_id,
                    electricity_type,
                    generator_type,
                    generator_capacity,
                    battery_type,
                    battery_number,
                    battery_capacity,
                    solar_system,
                    notes
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (CURRENT_SITE_ID,) + values)

            connection.commit()

        except Exception as error:

            connection.rollback()
            connection.close()

            show_message(
                "Database Error",
                str(error)
            )

            return

        connection.close()

        show_message(
            "Success",
            "Power System saved successfully."
        )


# ============================================================
# APP
# ============================================================

class TelecomApp(App):

    def build(self):

        initialize_database()

        manager = ScreenManager()

        # ====================================================
        # HOME
        # ====================================================

        manager.add_widget(
            HomeScreen(name="home")
        )

        # ====================================================
        # SITES
        # ====================================================

        manager.add_widget(
            SitesScreen(name="sites")
        )

        manager.add_widget(
            SiteFormScreen(name="site_form")
        )

        manager.add_widget(
            SiteDetailsScreen(name="site_details")
        )

        # ====================================================
        # PROFESSIONAL DASHBOARD
        # ====================================================

        manager.add_widget(
            SystemsScreen(name="systems")
        )

        manager.add_widget(
            SiteAlarmsScreen(name="site_alarms")
        )

        # ====================================================
        # 2G
        # ====================================================

        manager.add_widget(
            TwoGListScreen(name="2g_list")
        )

        manager.add_widget(
            TwoGFormScreen(name="2g_form")
        )

        manager.add_widget(
            TwoGSectorsScreen(name="2g_sectors")
        )

        manager.add_widget(
            TwoGSectorFormScreen(
                name="2g_sector_form"
            )
        )

        # ====================================================
        # LTE
        # ====================================================

        manager.add_widget(
            LTEListScreen(name="lte_list")
        )

        manager.add_widget(
            LTEFormScreen(name="lte_form")
        )

        manager.add_widget(
            LTERRUListScreen(name="lte_rru_list")
        )

        manager.add_widget(
            LTERRUFormScreen(name="lte_rru_form")
        )

        manager.add_widget(
            LTECellsListScreen(name="lte_cells_list")
        )

        manager.add_widget(
            LTECellFormScreen(name="lte_cell_form")
        )

        # ====================================================
        # MICROWAVE
        # ====================================================

        manager.add_widget(
            MicrowaveListScreen(
                name="microwave_list"
            )
        )

        manager.add_widget(
            MicrowaveFormScreen(
                name="microwave_form"
            )
        )

        # ====================================================
        # E1
        # ====================================================

        manager.add_widget(
            E1ListScreen(
                name="e1_list"
            )
        )

        # ====================================================
        # POWER
        # ====================================================

        manager.add_widget(
            PowerScreen(
                name="power"
            )
        )

        return manager


# ============================================================
# RUN APPLICATION
# ============================================================

if __name__ == "__main__":

    TelecomApp().run()