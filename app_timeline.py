import sys
import csv
import os
import subprocess
import re
from datetime import datetime, timedelta
from openpyxl import Workbook

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QHBoxLayout,
    QLabel, QFrame, QTableWidget, QTableWidgetItem, QHeaderView,
    QMessageBox, QFileDialog, QCheckBox, QStyledItemDelegate, QDateEdit
)
from PyQt6.QtGui import QIcon, QTextCharFormat, QKeySequence
from PyQt6.QtCore import Qt, QDate

from excel_writer import create_schedule_sheet
from utils import (
    load_stylesheet, global_exception_hook,
    HoverButton, resource_path, GlobalWheelEventFilter, BlankDoubleSpinBox
)
from constants import LOGO_MAP, DEBUG_MODE

# --- Constants for Table Columns ---
COL_ACTIVITY = 0
COL_WEIGHT = 1
COL_START_ORIG = 2
COL_END_ORIG = 3
COL_START_REV = 4
COL_END_REV = 5
COL_START_ACT = 6
COL_END_ACT = 7

DATE_FORMAT = "yyyy-MM-dd"

CANONICAL_HEADERS = {
    COL_ACTIVITY:   ['activity', 'activity name', 'task', 'task name', 'name', 'description', 'desc'],
    COL_WEIGHT:     ['weight', 'wt', 'cost', 'amount', 'budget', 'value', 'price'],
    COL_START_ORIG: ['start', 'start date', 'original start', 'planned start', 'baseline start'],
    COL_END_ORIG:   ['end', 'end date', 'original end', 'planned end', 'baseline end', 'finish'],
    COL_START_REV:  ['revised start', 'rev start', 'start (revised)', 'current start'],
    COL_END_REV:    ['revised end', 'rev end', 'end (revised)', 'current end'],
    COL_START_ACT:  ['actual start', 'act start', 'start (actual)'],
    COL_END_ACT:    ['actual end', 'act end', 'end (actual)']
}

# ------------------------------------------------------------------------
# --- CUSTOM TABLE WIDGET: COMPACT PASTE & SMART DATE PARSING ---
# ------------------------------------------------------------------------
class PasteableTableWidget(QTableWidget):
    def keyPressEvent(self, event):
        """
        Handles:
        1. Ctrl+V (Paste)
        2. Delete / Backspace (Clear Cells)
        """
        if event.matches(QKeySequence.StandardKey.Paste):
            self.paste_from_clipboard()
            return

        if event.key() in (Qt.Key.Key_Delete, Qt.Key.Key_Backspace):
            self.clear_selected_cells()
            return

        super().keyPressEvent(event)

    def clear_selected_cells(self):
        """
        Manually clears data from selected cells.
        """
        for item in self.selectedItems():
            item.setText("")
            item.setData(Qt.ItemDataRole.DisplayRole, None)

    def paste_from_clipboard(self):
        clipboard = QApplication.clipboard()
        text = clipboard.text()

        if not text:
            return

        # 1. Split into raw rows
        raw_rows = text.strip('\n').split('\n')

        # 2. COMPACT THE DATA: Remove completely empty rows immediately.
        #    This fixes issues with Excel merged cells creating "ghost" rows.
        rows = [r for r in raw_rows if r.strip()]

        if not rows:
            return

        # Get the starting cell
        selected = self.selectedIndexes()
        if not selected:
            start_row = 0
            start_col = 0
        else:
            selected.sort(key=lambda x: (x.row(), x.column()))
            start_row = selected[0].row()
            start_col = selected[0].column()

        date_cols = [COL_START_ORIG, COL_END_ORIG,
                     COL_START_REV, COL_END_REV,
                     COL_START_ACT, COL_END_ACT]

        # --- PASS 1: INFER DATE FORMAT (DMY vs MDY) ---
        date_candidates = []
        for r_idx, row_text in enumerate(rows):
            columns = row_text.split('\t')
            # Check bounds just for scanning
            if start_row + r_idx >= self.rowCount(): break

            for c_idx, data in enumerate(columns):
                target_col = start_col + c_idx
                if target_col in date_cols and data.strip():
                    date_candidates.append(data.strip())

        is_dmy_preference = self.infer_date_order(date_candidates)

        # --- PASS 2: PASTE COMPACTED DATA ---
        for r_idx, row_text in enumerate(rows):
            columns = row_text.split('\t')
            target_row = start_row + r_idx

            # Stop if we run out of table rows
            if target_row >= self.rowCount():
                break

            for c_idx, data in enumerate(columns):
                target_col = start_col + c_idx

                # Stop if we run out of table columns
                if target_col >= self.columnCount():
                    break

                val = data.strip()

                # Handle specific empty cells in a valid row
                # (We want to clear the cell if the source is blank)
                if not val:
                    self.setItem(target_row, target_col, QTableWidgetItem(""))
                    continue

                # --- 1. WEIGHT COLUMN ---
                if target_col == COL_WEIGHT:
                    try:
                        clean_num = val.replace(',', '').replace('$', '').replace('£', '')
                        float_val = float(clean_num)

                        item = self.item(target_row, target_col)
                        if not item:
                            item = QTableWidgetItem()
                            self.setItem(target_row, target_col, item)

                        item.setData(Qt.ItemDataRole.DisplayRole, float_val)
                    except ValueError:
                        self.setItem(target_row, target_col, QTableWidgetItem(""))

                # --- 2. DATE COLUMNS ---
                elif target_col in date_cols:
                    formatted_date = self.parse_date_smart(val, is_dmy_preference)
                    if formatted_date:
                        self.setItem(target_row, target_col, QTableWidgetItem(formatted_date))
                    else:
                        self.setItem(target_row, target_col, QTableWidgetItem(""))

                # --- 3. TEXT COLUMNS ---
                else:
                    self.setItem(target_row, target_col, QTableWidgetItem(val))

    @staticmethod
    def clean_numeric_date_str(date_str):
        """
        Strips time and normalizes separators.
        Returns None if alphanumeric.
        """
        if re.search(r'[a-zA-Z]', date_str):
            return None
        base_date = date_str.split(' ')[0]
        normalized = re.sub(r'[.\-\\]', '/', base_date)
        return normalized

    def infer_date_order(self, date_strings):
        """
        Check batch for DMY vs MDY evidence.
        """
        has_dmy = False
        has_mdy = False

        for ds in date_strings:
            clean = self.clean_numeric_date_str(ds)
            if not clean: continue

            parts = clean.split('/')
            if len(parts) == 3:
                try:
                    p0 = int(parts[0])
                    p1 = int(parts[1])
                    if p0 > 1000: continue
                    if p0 > 12 >= p1: has_dmy = True
                    if p1 > 12 >= p0: has_mdy = True
                except ValueError:
                    continue

        if has_dmy and not has_mdy: return True
        if has_mdy and not has_dmy: return False
        return None

    def parse_date_smart(self, date_str, is_dmy_preference):
        if not date_str: return None

        # 1. ISO Match
        iso_match = re.search(r'(\d{4})-(\d{1,2})-(\d{1,2})', date_str)
        if iso_match:
            try:
                return f"{iso_match.group(1)}-{iso_match.group(2).zfill(2)}-{iso_match.group(3).zfill(2)}"
            except:
                pass

        # 2. Text Formats
        text_formats = [
            "%d-%b-%y", "%d-%b-%Y", "%d %b %Y", "%d %b %y",
            "%b %d, %Y", "%b %d %Y"
        ]
        for fmt in text_formats:
            try:
                return datetime.strptime(date_str, fmt).strftime("%Y-%m-%d")
            except ValueError:
                continue

        # 3. Numeric Formats
        clean = self.clean_numeric_date_str(date_str)
        if not clean: return None

        dmy = ["%d/%m/%Y", "%d/%m/%y"]
        mdy = ["%m/%d/%Y", "%m/%d/%y"]
        ymd = ["%Y/%m/%d"]

        if is_dmy_preference is True:
            candidates = dmy + mdy + ymd
        elif is_dmy_preference is False:
            candidates = mdy + dmy + ymd
        else:
            candidates = mdy + dmy + ymd

        for fmt in candidates:
            try:
                return datetime.strptime(clean, fmt).strftime("%Y-%m-%d")
            except ValueError:
                continue

        return None

# ------------------------------------------------------------------------
# --- DELEGATES & MAIN WINDOW ---
# ------------------------------------------------------------------------
class NumberDelegate(QStyledItemDelegate):
    def displayText(self, value, locale):
        """
        Format the number for display: 1,234.56
        """
        try:
            val = float(value)
            return f"{val:,.2f}"
        except (ValueError, TypeError):
            return str(value)

    def createEditor(self, parent, option, index):
        return BlankDoubleSpinBox(0, 9_999_999_999.9999, parent=parent)


class DateDelegate(QStyledItemDelegate):
    # Constants for formats
    ISO_FMT = "yyyy-MM-dd"
    DISP_FMT = "MMM d, yyyy"

    def displayText(self, value, locale):
        if not value: return ""
        try:
            dt = datetime.strptime(str(value), "%Y-%m-%d")
            return dt.strftime("%b %d, %Y")
        except ValueError:
            return str(value)

    def createEditor(self, parent, option, index):
        editor = QDateEdit(parent)
        editor.setDisplayFormat(self.DISP_FMT)
        editor.setCalendarPopup(True)
        editor.setFrame(False)

        calendar = editor.calendarWidget()
        fmt = QTextCharFormat()
        calendar.setWeekdayTextFormat(Qt.DayOfWeek.Saturday, fmt)
        calendar.setWeekdayTextFormat(Qt.DayOfWeek.Sunday, fmt)
        return editor

    def setEditorData(self, editor, index):
        text = index.model().data(index, Qt.ItemDataRole.EditRole)
        if text:
            try:
                qdate = QDate.fromString(str(text), self.ISO_FMT)
                editor.setDate(qdate)
            except ValueError:
                editor.setDate(QDate.currentDate())
        else:
            editor.setDate(QDate.currentDate())

    def setModelData(self, editor, model, index):
        date_str = editor.date().toString(self.ISO_FMT)
        model.setData(index, date_str, Qt.ItemDataRole.EditRole)


class TimelineWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Timeline & S-Curve Generator')
        self.setWindowIcon(QIcon(resource_path(LOGO_MAP['app_timeline'])))
        self.setGeometry(50, 50, 900, 500)

        # Main Layout
        main_widget = QFrame()
        main_widget.setObjectName('timelinePage')
        main_widget.setProperty('class', 'page')
        main_layout = QVBoxLayout(main_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        self.setCentralWidget(main_widget)

        # --- 1. Top Bar ---
        top_bar = QFrame()
        top_bar.setProperty('class', 'title-row')
        top_layout = QHBoxLayout(top_bar)
        top_layout.setContentsMargins(0, 0, 0, 0)
        top_layout.setSpacing(3)

        title = QLabel("Project Schedule")
        title.setProperty('class', 'header-1')

        top_layout.addWidget(title)
        top_layout.addStretch()
        main_layout.addWidget(top_bar)

        # --- 2. Configuration Panel ---
        config_panel = QFrame()
        config_panel.setProperty('class', 'panel')
        config_layout = QHBoxLayout(config_panel)
        config_layout.setContentsMargins(0, 0, 0, 0)
        config_layout.setSpacing(3)

        self.chk_scurve = QCheckBox("S-Curve")
        self.chk_scurve.setChecked(True)
        self.chk_scurve.setProperty('class', 'check-box')
        self.chk_scurve.toggled.connect(self.update_column_visibility)

        self.chk_rev = QCheckBox("Add Revised")
        self.chk_rev.setProperty('class', 'check-box')
        self.chk_rev.toggled.connect(self.update_column_visibility)

        self.chk_act = QCheckBox("Add Actual")
        self.chk_act.setProperty('class', 'check-box')
        self.chk_act.toggled.connect(self.update_column_visibility)

        self.btn_import = HoverButton('')
        self.btn_import.setProperty('class', 'blue-button import-button')
        self.btn_import.setToolTip('Import CSV')
        self.btn_import.clicked.connect(self.import_csv)

        self.btn_add = HoverButton("+")
        self.btn_add.setProperty('class', 'green-button add-button')
        self.btn_add.clicked.connect(self.add_row)

        self.btn_del = HoverButton("-")
        self.btn_del.setProperty('class', 'red-button remove-button')
        self.btn_del.clicked.connect(self.remove_row)

        config_layout.addWidget(self.chk_scurve)
        config_layout.addWidget(self.chk_rev)
        config_layout.addWidget(self.chk_act)
        config_layout.addStretch()
        config_layout.addWidget(self.btn_import)
        config_layout.addSpacing(3)
        config_layout.addWidget(self.btn_add)
        config_layout.addWidget(self.btn_del)

        main_layout.addWidget(config_panel)

        # --- 3. Spreadsheet Area (UPDATED CLASS) ---
        self.table = PasteableTableWidget()  # <--- Using Custom Class Here
        self.setup_table()

        self.table.setFrameShape(QFrame.Shape.NoFrame)
        self.table.setAlternatingRowColors(True)
        main_layout.addWidget(self.table)

        # --- 4. Bottom Controls ---
        bottom_bar = QFrame()
        bottom_bar.setProperty('class', 'bottom-nav')
        bottom_layout = QHBoxLayout(bottom_bar)
        bottom_layout.setContentsMargins(0, 0, 0, 0)
        bottom_layout.setSpacing(0)

        self.btn_export = HoverButton("Generate Excel")
        self.btn_export.setProperty('class', 'green-button next-button')
        self.btn_export.clicked.connect(self.generate_excel)

        bottom_layout.addStretch()
        bottom_layout.addWidget(self.btn_export)

        main_layout.addWidget(bottom_bar)

        # Initial State
        self.update_column_visibility()
        if DEBUG_MODE:
            self.prefill_data()
        else:
            self.add_row()

    def setup_table(self):
        cols = [
            "Activity Name", "Weight",
            "Start", "End",
            "Start (Revised)", "End (Revised)",
            "Start (Actual)", "End (Actual)"
        ]
        self.table.setColumnCount(len(cols))
        self.table.setHorizontalHeaderLabels(cols)

        # 1. Apply Date Delegate
        date_delegate = DateDelegate(self.table)
        for col_idx in [COL_START_ORIG, COL_END_ORIG, COL_START_REV, COL_END_REV, COL_START_ACT, COL_END_ACT]:
            self.table.setItemDelegateForColumn(col_idx, date_delegate)

        # 2. Apply Number Delegate to Weight Column
        number_delegate = NumberDelegate(self.table)
        self.table.setItemDelegateForColumn(COL_WEIGHT, number_delegate)

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setMinimumSectionSize(120)
        self.table.verticalHeader().setDefaultAlignment(Qt.AlignmentFlag.AlignCenter)

    def update_column_visibility(self):
        self.table.setColumnHidden(COL_WEIGHT, not self.chk_scurve.isChecked())
        is_rev = self.chk_rev.isChecked()
        self.table.setColumnHidden(COL_START_REV, not is_rev)
        self.table.setColumnHidden(COL_END_REV, not is_rev)
        is_act = self.chk_act.isChecked()
        self.table.setColumnHidden(COL_START_ACT, not is_act)
        self.table.setColumnHidden(COL_END_ACT, not is_act)

    def add_row(self):
        row_idx = self.table.rowCount()
        self.table.insertRow(row_idx)
        self.table.setItem(row_idx, COL_ACTIVITY, QTableWidgetItem(f"Task {row_idx + 1}"))

        item_wt = QTableWidgetItem()
        item_wt.setData(Qt.ItemDataRole.DisplayRole, 1.0)
        self.table.setItem(row_idx, COL_WEIGHT, item_wt)

        for c in range(2, 8):
            self.table.setItem(row_idx, c, QTableWidgetItem(""))

    def remove_row(self):
        rows = sorted(set(index.row() for index in self.table.selectedIndexes()))
        if not rows and self.table.rowCount() > 0:
            self.table.removeRow(self.table.rowCount() - 1)
        else:
            for row in reversed(rows):
                self.table.removeRow(row)

    def clear_table(self):
        self.table.setRowCount(0)
        self.add_row()

    def get_table_data(self):
        data = []
        rows = self.table.rowCount()
        for r in range(rows):
            try:
                activity_item = self.table.item(r, COL_ACTIVITY)
                if not activity_item: continue
                activity = activity_item.text()
                if not activity.strip(): continue

                if self.chk_scurve.isChecked():
                    weight_str = self.table.item(r, COL_WEIGHT).text()
                    try:
                        weight = float(weight_str.replace(',', ''))
                    except ValueError:
                        weight = 0.00
                else:
                    weight = 0.01

                def p_date(col):
                    item = self.table.item(r, col)
                    if not item or not item.text().strip(): return None
                    return datetime.strptime(item.text(), "%Y-%m-%d").date()

                row_data = {
                    'name': activity,
                    'weight': weight,
                    'orig': (p_date(COL_START_ORIG), p_date(COL_END_ORIG)),
                    'rev': (p_date(COL_START_REV), p_date(COL_END_REV)),
                    'act': (p_date(COL_START_ACT), p_date(COL_END_ACT)),
                }
                data.append(row_data)
            except ValueError:
                continue
        return data

    def import_csv(self):
        path, _ = QFileDialog.getOpenFileName(self, "Import CSV", "", "CSV Files (*.csv)")
        if not path:
            return

        try:
            with open(path, mode='r', newline='', encoding='utf-8-sig') as f:
                reader = csv.reader(f)
                all_rows = [row for row in reader if any(field.strip() for field in row)]

            if not all_rows:
                QMessageBox.warning(self, "Import Error", "The selected CSV file is empty.")
                return

            # Analyze headers and data
            mapping, start_row_index = self._detect_csv_layout(all_rows[0])

            # Extract data rows based on analysis
            data_rows = all_rows[start_row_index:]

            if not data_rows:
                QMessageBox.warning(self, "Import Warning", "Header row detected, but no data rows were found.")
                return

            # Confirm with user if overwriting existing data
            if self.table.rowCount() > 0:
                is_empty = (self.table.rowCount() == 1 and
                            self.table.item(0, 0) and
                            self.table.item(0, 0).text() == "Task 1")

                if not is_empty:
                    reply = QMessageBox.question(
                        self, "Overwrite Data?",
                        "Importing will overwrite the current table. Continue?",
                        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
                    )
                    if reply == QMessageBox.StandardButton.No:
                        return

            # --- AUTO-CONFIGURE UI BASED ON COLUMNS ---
            # We explicitly check/uncheck boxes based on what data is present.
            # Since these checkboxes have signals connected to update_column_visibility,
            # the table columns will hide/show automatically.

            detected_cols = set(mapping.values())

            # 1. Weight / S-Curve
            # If we have a weight column, check the box. If not, uncheck it.
            self.chk_scurve.setChecked(COL_WEIGHT in detected_cols)

            # 2. Revised Dates
            # If we have either start or end revised, enable the revised view
            has_revised = (COL_START_REV in detected_cols) or (COL_END_REV in detected_cols)
            self.chk_rev.setChecked(has_revised)

            # 3. Actual Dates
            # If we have either start or end actual, enable the actual view
            has_actual = (COL_START_ACT in detected_cols) or (COL_END_ACT in detected_cols)
            self.chk_act.setChecked(has_actual)

            # --- EXECUTE IMPORT ---
            self.table.setRowCount(0)  # Clear Table

            # Prepare date preference inference
            date_cols_indices = [
                idx for idx, tbl_col in mapping.items()
                if tbl_col in [COL_START_ORIG, COL_END_ORIG, COL_START_REV, COL_END_REV, COL_START_ACT, COL_END_ACT]
            ]

            sample_dates = []
            for row in data_rows[:10]:
                for csv_idx in date_cols_indices:
                    if csv_idx < len(row) and row[csv_idx].strip():
                        sample_dates.append(row[csv_idx])

            is_dmy = self.table.infer_date_order(sample_dates)

            self.table.setRowCount(len(data_rows))

            for r_idx, row_data in enumerate(data_rows):
                # Ensure Activity Name exists
                if COL_ACTIVITY not in mapping.values():
                    self.table.setItem(r_idx, COL_ACTIVITY, QTableWidgetItem(f"Activity {r_idx + 1}"))

                for csv_col_idx, table_col_idx in mapping.items():
                    if csv_col_idx >= len(row_data):
                        continue

                    val = row_data[csv_col_idx].strip()
                    if not val:
                        continue

                    # Weight
                    if table_col_idx == COL_WEIGHT:
                        try:
                            clean_num = val.replace(',', '').replace('$', '').replace('£', '')
                            float_val = float(clean_num)
                            item = QTableWidgetItem()
                            item.setData(Qt.ItemDataRole.DisplayRole, float_val)
                            self.table.setItem(r_idx, table_col_idx, item)
                        except ValueError:
                            pass

                    # Dates
                    elif table_col_idx in [COL_START_ORIG, COL_END_ORIG, COL_START_REV, COL_END_REV, COL_START_ACT,
                                           COL_END_ACT]:
                        formatted_date = self.table.parse_date_smart(val, is_dmy)
                        if formatted_date:
                            self.table.setItem(r_idx, table_col_idx, QTableWidgetItem(formatted_date))

                    # Text
                    else:
                        self.table.setItem(r_idx, table_col_idx, QTableWidgetItem(val))

            QMessageBox.information(self, "Success", f"Successfully imported {len(data_rows)} rows.")

        except Exception as e:
            QMessageBox.critical(self, "Import Error", f"Failed to import CSV.\n{e}")

    def _detect_csv_layout(self, header_row):
        """
        Returns:
            mapping (dict): {csv_column_index: table_column_constant}
            start_row_index (int): 0 if no header found (start immediately), 1 if header found.
        """
        mapping = {}
        matches_found = 0

        # Check against canonical headers
        for csv_idx, header_text in enumerate(header_row):
            clean_header = header_text.lower().strip()

            for table_col, candidates in CANONICAL_HEADERS.items():
                if clean_header in candidates:
                    mapping[csv_idx] = table_col
                    matches_found += 1
                    break

        # HEURISTIC:
        # If we matched at least 2 columns (e.g. "Activity" and "Start"),
        # we assume this is a valid header row.
        if matches_found >= 2:
            return mapping, 1

        # If headers weren't detected, assume standard column order matches the table
        # strict order: Name, Weight, Start, End...
        default_mapping = {}
        limit = min(len(header_row), self.table.columnCount())
        for i in range(limit):
            default_mapping[i] = i

        return default_mapping, 0

    def generate_excel(self):
        raw_data = self.get_table_data()
        if not raw_data:
            QMessageBox.warning(self, "No Data", "Table is empty or invalid.")
            return

        categories = ['orig']
        if self.chk_rev.isChecked(): categories.append('rev')
        if self.chk_act.isChecked(): categories.append('act')
        check_boxes = {'Actual': self.chk_act.isChecked(), 'Revised': self.chk_rev.isChecked(), 'S-Curve': self.chk_scurve.isChecked()}

        all_dates = []
        for row in raw_data:
            for cat in categories:
                start, end = row[cat]
                if start: all_dates.append(start)
                if end: all_dates.append(end)

        if not all_dates:
            QMessageBox.warning(self, "Date Error", "No valid dates found in the selected plans.")
            return

        global_min = min(all_dates)
        global_max = max(all_dates)
        table_start_date = datetime(global_min.year, global_min.month, 1)
        tmp_end_date = global_max + timedelta(days=31)
        table_end_date = datetime(tmp_end_date.year, tmp_end_date.month, 1) - timedelta(days=1)
        total_days = (table_end_date - table_start_date).days + 1

        wb = Workbook()
        default_ws = wb.active
        wb.remove(default_ws)

        ws = wb.create_sheet('Schedule')
        ws = create_schedule_sheet(ws, raw_data, check_boxes, table_start_date, total_days)

        save_path, _ = QFileDialog.getSaveFileName(
            self, 'Save Schedule', 'project_timeline.xlsx', 'Excel Files (*.xlsx)'
        )
        if save_path:
            try:
                wb.save(save_path)
                self.open_file(save_path)
            except Exception as e:
                QMessageBox.critical(self, "Save Error", f"Could not save file.\n{e}")

    def open_file(self, path):
        if sys.platform == 'win32':
            os.startfile(path)
        elif sys.platform == 'darwin':
            subprocess.call(['open', path])
        else:
            subprocess.call(['xdg-open', path])

    def prefill_data(self):
        data = [
            ("Site Clearing", 5000.0, "2023-11-01", "2023-11-05"),
            ("Excavation", 12000.55, "2023-11-04", "2023-11-10"),
        ]
        self.table.setRowCount(0)
        for name, wt, start, end in data:
            r = self.table.rowCount()
            self.table.insertRow(r)
            self.table.setItem(r, COL_ACTIVITY, QTableWidgetItem(name))
            wt_item = QTableWidgetItem()
            wt_item.setData(Qt.ItemDataRole.DisplayRole, wt)
            self.table.setItem(r, COL_WEIGHT, wt_item)
            self.table.setItem(r, COL_START_ORIG, QTableWidgetItem(start))
            self.table.setItem(r, COL_END_ORIG, QTableWidgetItem(end))
            rev_end = (datetime.strptime(end, "%Y-%m-%d") + timedelta(days=2)).strftime("%Y-%m-%d")
            self.table.setItem(r, COL_START_REV, QTableWidgetItem(start))
            self.table.setItem(r, COL_END_REV, QTableWidgetItem(rev_end))

    def keyPressEvent(self, event):
        if event.key() == Qt.Key.Key_Escape:
            self.setFocus()
        else:
            super().keyPressEvent(event)


if __name__ == '__main__':
    sys.excepthook = global_exception_hook
    app = QApplication(sys.argv)
    app.installEventFilter(GlobalWheelEventFilter())
    app.setStyleSheet(load_stylesheet('style.qss'))
    window = TimelineWindow()
    window.show()
    sys.exit(app.exec())