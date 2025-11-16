from typing import Any

from PyQt6.QtSvg import QSvgRenderer
from PyQt6.QtWidgets import (
    QLabel, QApplication, QMessageBox, QWidget, QVBoxLayout, QSpinBox, QDoubleSpinBox, QPushButton, QGroupBox,
    QLineEdit, QComboBox, QTextEdit, QCheckBox, QScrollArea, QStackedWidget
)
import traceback
from PyQt6.QtGui import QPixmap, QCursor, QEnterEvent, QPainter, QColor
from PyQt6.QtCore import Qt, QEvent, pyqtSignal, QObject, QPropertyAnimation, QEasingCurve, QPoint, \
    QParallelAnimationGroup
import re
from typing import Literal
import os
import sys


class GlobalWheelEventFilter(QObject):
    """
    An event filter that intercepts and ignores wheel events for specific widgets
    to prevent accidental value changes when scrolling.
    """

    def eventFilter(self, obj, event):
        # Check if the event is a wheel event
        if event.type() == QEvent.Type.Wheel:
            # Check if the widget is one of the types we want to ignore scrolling on
            if isinstance(obj, (QComboBox, QSpinBox, QDoubleSpinBox)):
                # Return True to indicate the event has been handled and should be ignored.
                return True

        # For all other objects and events, pass them to the default implementation.
        return super().eventFilter(obj, event)

class AnimatedStackedWidget(QStackedWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._animation_duration = 500  # Animation duration in milliseconds
        self._easing_curve = QEasingCurve.Type.OutCubic
        self._next_widget_index = -1
        self._current_widget_index = 0
        self._p_now = QPoint(0, 0)
        self._p_next = QPoint(0, 0)
        self.animation_group = QParallelAnimationGroup(self)
        self.animation_group.finished.connect(self._animation_finished)

    def set_animation_duration(self, duration):
        self._animation_duration = duration

    def set_easing_curve(self, curve):
        self._easing_curve = curve

    def setCurrentIndex(self, index):
        if self.currentIndex() == index:
            return

        self._current_widget_index = self.currentIndex()
        self._next_widget_index = index

        current_widget = self.widget(self._current_widget_index)
        next_widget = self.widget(self._next_widget_index)

        if not current_widget or not next_widget:
            super().setCurrentIndex(index)
            return

        offset_x = self.frameRect().width()
        if self._current_widget_index < self._next_widget_index:
            # Slide from right to left
            self._p_now = QPoint(0, 0)
            self._p_next = QPoint(offset_x, 0)
            next_widget.setGeometry(offset_x, 0, self.width(), self.height())
        else:
            # Slide from left to right
            self._p_now = QPoint(0, 0)
            self._p_next = QPoint(-offset_x, 0)
            next_widget.setGeometry(-offset_x, 0, self.width(), self.height())

        next_widget.show()
        next_widget.raise_()

        anim_now = QPropertyAnimation(current_widget, b"pos")
        anim_now.setStartValue(self._p_now)
        anim_now.setEndValue(-self._p_next)
        anim_now.setDuration(self._animation_duration)
        anim_now.setEasingCurve(self._easing_curve)

        anim_next = QPropertyAnimation(next_widget, b"pos")
        anim_next.setStartValue(self._p_next)
        anim_next.setEndValue(self._p_now)
        anim_next.setDuration(self._animation_duration)
        anim_next.setEasingCurve(self._easing_curve)

        self.animation_group.clear()
        self.animation_group.addAnimation(anim_now)
        self.animation_group.addAnimation(anim_next)
        self.animation_group.start()

    def _animation_finished(self):
        super().setCurrentIndex(self._next_widget_index)
        current_widget = self.widget(self._current_widget_index)
        if current_widget:
            current_widget.hide()
            current_widget.move(self._p_now)

class MemoryGroupBox(QGroupBox):
    def __init__(self, title='', parent=None):
        super().__init__(title, parent)
        self.setCheckable(True)
        self._cache = {}
        self.toggled.connect(self.on_toggled)
        self.toggled.connect(self.update_group_box_style)

    def on_toggled(self, checked):
        if checked:
            self.restore_children()
        else:
            self.save_children()
            self.reset_children()

    def save_children(self):
        self._cache.clear()
        for w in self.findChildren(QWidget):
            if isinstance(w, QLineEdit):
                self._cache[w] = w.text()
            elif isinstance(w, QComboBox):
                self._cache[w] = w.currentIndex()
            elif isinstance(w, QSpinBox) or isinstance(w, QDoubleSpinBox):
                self._cache[w] = w.value()
            elif isinstance(w, QTextEdit):
                self._cache[w] = w.toPlainText()
            elif isinstance(w, QCheckBox):
                self._cache[w] = w.isChecked()

    def restore_children(self):
        for w, val in self._cache.items():
            if isinstance(w, QLineEdit):
                w.setText(val)
            elif isinstance(w, QComboBox):
                w.setCurrentIndex(val)
            elif isinstance(w, QSpinBox) or isinstance(w, QDoubleSpinBox):
                w.setValue(val)
            elif isinstance(w, QTextEdit):
                w.setPlainText(val)
            elif isinstance(w, QCheckBox):
                w.setChecked(val)

    def reset_children(self):
        for w in self.findChildren(QWidget):
            if isinstance(w, (QLineEdit, QTextEdit)):
                w.clear()
            elif isinstance(w, QComboBox):
                w.setCurrentIndex(-1)
            elif isinstance(w, (QSpinBox, QDoubleSpinBox)):
                w.setValue(w.minimum())
                w.setSuffix('')
            elif isinstance(w, QCheckBox):
                w.setChecked(True)

    def update_group_box_style(self) -> None:
        """Updates the dynamic 'class' property and forces a style refresh."""
        # Set a dynamic property for styling
        is_checked = self.isChecked()
        self.setProperty('checkedState', 'checked' if is_checked else 'unchecked')

        # Re-polish the widget to apply the new style
        self.style().polish(self)
        # Also re-polish children to update their styles (like labels)
        for child in self.findChildren(QWidget):
            assert isinstance(child, QWidget)
            child.style().polish(child)


class LinkSpinboxes(QCheckBox):
    """
    A QCheckBox that locks the value of one QSpinBox or QDoubleSpinBox to another.

    When checked, this checkbox disables the `copy_to_spinbox` and ensures its
    value always matches the `copy_from_spinbox`. When unchecked, it re-enables
    the `copy_to_spinbox`, allowing its value to be changed independently.
    """

    def __init__(
            self,
            copy_from_spinbox: QSpinBox | QDoubleSpinBox,
            copy_to_spinbox: QSpinBox | QDoubleSpinBox,
            tooltip: str | None = None,
            parent=None
    ):
        """
        Initializes the LockRatioCheckBox.

        Args:
            copy_from_spinbox: The spinbox to copy the value from.
            copy_to_spinbox: The spinbox to copy the value to.
            tooltip: An optional tooltip string for the checkbox.
            parent: An optional parent widget.
        """
        super().__init__(parent)

        self.copy_from_spinbox = copy_from_spinbox
        self.copy_to_spinbox = copy_to_spinbox

        self.setProperty('class', 'lock-ratio')
        self.setChecked(True)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        if tooltip:
            self.setToolTip(tooltip)

        # Connect signals
        self.toggled.connect(self._on_toggled)
        self.copy_from_spinbox.valueChanged.connect(self._on_source_value_changed)

        self.copy_to_spinbox.setEnabled(False)
        self.copy_to_spinbox.setValue(self.copy_from_spinbox.value())

    def _on_toggled(self, checked: bool):
        """
        Handles the toggled signal of the checkbox.

        Enables or disables the target spinbox based on the checked state
        and syncs the value if checked.
        """
        self.copy_to_spinbox.setEnabled(not checked)
        if checked:
            self.copy_to_spinbox.setValue(self.copy_from_spinbox.value())

    def _on_source_value_changed(self, value: int | float):
        """
        Handles the valueChanged signal of the source spinbox.

        Updates the target spinbox's value if the checkbox is checked.
        """
        if self.isChecked():
            self.copy_to_spinbox.setValue(value)

class BlankSpinBox(QSpinBox):
    """Integer spinbox with blank special value."""
    def __init__(self, minimum: int, maximum: int, initial: int | None = None, suffix: str | None = None, parent: QWidget | None = None):
        super().__init__(parent)
        self.setRange(minimum, maximum)
        if initial:
            self.setValue(initial)
        else:
            self.setSpecialValueText(' ')
        if suffix:
            self.setSuffix(suffix)
        self.setGroupSeparatorShown(True)

class BlankDoubleSpinBox(QDoubleSpinBox):
    """Float spinbox with configurable decimals, and blank special value."""
    def __init__(
        self,
        minimum: float,
        maximum: float,
        decimals: int = 2,
        suffix: str | None = None,
        parent: QWidget | None = None
    ):
        super().__init__(parent)
        self.setRange(minimum, maximum)
        self.setDecimals(decimals)
        self.setSpecialValueText(' ')
        if suffix:
            self.setSuffix(suffix)

class InfoPopup(QWidget):
    """
    A simple, frameless popup widget to display informational text.
    Styled via QSS with the object name 'infoPopup'.
    """

    def __init__(self, parent: QWidget | None = None) -> None:
        """
        Initializes a simple, frameless popup widget for displaying information.

        Args:
            parent: The parent widget, if any.
        """
        super().__init__(parent)
        # Use ToolTip flag to make it float on top and not appear in the taskbar
        self.setWindowFlags(Qt.WindowType.ToolTip | Qt.WindowType.FramelessWindowHint)
        self.setObjectName('infoPopup')  # For styling

        self.layout = QVBoxLayout(self)
        self.label = QLabel(self)
        self.label.setWordWrap(True)  # Ensure text wraps if it's long
        self.layout.addWidget(self.label)

    def set_info_text(self, text: str) -> None:
        """
        Sets the text content of the popup's label.

        Args:
            text: The string to display, which can include rich text formatting.
        """
        self.label.setText(text)


class HoverLabel(QLabel):
    """
    A QLabel subclass that emits signals on mouse enter and leave events.
    """
    # noinspection PyUnresolvedReferences
    mouseEntered = pyqtSignal()
    # noinspection PyUnresolvedReferences
    mouseLeft = pyqtSignal()

    def enterEvent(self, event: QEnterEvent) -> None:  # <-- Corrected type hint
        """
        Emits the mouseEntered signal when the cursor enters the label's area.

        Args:
            event: The enter event, specifically a QEnterEvent.
        """
        # noinspection PyUnresolvedReferences
        self.mouseEntered.emit()
        super().enterEvent(event)

    def leaveEvent(self, event: QEvent) -> None:  # <-- This hint is already correct
        """
        Emits the mouseLeft signal when the cursor leaves the label's area.

        Args:
            event: The leave event.
        """
        # noinspection PyUnresolvedReferences
        self.mouseLeft.emit()
        super().leaveEvent(event)

class HoverButton(QPushButton):
    """
    A custom QPushButton that automatically sets a pointing hand cursor
    on hover.
    """
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))

def load_stylesheet(filename):
    """
    Loads a QSS file, replaces relative image paths with absolute paths,
    and returns the processed content. This is crucial for PyInstaller.
    """
    # First, get the absolute path to the QSS file itself
    qss_file_path = resource_path(filename)

    # Read the content of the QSS file
    with open(qss_file_path, 'r') as f:
        stylesheet = f.read()

    # Get the correct absolute path to the 'images' directory,
    # ensuring it uses forward slashes, which CSS/QSS requires.
    image_dir_path = resource_path('images').replace('\\', '/')

    # Replace the relative path prefix with the absolute one
    # This finds 'url(images/' and replaces it with 'url(C:/.../_MEIxxxx/images/'
    processed_stylesheet = stylesheet.replace('url(images/', f'url({image_dir_path}/')

    return processed_stylesheet

def get_img(path: str, width: int, height: int, class_name: str = 'image-container',
            alignment: Qt.AlignmentFlag = None, return_pixmap: bool = False) -> QLabel | QPixmap:
    """
    Creates a QLabel with a scaled pixmap from an image file.

    Args:
        path: The file path of the image.
        width: The target width for scaling.
        height: The target height for scaling.
        class_name: The CSS class name to set on the QLabel.
        alignment: The alignment of the pixmap within the QLabel.
        return_pixmap: If True, returns the QPixmap object instead of the QLabel.

    Returns:
        A configured QLabel or the raw QPixmap.
    """
    container = QLabel()
    container.setProperty('class', class_name)
    if alignment is not None:
        container.setAlignment(alignment)
    pixmap = QPixmap(path)
    if pixmap.isNull():
        if return_pixmap:
            return pixmap
        container.setText(f'Image {path} not found or could not be loaded.')
    else:
        # Scaling can also theoretically fail, e.g., with a MemoryError on huge images
        try:
            scaled_pixmap = pixmap.scaled(width, height, Qt.AspectRatioMode.KeepAspectRatio,
                                          Qt.TransformationMode.SmoothTransformation)
            if return_pixmap:
                return scaled_pixmap
            container.setPixmap(scaled_pixmap)
        except MemoryError:
            container.setText(f'Not enough memory to scale the image {path}.')
    return container

def update_image(selected_text: str, image_map: dict[str, str], update_this_object: QLabel, width: int | None = None,
                 fallback: str = None) -> None:
    """
    Updates the pixmap of a QLabel based on a text key.

    Args:
        selected_text: The key to look up in the image_map.
        image_map: A dictionary mapping text keys to image file paths.
        update_this_object: The QLabel widget whose pixmap will be updated.
        width: The target width for the new image. If None, uses the current pixmap's width.
        fallback: The image to serve if the index is wrong.
    """
    if fallback is None:
        fallback = resource_path('images/logo.png')
    if width is None:
        width = update_this_object.pixmap().width()
    path = image_map.get(selected_text, fallback)
    pixmap = get_img(path, width, width, return_pixmap=True)
    if pixmap.isNull():
        update_this_object.setText(f'Image path {path} not found.')
    else:
        update_this_object.setPixmap(get_img(path, width, width, return_pixmap=True))

def toggle_obj_visibility(selected_text: str, target_text: str, objs: QWidget | list[QWidget], hide_when_target: bool = False) -> None:
    """
    Toggles the visibility of one or more widgets based on a text comparison.

    Args:
        selected_text: The text from the source widget (e.g., QComboBox).
        target_text: The text to check for within the selected_text.
        objs: A single widget or a list of widgets to toggle.
        hide_when_target: If True, hides the object(s) on a match; otherwise, shows them.
    """
    def toggle(item):
        if hide_when_target:
            if target_text in selected_text:
                item.setVisible(False)
            else:
                item.setVisible(True)
        else:
            if target_text in selected_text:
                item.setVisible(True)
            else:
                item.setVisible(False)

    if isinstance(objs, list):
        for obj in objs:
            toggle(obj)
    else:
        toggle(objs)

def is_widget_empty(widget: QWidget) -> bool:
    """
    Checks if a given input widget is effectively empty.

    Args:
        widget: The widget to check.

    Returns:
        True if the widget is considered empty, False otherwise.
    """
    if isinstance(widget, (QSpinBox, QDoubleSpinBox)) and hasattr(widget, 'specialValueText'):
        # This is our custom BlankSpinBox or BlankDoubleSpinBox
        return widget.text() == widget.specialValueText()
    elif isinstance(widget, QLineEdit):
        return not widget.text().strip()
    elif isinstance(widget, QTextEdit):
        return not widget.toPlainText().strip()
    # For other widgets like QComboBox, we assume a selection is always valid if visible.
    return False


def parse_spacing_string(text: str) -> list[tuple[float]]:
    """
    Parses a spacing string into a list of tuples, enforcing strict validation.

    The string is a comma-separated list of entries, where each entry
    is in the format 'value@spacing' or 'value at spacing'.
    - 'value' can be an integer or 'rest' (case-insensitive).
    - 'spacing' must be an integer, optionally followed by 'mm' (case-insensitive).
    - The delimiter '@' or 'at' (case-insensitive) is required.

    Raises:
        TypeError: If the input is not a string.
        ValueError: If any entry is malformed or contains invalid units.

    For example:
    '1 @ 50mm , 5 at 100, rest AT 200 mm'
    becomes:
    [(1, 50), (5, 100), ('rest', 200)]
    """
    if not isinstance(text, str):
        raise TypeError('Input must be a string.')

    results = []
    entries = re.split(r'\s*,\s*', text.strip())

    for entry in entries:
        if not entry:
            continue

        parts = re.split(r'\s*(?:@|at)\s*', entry, flags=re.IGNORECASE)
        if len(parts) != 2 or not parts[0] or not parts[1]:
            raise ValueError(
                f'Invalid format in {entry}. Each part must use @ or at to separate a value and a number.')

        value_str, spacing_str = parts
        spacing_str = spacing_str.strip()

        # Check for and remove the 'mm' unit, case-insensitively.
        if spacing_str.lower().endswith('mm'):
            spacing_str = spacing_str[:-2].strip()

        try:
            spacing = safe_parse_to_num(spacing_str)
        except ValueError:
            raise ValueError(
                f'Invalid spacing {parts[1]}. Spacing must be a whole number, optionally followed by mm.') from None

        if value_str.lower() == 'rest':
            value = 'rest'
        else:
            try:
                value = safe_parse_to_num(value_str)
            except ValueError:
                raise ValueError(
                    f'Invalid value {value_str} in {entry}. Value must be a whole number or rest.') from None
            if isinstance(value, float):
                raise ValueError(f'Invalid value {value_str} in {entry}. Value must be a whole number or rest.')

        results.append((value, spacing))

    return results

def style_invalid_input(widget: QWidget, is_valid: bool) -> None:
    """
    Applies or removes a CSS class to indicate invalid input for QLineEdit, QSpinBox, etc.

    Args:
        widget: The widget to style (e.g., QSpinBox).
        is_valid: True to remove invalid style, False to apply it.
    """
    if not hasattr(widget, 'property') or not hasattr(widget, 'setProperty'):
        return  # Not a widget we can style this way

    current_property = widget.property('class') or ''
    if is_valid:
        # Remove any invalid styling
        new_class = current_property.replace('invalid-input', '').strip()
    else:
        # Apply invalid styling if it's not already there
        if 'invalid-input' not in current_property:
            new_class = (current_property + ' invalid-input').strip()
        else:
            new_class = current_property

    if new_class != current_property:
        widget.setProperty('class', new_class)
        widget.style().polish(widget)

def parse_nested_dict(data: dict[str, Any]) -> dict[str, Any]:
    """
    Recursively traverses a dictionary, parsing widget text values into numbers.

    Args:
        data: The dictionary containing Qt widgets or other data.

    Returns:
        A new dictionary with widget values replaced by parsed data.
    """

    def recurse(obj):
        if isinstance(obj, dict):
            return {k: recurse(v) for k, v in obj.items()}
        elif isinstance(obj, list):
            return [recurse(v) for v in obj]
        elif isinstance(obj, QLineEdit):
            text = obj.text()
            try:
                return safe_parse_to_num(text)
            except ValueError:
                return text
        elif isinstance(obj, QComboBox):
            text = obj.currentText()
            try:
                return safe_parse_to_num(text)
            except ValueError:
                return text
        elif isinstance(obj, QTextEdit):
            return obj.toPlainText()
        elif isinstance(obj, (QSpinBox, QDoubleSpinBox)):
            return obj.value()
        else:
            return obj

    return recurse(data)

def get_bar_dia(code: int | str, system: Literal['ph', 'soft_metric', 'imperial'] = 'ph') -> float:
    if isinstance(code, str) and code.startswith('#'):
        code = int(code[1:])

    if system == 'imperial':
        bar_sizes = {
            3: 9.525,
            4: 12.7,
            5: 15.875,
            6: 19.05,
            7: 22.225,
            8: 25.4,
            9: 28.65,
            10: 32.26,
            11: 35.81,
            14: 43,
            18: 57.33
        }
    elif system == 'soft_metric':
        bar_sizes = {
            10: 9.525,
            13: 12.7,
            16: 15.875,
            19: 19.05,
            22: 22.225,
            25: 25.4,
            29: 28.65,
            32: 32.26,
            36: 35.81,
            43: 43,
            57: 57.33
        }
    elif system == 'ph':  # PH Standard Bar Naming
        bar_sizes = {
            10: 9.525,
            12: 12.7,
            16: 15.875,
            20: 19.05,
            25: 25.4,
            28: 28.65,
            32: 32.26,
            36: 35.81,
            40: 38.7,
            50: 50.8
        }
    else:
        raise ValueError(
            f"Invalid bar sizing scheme {system}. Valid choices are 'ph', 'soft_metric', and 'imperial'.")
    return bar_sizes[code]

def get_dia_code(mm: int | float, system: Literal['ph', 'soft_metric', 'imperial'] = 'ph') -> str:
    if system == 'imperial':
        rebar_code = {9.525: '#3', 12.7: '#4', 15.875: '#5', 19.05: '#6', 22.225: '#7',
                     25.4: '#8', 28.65: '#9', 32.26: '#10', 35.81: '#11', 43: '#14', 57.33: '#18'}
    elif system == 'soft_metric':
        rebar_code = {9.525: '#10', 12.7: '#13', 15.875: '#16', 19.05: '#19', 22.225: '#22',
                     25.4: '#25', 28.65: '#29', 32.26: '#32', 35.81: '#36', 43: '#43', 57.33: '#57'}
    elif system == 'ph':  # PH Standard Bar Naming
        rebar_code = {9.525: '#10', 12.7: '#12', 15.875: '#16', 19.05: '#20', 25.4: '#25',
                     28.65: '#28', 32.26: '#32', 35.81: '#36', 38.7: '#40', 50.8: '#50'}
    else:
        raise ValueError(
            f"Invalid bar sizing scheme {system}. Valid choices are 'ph', 'soft_metric', and 'imperial'.")

    # tolerance-based lookup
    for dia, code in rebar_code.items():
        if abs(mm - dia) < 0.2:  # tolerance 0.2 mm
            return code

    raise KeyError(f'No bar code found for {mm} mm in {system}')

def safe_parse_to_num(text: str) -> float:
    """
    Parses a string into a float or int, safely handling commas.
    Raises ValueError if parsing is not possible.
    """
    if not isinstance(text, str) or not text.strip():
        raise ValueError('Input cannot be empty.')

    # Remove commas and strip whitespace
    cleaned_text = text.replace(',', '').strip()

    if '.' in cleaned_text:
        try:
            num_float = float(cleaned_text)
            if num_float.is_integer():
                return int(num_float)
            return num_float
        except ValueError:
            # Re-raise with a more specific message if needed, or let the caller handle it.
            raise ValueError(f'Could not convert {text} to float.')
    try:
        return int(cleaned_text)
    except ValueError:
        # Re-raise with a more specific message if needed, or let the caller handle it.
        raise ValueError(f'Could not convert {text} to integer.')

def global_exception_hook(exc_type, exc_value, exc_traceback):
    """
    Catches any unhandled exceptions, extracts detailed location info,
    displays them in a dialog, and prints them to the console.
    """
    # Log the full error to the console (standard behavior)
    print('--- Unhandled Exception Caught ---')
    traceback.print_exception(exc_type, exc_value, exc_traceback)
    print('---------------------------------')

    # The last entry in the traceback is the actual line where the error happened.
    last_frame = traceback.extract_tb(exc_traceback)[-1]

    # Use os.path.basename to get just the filename, not the full path.
    file_name = os.path.basename(last_frame.filename)
    line_number = last_frame.lineno
    function_name = last_frame.name

    # Create a user-friendly and informative error message
    error_message = (
        f'An unexpected error occurred, and the application may need to close.\n\n'
        f'<b>File:</b> {file_name}<br>'
        f'<b>Function:</b> {function_name}<br>'
        f'<b>Line:</b> {line_number}<br><br>'
        f'<b>Error Type:</b> {exc_type.__name__}<br>'
        f'<b>Error Message:</b> {exc_value}'
    )

    # Show the error in a message box
    error_app = QApplication.instance() or QApplication([])
    error_dialog = QMessageBox()
    error_dialog.setIcon(QMessageBox.Icon.Critical)
    error_dialog.setText('Application Error')

    # Use setTextFormat to allow rich text (for bold tags)
    error_dialog.setTextFormat(Qt.TextFormat.RichText)
    error_dialog.setInformativeText(error_message)

    error_dialog.setWindowTitle('Unexpected Error')
    error_dialog.setStandardButtons(QMessageBox.StandardButton.Ok)
    error_dialog.exec()

    if not QApplication.instance():
        error_app.quit()

def resource_path(relative_path: str) -> str:
    """
    Get the absolute path to a resource, working for both development and PyInstaller.
    """
    # Check if the PyInstaller attribute exists
    if hasattr(sys, '_MEIPASS'):
        # Running in a bundled PyInstaller app
        base_path = getattr(sys, '_MEIPASS')
    else:
        # Running in a normal development environment
        base_path = os.path.abspath('.')

    return os.path.join(base_path, relative_path)

def make_scrollable(widget: QWidget, always_on: bool = False) -> QScrollArea:
    """
    Wraps a widget in a QScrollArea.

    Args:
        widget: The widget to make scrollable.
        always_on: If True, the vertical scrollbar is always visible.

    Returns:
        The configured QScrollArea containing the widget.
    """
    scroll = QScrollArea()
    scroll.setProperty('class', 'scroll-bar')
    scroll.setWidget(widget)
    scroll.setWidgetResizable(True)
    scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
    if always_on:
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
    else:
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
    return scroll

def svg_to_pixmap(svg_filename: str, width: int, height: int, color: QColor) -> QPixmap:
    renderer = QSvgRenderer(svg_filename)
    pixmap = QPixmap(width, height)
    pixmap.fill(Qt.GlobalColor.transparent)
    painter = QPainter(pixmap)
    renderer.render(painter) # this is the destination, and only its alpha is used!
    painter.setCompositionMode(
        painter.CompositionMode.CompositionMode_SourceIn)
    painter.fillRect(pixmap.rect(), color)
    painter.end()
    return pixmap
