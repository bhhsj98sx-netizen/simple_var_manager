from PySide6.QtWidgets import QWidget, QLabel, QVBoxLayout
from PySide6.QtGui import QPixmap, QCursor
from PySide6.QtCore import Qt, Signal


class SceneCard(QWidget):
    clicked = Signal(str, str)            # scene_name, var_name
    selection_changed = Signal(str, bool) # kept for compatibility

    def __init__(self, scene_name: str, var_name: str, image_bytes: bytes | None):
        super().__init__()

        self.scene_name = scene_name
        self.var_name = var_name

        self._checked = False          # selected (selection mode)
        self._selection_mode = False
        self._active = False           # inspected (non-selection mode)

        self.setFixedWidth(220)
        self.setCursor(QCursor(Qt.PointingHandCursor))

        root = QVBoxLayout(self)
        root.setAlignment(Qt.AlignTop)
        root.setSpacing(6)

        # ------------------
        # Selection hint text
        # ------------------
        self.select_hint = QLabel("Click to select")
        self.select_hint.setVisible(False)
        self.select_hint.setStyleSheet("color: #bbb; font-size: 11px;")
        root.addWidget(self.select_hint)

        # ------------------
        # Image preview
        # ------------------
        self.image_label = QLabel()
        self.image_label.setFixedSize(200, 130)
        self.image_label.setAlignment(Qt.AlignCenter)
        self.image_label.setStyleSheet("border: 1px solid #444; background-color: #222;")

        if image_bytes:
            pixmap = QPixmap()
            pixmap.loadFromData(image_bytes)
            pixmap = pixmap.scaled(200, 130, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            self.image_label.setPixmap(pixmap)
        else:
            self.image_label.setText("No Preview")

        root.addWidget(self.image_label)

        # Scene label
        self.scene_label = QLabel(scene_name)
        self.scene_label.setWordWrap(True)
        self.scene_label.setStyleSheet("font-weight: bold;")
        root.addWidget(self.scene_label)

        # Var label
        self.var_label = QLabel(var_name)
        self.var_label.setWordWrap(True)
        self.var_label.setStyleSheet("color: #aaa; font-size: 10px;")
        root.addWidget(self.var_label)

        self.setLayout(root)
        self._apply_style()

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.clicked.emit(self.scene_name, self.var_name)
        super().mousePressEvent(event)

    # ------------------
    # API used by GUI
    # ------------------
    def set_selection_mode(self, enabled: bool):
        self._selection_mode = bool(enabled)
        self.select_hint.setVisible(self._selection_mode)

        if enabled:
            self._active = False  # clear active when entering selection mode

        self._apply_style()

    def set_checked(self, checked: bool):
        self._checked = bool(checked)
        self._apply_style()

    def is_checked(self) -> bool:
        return self._checked

    def set_active(self, active: bool):
        """
        Highlight card when inspecting (non-selection mode)
        """
        self._active = bool(active)
        self._apply_style()

    # ------------------
    # Styling logic
    # ------------------
    def _apply_style(self):
        # Priority:
        # 1. Selected (selection mode)
        # 2. Active (inspect mode)
        # 3. Normal / hover

        if self._selection_mode and self._checked:
            self.setStyleSheet("""
                QWidget {
                    border: 2px solid #3daee9;
                    border-radius: 8px;
                    background: rgba(61, 174, 233, 0.10);
                    padding: 4px;
                }
            """)
        elif (not self._selection_mode) and self._active:
            self.setStyleSheet("""
                QWidget {
                    border: 2px solid #f1c40f;
                    border-radius: 8px;
                    background: rgba(241, 196, 15, 0.10);
                    padding: 4px;
                }
            """)
        else:
            self.setStyleSheet("""
                QWidget {
                    border: 1px solid transparent;
                    border-radius: 8px;
                    background: transparent;
                    padding: 4px;
                }
                QWidget:hover {
                    border: 1px solid #666;
                    background: rgba(255,255,255,0.03);
                }
            """)
