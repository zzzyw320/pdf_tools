from PySide6.QtWidgets import (
    QDialog,
    QLabel,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QComboBox,
)
from PySide6.QtCore import Qt

from translations import TEXTS


class LanguageDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.selected_lang = None
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle(TEXTS["zh"]["lang_window_title"])
        self.setFixedSize(360, 180)

        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(18)

        self.label = QLabel(TEXTS["zh"]["lang_label"])
        self.label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)

        self.combo = QComboBox()
        self.combo.addItem(TEXTS["zh"]["lang_zh"], "zh")
        self.combo.addItem(TEXTS["zh"]["lang_en"], "en")

        button_layout = QHBoxLayout()
        button_layout.setSpacing(16)

        self.confirm_button = QPushButton(TEXTS["zh"]["lang_confirm"])
        self.confirm_button.clicked.connect(self.confirm_language)

        self.cancel_button = QPushButton(TEXTS["zh"]["lang_cancel"])
        self.cancel_button.clicked.connect(self.reject)

        button_layout.addWidget(self.confirm_button)
        button_layout.addWidget(self.cancel_button)

        layout.addWidget(self.label)
        layout.addWidget(self.combo)
        layout.addStretch()
        layout.addLayout(button_layout)

        self.setLayout(layout)

        self.setStyleSheet("""
            QWidget {
                font-size: 16px;
                color: black;
                background-color: white;
            }
            QLabel {
                color: black;
                background-color: white;
            }
            QComboBox {
                background-color: white;
                color: black;
                border: 1px solid #cfcfcf;
                border-radius: 6px;
                padding: 6px;
                min-height: 32px;
            }
            QPushButton {
                background-color: #e8f0fe;
                color: black;
                border: 1px solid #b7c9e2;
                border-radius: 8px;
                padding: 8px;
                min-height: 36px;
            }
            QPushButton:hover {
                background-color: #dbe9ff;
            }
        """)

    def confirm_language(self):
        self.selected_lang = self.combo.currentData()
        self.accept()