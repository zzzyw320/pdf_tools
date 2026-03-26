import os
import fitz

from PySide6.QtWidgets import (
    QWidget,
    QLabel,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QRadioButton,
    QButtonGroup,
    QLineEdit,
    QFileDialog,
    QMessageBox,
    QScrollArea,
    QFrame,
    QGridLayout,
    QSizePolicy,
    QProgressDialog,
)
from PySide6.QtCore import Qt, QCoreApplication
from PySide6.QtGui import QIntValidator, QPixmap, QImage

from pypdf import PdfReader, PdfWriter
from translations import TEXTS


class SplitWindow(QWidget):
    def __init__(self, pdf_path, lang="zh"):
        super().__init__()
        self.lang = lang
        self.texts = TEXTS[self.lang]

        self.pdf_path = pdf_path
        self.pdf_name = os.path.basename(pdf_path)
        self.pdf_base_name = os.path.splitext(self.pdf_name)[0]
        self.page_count = self.get_pdf_page_count(pdf_path)

        self.split_points = set()

        self.setWindowTitle(self.texts["split_window_title"])
        self.resize(980, 760)

        self.init_ui()

    def get_pdf_page_count(self, pdf_path):
        try:
            reader = PdfReader(pdf_path)
            return len(reader.pages)
        except Exception as e:
            QMessageBox.critical(
                self,
                self.texts["error"],
                f"{self.texts['read_pdf_failed']}\n{e}"
            )
            return 0

    def render_pdf_page_thumbnail(self, page_index, zoom=0.35):
        try:
            doc = fitz.open(self.pdf_path)
            page = doc.load_page(page_index)

            matrix = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=matrix, alpha=False)

            image_format = QImage.Format_RGB888
            qimage = QImage(
                pix.samples,
                pix.width,
                pix.height,
                pix.stride,
                image_format
            )

            pixmap = QPixmap.fromImage(qimage.copy())
            doc.close()
            return pixmap
        except Exception:
            return None

    def init_ui(self):
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)

        self.file_label = QLabel(
            self.texts["current_file_info"].format(
                filename=self.pdf_name,
                page_count=self.page_count
            )
        )
        self.file_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        main_layout.addWidget(self.file_label)

        mode_layout = QHBoxLayout()
        mode_layout.setSpacing(20)

        self.manual_radio = QRadioButton(self.texts["manual_select"])
        self.auto_radio = QRadioButton(self.texts["auto_select"])
        self.manual_radio.setChecked(True)

        self.mode_group = QButtonGroup(self)
        self.mode_group.addButton(self.manual_radio)
        self.mode_group.addButton(self.auto_radio)

        self.manual_radio.toggled.connect(self.update_mode_ui)
        self.auto_radio.toggled.connect(self.update_mode_ui)

        mode_layout.addWidget(self.manual_radio)
        mode_layout.addWidget(self.auto_radio)
        mode_layout.addStretch()

        main_layout.addLayout(mode_layout)

        self.manual_tip_label = QLabel(self.texts["manual_tip"])
        self.manual_tip_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        main_layout.addWidget(self.manual_tip_label)

        self.auto_input_container = QWidget()
        auto_input_layout = QHBoxLayout()
        auto_input_layout.setContentsMargins(0, 0, 0, 0)
        auto_input_layout.setSpacing(10)

        self.auto_label_left = QLabel(self.texts["every"])
        self.auto_page_input = QLineEdit()
        self.auto_page_input.setFixedWidth(100)
        self.auto_page_input.setPlaceholderText(self.texts["input_page_count"])
        self.auto_page_input.setValidator(QIntValidator(1, 999999, self))
        self.auto_label_right = QLabel(self.texts["pages_auto_split"])

        auto_input_layout.addWidget(self.auto_label_left)
        auto_input_layout.addWidget(self.auto_page_input)
        auto_input_layout.addWidget(self.auto_label_right)
        auto_input_layout.addStretch()

        self.auto_input_container.setLayout(auto_input_layout)
        main_layout.addWidget(self.auto_input_container)

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)

        self.pages_container = QWidget()
        self.pages_layout = QGridLayout()
        self.pages_layout.setSpacing(15)
        self.pages_layout.setContentsMargins(10, 10, 10, 10)

        self.pages_container.setLayout(self.pages_layout)
        self.scroll_area.setWidget(self.pages_container)

        main_layout.addWidget(self.scroll_area)

        bottom_layout = QHBoxLayout()
        bottom_layout.setSpacing(20)

        self.finish_button = QPushButton(self.texts["finish"])
        self.finish_button.setFixedHeight(45)
        self.finish_button.clicked.connect(self.finish_split)

        self.cancel_button = QPushButton(self.texts["cancel"])
        self.cancel_button.setFixedHeight(45)
        self.cancel_button.clicked.connect(self.close)

        bottom_layout.addWidget(self.finish_button)
        bottom_layout.addWidget(self.cancel_button)

        main_layout.addLayout(bottom_layout)

        self.setLayout(main_layout)

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
            QPushButton {
                background-color: #e8f0fe;
                color: black;
                border: 1px solid #b7c9e2;
                border-radius: 8px;
                padding: 8px;
            }
            QPushButton:hover {
                background-color: #dbe9ff;
                color: black;
            }
            QLineEdit {
                background-color: white;
                color: black;
                border: 1px solid #cfcfcf;
                border-radius: 6px;
                padding: 6px;
            }
            QScrollArea {
                border: 1px solid #cfcfcf;
                border-radius: 8px;
                background-color: white;
            }
            QRadioButton {
                color: black;
            }
        """)

        self.populate_page_placeholders()
        self.update_mode_ui()

    def populate_page_placeholders(self):
        columns = 4

        for page_index in range(self.page_count):
            page_number = page_index + 1

            page_frame = QFrame()
            page_frame.setFrameShape(QFrame.StyledPanel)
            page_frame.setStyleSheet("""
                QFrame {
                    border: 1px solid #bfbfbf;
                    border-radius: 8px;
                    background-color: #f8f8f8;
                }
            """)
            page_frame.setFixedSize(190, 320)

            page_layout = QVBoxLayout()
            page_layout.setContentsMargins(10, 10, 10, 10)
            page_layout.setSpacing(8)

            preview_label = QLabel()
            preview_label.setAlignment(Qt.AlignCenter)
            preview_label.setStyleSheet("border: none;")
            preview_label.setFixedHeight(220)

            pixmap = self.render_pdf_page_thumbnail(page_index)

            if pixmap is not None:
                scaled_pixmap = pixmap.scaled(
                    150, 210,
                    Qt.KeepAspectRatio,
                    Qt.SmoothTransformation
                )
                preview_label.setPixmap(scaled_pixmap)
            else:
                preview_label.setText(self.texts["thumbnail_load_failed"])
                preview_label.setStyleSheet("border: none; font-size: 14px; color: red;")

            page_num_label = QLabel(
                self.texts["page_label"].format(page_number=page_number)
            )
            page_num_label.setAlignment(Qt.AlignCenter)
            page_num_label.setStyleSheet("border: none; font-weight: bold;")

            split_button = QPushButton(self.texts["split_after_page"])
            split_button.setFixedHeight(36)
            split_button.clicked.connect(
                lambda checked=False, page=page_number, btn=split_button: self.toggle_split_point(page, btn)
            )

            if page_number == self.page_count:
                split_button.setText(self.texts["last_page"])
                split_button.setEnabled(False)

            page_layout.addWidget(preview_label)
            page_layout.addWidget(page_num_label)
            page_layout.addWidget(split_button)

            page_frame.setLayout(page_layout)

            row = page_index // columns
            col = page_index % columns
            self.pages_layout.addWidget(page_frame, row, col)

            QCoreApplication.processEvents()

        self.pages_container.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Maximum)

    def toggle_split_point(self, page_number, button):
        if page_number in self.split_points:
            self.split_points.remove(page_number)
            button.setText(self.texts["split_after_page"])
            button.setStyleSheet("")
        else:
            self.split_points.add(page_number)
            button.setText(self.texts["split_marked"])
            button.setStyleSheet("""
                QPushButton {
                    background-color: #d4edda;
                    color: black;
                    border: 1px solid #8ccf9b;
                    border-radius: 8px;
                    padding: 8px;
                }
                QPushButton:hover {
                    background-color: #c3e6cb;
                }
            """)

    def update_mode_ui(self):
        if self.manual_radio.isChecked():
            self.manual_tip_label.show()
            self.scroll_area.show()
            self.auto_input_container.hide()
        else:
            self.manual_tip_label.hide()
            self.scroll_area.hide()
            self.auto_input_container.show()

    def build_manual_ranges(self):
        if self.page_count <= 0:
            return []

        ranges = []
        sorted_points = sorted(self.split_points)

        start_page = 1
        for split_page in sorted_points:
            if split_page >= start_page:
                ranges.append((start_page, split_page))
                start_page = split_page + 1

        if start_page <= self.page_count:
            ranges.append((start_page, self.page_count))

        return ranges

    def build_auto_ranges(self, pages_per_split):
        ranges = []
        start = 1

        while start <= self.page_count:
            end = min(start + pages_per_split - 1, self.page_count)
            ranges.append((start, end))
            start = end + 1

        return ranges

    def export_ranges_to_pdfs(self, ranges, output_folder):
        reader = PdfReader(self.pdf_path)
        output_files = []

        for i, (start_page, end_page) in enumerate(ranges, start=1):
            writer = PdfWriter()

            for page_index in range(start_page - 1, end_page):
                writer.add_page(reader.pages[page_index])

            output_filename = f"{self.pdf_base_name}_part_{i}.pdf"
            output_path = os.path.join(output_folder, output_filename)

            with open(output_path, "wb") as f:
                writer.write(f)

            output_files.append(output_path)

        return output_files

    def show_busy_progress(self, title, label_text):
        progress = QProgressDialog(label_text, None, 0, 0, self)
        progress.setWindowTitle(title)
        progress.setWindowModality(Qt.ApplicationModal)
        progress.setMinimumDuration(0)
        progress.setAutoClose(False)
        progress.setAutoReset(False)
        progress.setCancelButton(None)
        progress.show()
        QCoreApplication.processEvents()
        return progress

    def finish_split(self):
        if self.auto_radio.isChecked():
            text = self.auto_page_input.text().strip()

            if not text:
                QMessageBox.warning(
                    self,
                    self.texts["tip"],
                    self.texts["input_auto_pages"]
                )
                return

            try:
                pages_per_split = int(text)
            except ValueError:
                QMessageBox.warning(
                    self,
                    self.texts["tip"],
                    self.texts["invalid_number"]
                )
                return

            if pages_per_split <= 0:
                QMessageBox.warning(
                    self,
                    self.texts["tip"],
                    self.texts["page_must_positive"]
                )
                return

            output_folder = QFileDialog.getExistingDirectory(
                self,
                self.texts["choose_split_folder"]
            )

            if not output_folder:
                return

            progress = self.show_busy_progress(
                self.texts["split_progress_title"],
                self.texts["split_auto_progress"]
            )

            try:
                ranges = self.build_auto_ranges(pages_per_split)
                output_files = self.export_ranges_to_pdfs(ranges, output_folder)
                progress.close()

                QMessageBox.information(
                    self,
                    self.texts["success"],
                    self.texts["split_auto_success"].format(
                        count=len(output_files),
                        folder=output_folder
                    )
                )
            except Exception as e:
                progress.close()
                QMessageBox.critical(
                    self,
                    self.texts["error"],
                    f"{self.texts['split_auto_fail']}\n{e}"
                )

        else:
            ranges = self.build_manual_ranges()
            if not ranges:
                QMessageBox.warning(
                    self,
                    self.texts["tip"],
                    self.texts["manual_range_invalid"]
                )
                return

            output_folder = QFileDialog.getExistingDirectory(
                self,
                self.texts["choose_split_folder"]
            )

            if not output_folder:
                return

            progress = self.show_busy_progress(
                self.texts["split_progress_title"],
                self.texts["split_manual_progress"]
            )

            try:
                output_files = self.export_ranges_to_pdfs(ranges, output_folder)
                progress.close()

                QMessageBox.information(
                    self,
                    self.texts["success"],
                    self.texts["split_manual_success"].format(
                        count=len(output_files),
                        folder=output_folder
                    )
                )
            except Exception as e:
                progress.close()
                QMessageBox.critical(
                    self,
                    self.texts["error"],
                    f"{self.texts['split_manual_fail']}\n{e}"
                )