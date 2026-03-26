import os
import fitz

from PySide6.QtWidgets import (
    QWidget,
    QLabel,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QFileDialog,
    QMessageBox,
    QListWidget,
    QListWidgetItem,
    QListView,
    QProgressDialog,
    QAbstractItemView,
)
from PySide6.QtCore import Qt, QSize, QCoreApplication
from PySide6.QtGui import QPixmap, QImage, QIcon

from pypdf import PdfReader, PdfWriter
from translations import TEXTS


class ProcessWindow(QWidget):
    def __init__(self, pdf_path, lang="zh"):
        super().__init__()
        self.lang = lang
        self.texts = TEXTS[self.lang]

        self.pdf_path = pdf_path
        self.pdf_name = os.path.basename(pdf_path)
        self.pdf_base_name = os.path.splitext(self.pdf_name)[0]
        self.page_count = self.get_pdf_page_count(pdf_path)

        self.setWindowTitle(self.texts["process_window_title"])
        self.resize(980, 820)

        self.init_ui()
        self.load_page_thumbnails()

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

    def render_pdf_page_thumbnail(self, page_index, zoom=0.18):
        try:
            doc = fitz.open(self.pdf_path)
            page = doc.load_page(page_index)

            matrix = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=matrix, alpha=False)

            qimage = QImage(
                pix.samples,
                pix.width,
                pix.height,
                pix.stride,
                QImage.Format_RGB888
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

        self.tip_label = QLabel(self.texts["process_tip"])
        self.tip_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        main_layout.addWidget(self.tip_label)

        center_layout = QHBoxLayout()
        center_layout.setSpacing(15)

        self.page_list_widget = QListWidget()
        self.page_list_widget.setViewMode(QListView.ListMode)
        self.page_list_widget.setFlow(QListView.TopToBottom)
        self.page_list_widget.setWrapping(False)
        self.page_list_widget.setIconSize(QSize(70, 95))
        self.page_list_widget.setSpacing(4)
        self.page_list_widget.setResizeMode(QListView.Adjust)
        self.page_list_widget.setMovement(QListView.Static)
        self.page_list_widget.setSelectionMode(QAbstractItemView.SingleSelection)
        self.page_list_widget.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.page_list_widget.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
        self.page_list_widget.itemSelectionChanged.connect(self.update_action_buttons)

        center_layout.addWidget(self.page_list_widget, 1)

        right_button_layout = QVBoxLayout()
        right_button_layout.setSpacing(12)

        self.move_top_button = QPushButton(self.texts["move_first"])
        self.move_top_button.setFixedSize(110, 45)
        self.move_top_button.clicked.connect(self.move_selected_to_top)

        self.move_up_button = QPushButton(self.texts["move_up"])
        self.move_up_button.setFixedSize(110, 45)
        self.move_up_button.clicked.connect(self.move_selected_up)

        self.delete_button = QPushButton(self.texts["delete_current_page"])
        self.delete_button.setFixedSize(110, 45)
        self.delete_button.clicked.connect(self.delete_current_page)

        self.move_down_button = QPushButton(self.texts["move_down"])
        self.move_down_button.setFixedSize(110, 45)
        self.move_down_button.clicked.connect(self.move_selected_down)

        self.move_bottom_button = QPushButton(self.texts["move_last"])
        self.move_bottom_button.setFixedSize(110, 45)
        self.move_bottom_button.clicked.connect(self.move_selected_to_bottom)

        right_button_layout.addWidget(self.move_top_button)
        right_button_layout.addWidget(self.move_up_button)
        right_button_layout.addWidget(self.delete_button)
        right_button_layout.addWidget(self.move_down_button)
        right_button_layout.addWidget(self.move_bottom_button)
        right_button_layout.addStretch()

        center_layout.addLayout(right_button_layout)
        main_layout.addLayout(center_layout)

        bottom_layout = QHBoxLayout()
        bottom_layout.setSpacing(20)

        self.finish_button = QPushButton(self.texts["finish"])
        self.finish_button.setFixedHeight(45)
        self.finish_button.clicked.connect(self.finish_process)

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
            QPushButton:disabled {
                background-color: #f2f2f2;
                color: #888888;
                border: 1px solid #d0d0d0;
            }
            QListWidget {
                border: 1px solid #cfcfcf;
                border-radius: 8px;
                background-color: white;
                color: black;
                padding: 8px;
            }
        """)

        self.update_action_buttons()

    def load_page_thumbnails(self):
        for page_index in range(self.page_count):
            page_number = page_index + 1
            pixmap = self.render_pdf_page_thumbnail(page_index)

            item = QListWidgetItem()
            item.setText("")
            item.setTextAlignment(Qt.AlignVCenter)
            item.setSizeHint(QSize(180, 110))

            if pixmap is not None:
                scaled_pixmap = pixmap.scaled(
                    70, 95,
                    Qt.KeepAspectRatio,
                    Qt.SmoothTransformation
                )
                item.setIcon(QIcon(scaled_pixmap))

            item.setData(Qt.UserRole, page_number)
            self.page_list_widget.addItem(item)

            QCoreApplication.processEvents()

        self.refresh_item_labels()
        self.update_action_buttons()

    def refresh_item_labels(self):
        for i in range(self.page_list_widget.count()):
            item = self.page_list_widget.item(i)
            original_page_number = item.data(Qt.UserRole)
            current_position = i + 1
            item.setText(
                self.texts["process_position_label"].format(
                    position=current_position,
                    original_page=original_page_number
                )
            )

    def update_action_buttons(self):
        count = self.page_list_widget.count()
        current_row = self.page_list_widget.currentRow()

        if count == 0 or current_row < 0:
            self.move_top_button.setEnabled(False)
            self.move_up_button.setEnabled(False)
            self.delete_button.setEnabled(False)
            self.move_down_button.setEnabled(False)
            self.move_bottom_button.setEnabled(False)
            return

        self.delete_button.setEnabled(True)

        is_first = (current_row == 0)
        self.move_top_button.setEnabled(not is_first)
        self.move_up_button.setEnabled(not is_first)

        is_last = (current_row == count - 1)
        self.move_down_button.setEnabled(not is_last)
        self.move_bottom_button.setEnabled(not is_last)

    def take_current_item(self):
        current_row = self.page_list_widget.currentRow()
        if current_row < 0:
            QMessageBox.warning(
                self,
                self.texts["tip"],
                self.texts["select_page_first"]
            )
            return None, None
        item = self.page_list_widget.takeItem(current_row)
        return current_row, item

    def move_selected_to_top(self):
        current_row, item = self.take_current_item()
        if item is None:
            return
        self.page_list_widget.insertItem(0, item)
        self.page_list_widget.setCurrentRow(0)
        self.refresh_item_labels()
        self.update_action_buttons()

    def move_selected_up(self):
        current_row, item = self.take_current_item()
        if item is None:
            return
        new_row = max(0, current_row - 1)
        self.page_list_widget.insertItem(new_row, item)
        self.page_list_widget.setCurrentRow(new_row)
        self.refresh_item_labels()
        self.update_action_buttons()

    def move_selected_down(self):
        current_row, item = self.take_current_item()
        if item is None:
            return
        new_row = min(self.page_list_widget.count(), current_row + 1)
        self.page_list_widget.insertItem(new_row, item)
        self.page_list_widget.setCurrentRow(new_row)
        self.refresh_item_labels()
        self.update_action_buttons()

    def move_selected_to_bottom(self):
        current_row, item = self.take_current_item()
        if item is None:
            return
        new_row = self.page_list_widget.count()
        self.page_list_widget.insertItem(new_row, item)
        self.page_list_widget.setCurrentRow(new_row)
        self.refresh_item_labels()
        self.update_action_buttons()

    def delete_current_page(self):
        current_row = self.page_list_widget.currentRow()

        if current_row < 0:
            QMessageBox.warning(
                self,
                self.texts["tip"],
                self.texts["select_page_delete"]
            )
            return

        item = self.page_list_widget.item(current_row)
        page_number = item.data(Qt.UserRole)

        reply = QMessageBox.question(
            self,
            self.texts["confirm_delete"],
            self.texts["delete_page_confirm"].format(page_number=page_number),
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            self.page_list_widget.takeItem(current_row)
            if self.page_list_widget.count() > 0:
                self.page_list_widget.setCurrentRow(min(current_row, self.page_list_widget.count() - 1))
            self.refresh_item_labels()
            self.update_action_buttons()

    def get_current_page_order(self):
        page_order = []
        for i in range(self.page_list_widget.count()):
            item = self.page_list_widget.item(i)
            original_page_number = item.data(Qt.UserRole)
            page_order.append(original_page_number)
        return page_order

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

    def export_reordered_pdf(self, page_order, output_path):
        reader = PdfReader(self.pdf_path)
        writer = PdfWriter()

        for original_page_number in page_order:
            page_index = original_page_number - 1
            writer.add_page(reader.pages[page_index])

        with open(output_path, "wb") as f:
            writer.write(f)

    def finish_process(self):
        page_order = self.get_current_page_order()

        if not page_order:
            QMessageBox.warning(
                self,
                self.texts["tip"],
                self.texts["no_pages_process"]
            )
            return

        output_path, _ = QFileDialog.getSaveFileName(
            self,
            self.texts["choose_process_output"],
            f"{self.pdf_base_name}_processed.pdf",
            "PDF Files (*.pdf)"
        )

        if not output_path:
            return

        if not output_path.lower().endswith(".pdf"):
            output_path += ".pdf"

        progress = self.show_busy_progress(
            self.texts["process_progress_title"],
            self.texts["process_progress_text"]
        )

        try:
            self.export_reordered_pdf(page_order, output_path)
            progress.close()

            QMessageBox.information(
                self,
                self.texts["success"],
                self.texts["process_success"].format(output_path=output_path)
            )
        except Exception as e:
            progress.close()
            QMessageBox.critical(
                self,
                self.texts["error"],
                f"{self.texts['process_fail']}\n{e}"
            )