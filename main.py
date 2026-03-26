import os
import sys
import fitz

from pypdf import PdfReader, PdfWriter
from pdf2docx import Converter

from split_window import SplitWindow
from process_window import ProcessWindow
from language_dialog import LanguageDialog
from translations import TEXTS

from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QPushButton,
    QListWidget,
    QListWidgetItem,
    QFileDialog,
    QVBoxLayout,
    QHBoxLayout,
    QMessageBox,
    QLabel,
    QAbstractItemView,
    QProgressDialog,
)
from PySide6.QtCore import Qt, QCoreApplication
from PySide6.QtGui import QDragEnterEvent, QDropEvent


class PDFDropListWidget(QListWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            for url in urls:
                if url.isLocalFile() and url.toLocalFile().lower().endswith(".pdf"):
                    event.acceptProposedAction()
                    return
        event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            for url in urls:
                if url.isLocalFile() and url.toLocalFile().lower().endswith(".pdf"):
                    event.acceptProposedAction()
                    return
        event.ignore()

    def dropEvent(self, event: QDropEvent):
        if not event.mimeData().hasUrls():
            event.ignore()
            return

        pdf_files = []
        for url in event.mimeData().urls():
            if url.isLocalFile():
                file_path = url.toLocalFile()
                if file_path.lower().endswith(".pdf"):
                    pdf_files.append(file_path)

        if pdf_files:
            parent_window = self.window()
            if hasattr(parent_window, "add_pdf_files_from_list"):
                parent_window.add_pdf_files_from_list(pdf_files)
            event.acceptProposedAction()
        else:
            event.ignore()


class PDFToolMainWindow(QWidget):
    def __init__(self, lang="zh"):
        super().__init__()
        self.lang = lang
        self.texts = TEXTS[self.lang]

        self.setWindowTitle(self.texts["main_title"])
        self.resize(760, 600)

        self.pdf_files = []

        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)

        # 顶部：添加 PDF 文件
        self.add_pdf_button = QPushButton(self.texts["add_pdf"])
        self.add_pdf_button.setFixedHeight(50)
        self.add_pdf_button.clicked.connect(self.add_pdf_files)
        main_layout.addWidget(self.add_pdf_button)

        # 列表标题 + 清空全部
        list_top_layout = QHBoxLayout()
        list_top_layout.setSpacing(10)

        self.info_label = QLabel(self.texts["added_pdf_label"])
        self.info_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)

        self.clear_all_button = QPushButton(self.texts["clear_all"])
        self.clear_all_button.setFixedHeight(36)
        self.clear_all_button.setFixedWidth(110)
        self.clear_all_button.clicked.connect(self.clear_all_pdfs)

        list_top_layout.addWidget(self.info_label)
        list_top_layout.addStretch()
        list_top_layout.addWidget(self.clear_all_button)

        main_layout.addLayout(list_top_layout)

        # 中间：列表 + 右侧删除按钮
        center_layout = QHBoxLayout()
        center_layout.setSpacing(12)

        self.pdf_list_widget = PDFDropListWidget()
        self.pdf_list_widget.setSelectionMode(QAbstractItemView.SingleSelection)
        self.pdf_list_widget.setAlternatingRowColors(True)
        self.pdf_list_widget.itemSelectionChanged.connect(self.update_delete_button_state)
        center_layout.addWidget(self.pdf_list_widget, 1)

        right_button_layout = QVBoxLayout()
        right_button_layout.setSpacing(10)

        self.delete_button = QPushButton(self.texts["delete"])
        self.delete_button.setFixedSize(90, 40)
        self.delete_button.clicked.connect(self.delete_selected_pdf)
        self.delete_button.setEnabled(False)

        right_button_layout.addWidget(self.delete_button)
        right_button_layout.addStretch()

        center_layout.addLayout(right_button_layout)
        main_layout.addLayout(center_layout)

        # 底部按钮区域：两行
        bottom_layout = QVBoxLayout()
        bottom_layout.setSpacing(12)

        # 第一行
        top_button_row = QHBoxLayout()
        top_button_row.setSpacing(20)

        self.merge_button = QPushButton(self.texts["merge"])
        self.merge_button.setFixedHeight(45)
        self.merge_button.clicked.connect(self.merge_pdfs_placeholder)

        self.split_button = QPushButton(self.texts["split"])
        self.split_button.setFixedHeight(45)
        self.split_button.clicked.connect(self.split_pdfs_placeholder)

        self.process_button = QPushButton(self.texts["process"])
        self.process_button.setFixedHeight(45)
        self.process_button.clicked.connect(self.open_process_window)

        top_button_row.addWidget(self.merge_button)
        top_button_row.addWidget(self.split_button)
        top_button_row.addWidget(self.process_button)

        # 第二行
        bottom_button_row = QHBoxLayout()
        bottom_button_row.setSpacing(20)

        self.pdf_to_word_button = QPushButton(self.texts["pdf_to_word"])
        self.pdf_to_word_button.setFixedHeight(45)
        self.pdf_to_word_button.clicked.connect(self.pdf_to_word)

        self.pdf_to_jpg_button = QPushButton(self.texts["pdf_to_jpg"])
        self.pdf_to_jpg_button.setFixedHeight(45)
        self.pdf_to_jpg_button.clicked.connect(self.pdf_to_jpg)

        bottom_button_row.addWidget(self.pdf_to_word_button)
        bottom_button_row.addWidget(self.pdf_to_jpg_button)

        bottom_layout.addLayout(top_button_row)
        bottom_layout.addLayout(bottom_button_row)
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
                padding: 6px;
            }
        """)

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

    def add_pdf_files_from_list(self, file_paths):
        if not file_paths:
            return

        added_count = 0

        for file_path in file_paths:
            if file_path.lower().endswith(".pdf") and file_path not in self.pdf_files:
                self.pdf_files.append(file_path)
                file_name = os.path.basename(file_path)

                item = QListWidgetItem(file_name)
                item.setToolTip(file_path)
                self.pdf_list_widget.addItem(item)
                added_count += 1

        if added_count == 0:
            QMessageBox.information(self, self.texts["tip"], self.texts["msg_already_added"])

        self.update_delete_button_state()

    def add_pdf_files(self):
        file_paths, _ = QFileDialog.getOpenFileNames(
            self,
            self.texts["choose_pdf"],
            "",
            "PDF Files (*.pdf)"
        )
        self.add_pdf_files_from_list(file_paths)

    def delete_selected_pdf(self):
        current_row = self.pdf_list_widget.currentRow()

        if current_row < 0:
            QMessageBox.warning(self, self.texts["tip"], self.texts["msg_select_pdf_process"])
            return

        reply = QMessageBox.question(
            self,
            self.texts["confirm_delete"],
            self.texts["msg_delete_confirm"],
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            self.pdf_list_widget.takeItem(current_row)
            del self.pdf_files[current_row]

        self.update_delete_button_state()

    def clear_all_pdfs(self):
        if not self.pdf_files:
            QMessageBox.information(self, self.texts["tip"], self.texts["msg_no_clear"])
            return

        reply = QMessageBox.question(
            self,
            self.texts["confirm_clear"],
            self.texts["msg_clear_confirm"],
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            self.pdf_files.clear()
            self.pdf_list_widget.clear()

        self.update_delete_button_state()

    def update_delete_button_state(self):
        has_selection = self.pdf_list_widget.currentRow() >= 0
        self.delete_button.setEnabled(has_selection)

    def merge_pdfs(self, input_paths, output_path):
        writer = PdfWriter()

        for pdf_path in input_paths:
            reader = PdfReader(pdf_path)
            for page in reader.pages:
                writer.add_page(page)

        with open(output_path, "wb") as f:
            writer.write(f)

    def merge_pdfs_placeholder(self):
        if len(self.pdf_files) < 2:
            QMessageBox.warning(self, self.texts["tip"], self.texts["msg_need_two_pdf"])
            return

        output_path, _ = QFileDialog.getSaveFileName(
            self,
            self.texts["choose_merge_output"],
            "merged.pdf",
            "PDF Files (*.pdf)"
        )

        if not output_path:
            return

        if not output_path.lower().endswith(".pdf"):
            output_path += ".pdf"

        progress = self.show_busy_progress(
            self.texts["progress_merge_title"],
            self.texts["progress_merge"]
        )

        try:
            self.merge_pdfs(self.pdf_files, output_path)
            progress.close()
            QMessageBox.information(
                self,
                self.texts["success"],
                f"{self.texts['msg_merge_success']}\n\n{output_path}"
            )
        except Exception as e:
            progress.close()
            QMessageBox.critical(
                self,
                self.texts["error"],
                f"{self.texts['msg_merge_fail']}\n{e}"
            )

    def split_pdfs_placeholder(self):
        if len(self.pdf_files) < 1:
            QMessageBox.warning(self, self.texts["tip"], self.texts["msg_need_one_pdf"])
            return

        current_row = self.pdf_list_widget.currentRow()
        if current_row < 0:
            QMessageBox.warning(self, self.texts["tip"], self.texts["msg_select_pdf_split"])
            return

        selected_pdf_path = self.pdf_files[current_row]

        # 这里暂时先不传语言，下一步再把 split_window.py 做成双语
        self.split_window = SplitWindow(selected_pdf_path, lang=self.lang)
        self.split_window.show()

    def open_process_window(self):
        if len(self.pdf_files) < 1:
            QMessageBox.warning(self, self.texts["tip"], self.texts["msg_need_one_pdf"])
            return

        current_row = self.pdf_list_widget.currentRow()
        if current_row < 0:
            QMessageBox.warning(self, self.texts["tip"], self.texts["msg_select_pdf_process"])
            return

        selected_pdf_path = self.pdf_files[current_row]

        # 这里暂时先不传语言，下一步再把 process_window.py 做成双语
        self.process_window = ProcessWindow(selected_pdf_path, lang=self.lang)
        self.process_window.show()

    def convert_pdf_to_word(self, input_pdf_path, output_docx_path):
        cv = Converter(input_pdf_path)
        try:
            cv.convert(output_docx_path, start=0, end=None)
        finally:
            cv.close()

    def pdf_to_word(self):
        if len(self.pdf_files) < 1:
            QMessageBox.warning(self, self.texts["tip"], self.texts["msg_need_one_pdf"])
            return

        current_row = self.pdf_list_widget.currentRow()
        if current_row < 0:
            QMessageBox.warning(self, self.texts["tip"], self.texts["msg_select_pdf_convert"])
            return

        selected_pdf_path = self.pdf_files[current_row]
        default_name = os.path.splitext(os.path.basename(selected_pdf_path))[0] + ".docx"

        output_path, _ = QFileDialog.getSaveFileName(
            self,
            self.texts["choose_word_output"],
            default_name,
            "Word Files (*.docx)"
        )

        if not output_path:
            return

        if not output_path.lower().endswith(".docx"):
            output_path += ".docx"

        progress = self.show_busy_progress(
            self.texts["progress_convert_title"],
            self.texts["progress_word"]
        )

        try:
            self.convert_pdf_to_word(selected_pdf_path, output_path)
            progress.close()
            QMessageBox.information(
                self,
                self.texts["success"],
                f"{self.texts['msg_word_success']}\n\n{output_path}"
            )
        except Exception as e:
            progress.close()
            QMessageBox.critical(
                self,
                self.texts["error"],
                f"{self.texts['msg_word_fail']}\n{e}"
            )

    def convert_pdf_to_jpg(self, input_pdf_path, output_folder, zoom=2.0):
        pdf_name = os.path.splitext(os.path.basename(input_pdf_path))[0]
        doc = fitz.open(input_pdf_path)

        try:
            for page_index in range(len(doc)):
                page = doc.load_page(page_index)
                matrix = fitz.Matrix(zoom, zoom)
                pix = page.get_pixmap(matrix=matrix, alpha=False)

                output_filename = f"{pdf_name}_page_{page_index + 1}.jpg"
                output_path = os.path.join(output_folder, output_filename)
                pix.save(output_path)
        finally:
            doc.close()

    def pdf_to_jpg(self):
        if len(self.pdf_files) < 1:
            QMessageBox.warning(self, self.texts["tip"], self.texts["msg_need_one_pdf"])
            return

        current_row = self.pdf_list_widget.currentRow()
        if current_row < 0:
            QMessageBox.warning(self, self.texts["tip"], self.texts["msg_select_pdf_convert"])
            return

        selected_pdf_path = self.pdf_files[current_row]

        output_folder = QFileDialog.getExistingDirectory(
            self,
            self.texts["choose_jpg_output"]
        )

        if not output_folder:
            return

        progress = self.show_busy_progress(
            self.texts["progress_convert_title"],
            self.texts["progress_jpg"]
        )

        try:
            self.convert_pdf_to_jpg(selected_pdf_path, output_folder)
            progress.close()
            QMessageBox.information(
                self,
                self.texts["success"],
                f"{self.texts['msg_jpg_success']}\n\n{output_folder}"
            )
        except Exception as e:
            progress.close()
            QMessageBox.critical(
                self,
                self.texts["error"],
                f"{self.texts['msg_jpg_fail']}\n{e}"
            )


if __name__ == "__main__":
    app = QApplication(sys.argv)

    lang_dialog = LanguageDialog()
    result = lang_dialog.exec()

    if result:
        selected_lang = lang_dialog.selected_lang or "zh"
        window = PDFToolMainWindow(lang=selected_lang)
        window.show()
        sys.exit(app.exec())
    else:
        sys.exit(0)