import os
import sys
import fitz
from pdf2docx import Converter
from pypdf import PdfReader, PdfWriter
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
)
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QProgressDialog
from PySide6.QtCore import QCoreApplication
from PySide6.QtGui import QDragEnterEvent, QDropEvent

from split_window import SplitWindow
from process_window import ProcessWindow

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
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF 处理工具")
        self.resize(760, 540)

        # 保存已添加 PDF 的完整路径
        self.pdf_files = []

        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)

        # ===== 顶部：添加 PDF 文件 =====
        self.add_pdf_button = QPushButton("添加 PDF 文件")
        self.add_pdf_button.setFixedHeight(50)
        self.add_pdf_button.clicked.connect(self.add_pdf_files)
        main_layout.addWidget(self.add_pdf_button)

        # ===== 列表标题 + 清空全部 =====
        list_top_layout = QHBoxLayout()
        list_top_layout.setSpacing(10)

        self.info_label = QLabel("已添加的 PDF 文件（也可直接拖入 PDF 到下方区域）：")
        self.info_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)

        self.clear_all_button = QPushButton("清空全部")
        self.clear_all_button.setFixedHeight(36)
        self.clear_all_button.setFixedWidth(110)
        self.clear_all_button.clicked.connect(self.clear_all_pdfs)

        list_top_layout.addWidget(self.info_label)
        list_top_layout.addStretch()
        list_top_layout.addWidget(self.clear_all_button)

        main_layout.addLayout(list_top_layout)

        # ===== 中间：列表 + 右侧删除按钮 =====
        center_layout = QHBoxLayout()
        center_layout.setSpacing(12)

        self.pdf_list_widget = PDFDropListWidget()
        self.pdf_list_widget.setSelectionMode(QAbstractItemView.SingleSelection)
        self.pdf_list_widget.setAlternatingRowColors(True)
        self.pdf_list_widget.itemSelectionChanged.connect(self.update_delete_button_state)
        center_layout.addWidget(self.pdf_list_widget, 1)

        right_button_layout = QVBoxLayout()
        right_button_layout.setSpacing(10)

        self.delete_button = QPushButton("删除")
        self.delete_button.setFixedSize(90, 40)
        self.delete_button.clicked.connect(self.delete_selected_pdf)
        self.delete_button.setEnabled(False)

        right_button_layout.addWidget(self.delete_button)
        right_button_layout.addStretch()

        center_layout.addLayout(right_button_layout)

        main_layout.addLayout(center_layout)

        # ===== 底部按钮区域：两行 =====
        bottom_layout = QVBoxLayout()
        bottom_layout.setSpacing(12)

        # 第一行：合并 / 拆分 / 处理
        top_button_row = QHBoxLayout()
        top_button_row.setSpacing(20)

        self.merge_button = QPushButton("合并")
        self.merge_button.setFixedHeight(45)
        self.merge_button.clicked.connect(self.merge_pdfs_placeholder)

        self.split_button = QPushButton("拆分")
        self.split_button.setFixedHeight(45)
        self.split_button.clicked.connect(self.split_pdfs_placeholder)

        self.process_button = QPushButton("处理")
        self.process_button.setFixedHeight(45)
        self.process_button.clicked.connect(self.open_process_window)

        top_button_row.addWidget(self.merge_button)
        top_button_row.addWidget(self.split_button)
        top_button_row.addWidget(self.process_button)

        # 第二行：PDF转Word / PDF转JPG
        bottom_button_row = QHBoxLayout()
        bottom_button_row.setSpacing(20)

        self.pdf_to_word_button = QPushButton("PDF转Word")
        self.pdf_to_word_button.setFixedHeight(45)
        self.pdf_to_word_button.clicked.connect(self.pdf_to_word)

        self.pdf_to_jpg_button = QPushButton("PDF转JPG")
        self.pdf_to_jpg_button.setFixedHeight(45)
        self.pdf_to_jpg_button.clicked.connect(self.pdf_to_jpg)

        bottom_button_row.addWidget(self.pdf_to_word_button)
        bottom_button_row.addWidget(self.pdf_to_jpg_button)

        bottom_layout.addLayout(top_button_row)
        bottom_layout.addLayout(bottom_button_row)

        main_layout.addLayout(bottom_layout)

        self.setLayout(main_layout)

        # ===== 样式 =====
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

    def add_pdf_files(self):
        """通过文件选择框添加多个 PDF 文件"""
        file_paths, _ = QFileDialog.getOpenFileNames(
            self,
            "选择 PDF 文件",
            "",
            "PDF Files (*.pdf)"
        )

        self.add_pdf_files_from_list(file_paths)

    def add_pdf_files_from_list(self, file_paths):
        """从文件路径列表中添加 PDF 文件"""
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
            QMessageBox.information(self, "提示", "拖入或选择的 PDF 文件都已经添加过了。")

        self.update_delete_button_state()

    def delete_selected_pdf(self):
        """删除当前选中的 PDF"""
        current_row = self.pdf_list_widget.currentRow()

        if current_row < 0:
            QMessageBox.warning(self, "提示", "请先选中一个 PDF 文件。")
            return

        reply = QMessageBox.question(
            self,
            "确认删除",
            "确定要删除当前选中的 PDF 文件吗？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            self.pdf_list_widget.takeItem(current_row)
            del self.pdf_files[current_row]

        self.update_delete_button_state()

    def clear_all_pdfs(self):
        """清空全部 PDF"""
        if not self.pdf_files:
            QMessageBox.information(self, "提示", "当前没有可清空的 PDF 文件。")
            return

        reply = QMessageBox.question(
            self,
            "确认清空",
            "确定要清空全部已添加的 PDF 文件吗？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            self.pdf_files.clear()
            self.pdf_list_widget.clear()

        self.update_delete_button_state()

    def update_delete_button_state(self):
        """根据是否选中条目，更新删除按钮状态"""
        has_selection = self.pdf_list_widget.currentRow() >= 0
        self.delete_button.setEnabled(has_selection)

    def show_busy_progress(self, title, label_text):
        """显示忙碌状态的进度条对话框"""
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

    def merge_pdfs(self, input_paths, output_path):
        """按顺序合并多个 PDF"""
        writer = PdfWriter()

        for pdf_path in input_paths:
            reader = PdfReader(pdf_path)
            for page in reader.pages:
                writer.add_page(page)

        with open(output_path, "wb") as f:
            writer.write(f)

    def merge_pdfs_placeholder(self):
        """真正执行 PDF 合并，并显示进度条"""
        if len(self.pdf_files) < 2:
            QMessageBox.warning(self, "提示", "至少需要添加 2 个 PDF 文件才能合并。")
            return

        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "选择合并后的输出位置",
            "merged.pdf",
            "PDF Files (*.pdf)"
        )

        if not output_path:
            return

        if not output_path.lower().endswith(".pdf"):
            output_path += ".pdf"

        progress = self.show_busy_progress("合并中", "正在合并 PDF，请稍候...")

        try:
            self.merge_pdfs(self.pdf_files, output_path)
            progress.close()

            QMessageBox.information(
                self,
                "合并成功",
                f"PDF 合并完成！\n\n输出文件：\n{output_path}"
            )
        except Exception as e:
            progress.close()
            QMessageBox.critical(
                self,
                "错误",
                f"PDF 合并失败：\n{e}"
            )

    def split_pdfs_placeholder(self):
        """打开拆分窗口"""
        if len(self.pdf_files) < 1:
            QMessageBox.warning(self, "提示", "请先添加至少 1 个 PDF 文件。")
            return

        current_row = self.pdf_list_widget.currentRow()
        if current_row < 0:
            QMessageBox.warning(self, "提示", "请先在列表中选中一个要拆分的 PDF 文件。")
            return

        selected_pdf_path = self.pdf_files[current_row]

        self.split_window = SplitWindow(selected_pdf_path)
        self.split_window.show()

    def open_process_window(self):
        """打开处理窗口"""
        if len(self.pdf_files) < 1:
            QMessageBox.warning(self, "提示", "请先添加至少 1 个 PDF 文件。")
            return

        current_row = self.pdf_list_widget.currentRow()
        if current_row < 0:
            QMessageBox.warning(self, "提示", "请先在列表中选中一个要处理的 PDF 文件。")
            return

        selected_pdf_path = self.pdf_files[current_row]

        self.process_window = ProcessWindow(selected_pdf_path)
        self.process_window.show()

    def convert_pdf_to_word(self, input_pdf_path, output_docx_path):
        """把 PDF 转成 Word"""
        cv = Converter(input_pdf_path)
        try:
            cv.convert(output_docx_path, start=0, end=None)
        finally:
            cv.close()

    def pdf_to_word(self):
        """执行 PDF 转 Word"""
        if len(self.pdf_files) < 1:
            QMessageBox.warning(self, "提示", "请先添加至少 1 个 PDF 文件。")
            return

        current_row = self.pdf_list_widget.currentRow()
        if current_row < 0:
            QMessageBox.warning(self, "提示", "请先在列表中选中一个要转换的 PDF 文件。")
            return

        selected_pdf_path = self.pdf_files[current_row]
        default_name = os.path.splitext(os.path.basename(selected_pdf_path))[0] + ".docx"

        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "选择 Word 输出位置",
            default_name,
            "Word Files (*.docx)"
        )

        if not output_path:
            return

        if not output_path.lower().endswith(".docx"):
            output_path += ".docx"

        progress = self.show_busy_progress("转换中", "正在将 PDF 转换为 Word，请稍候...")

        try:
            self.convert_pdf_to_word(selected_pdf_path, output_path)
            progress.close()

            QMessageBox.information(
                self,
                "转换成功",
                f"PDF 转 Word 完成！\n\n输出文件：\n{output_path}"
            )
        except Exception as e:
            progress.close()
            QMessageBox.critical(
                self,
                "错误",
                f"PDF 转 Word 失败：\n{e}"
            )

    def convert_pdf_to_jpg(self, input_pdf_path, output_folder, zoom=2.0):
        """把 PDF 每一页导出为 JPG"""
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
        """执行 PDF 转 JPG"""
        if len(self.pdf_files) < 1:
            QMessageBox.warning(self, "提示", "请先添加至少 1 个 PDF 文件。")
            return

        current_row = self.pdf_list_widget.currentRow()
        if current_row < 0:
            QMessageBox.warning(self, "提示", "请先在列表中选中一个要转换的 PDF 文件。")
            return

        selected_pdf_path = self.pdf_files[current_row]

        output_folder = QFileDialog.getExistingDirectory(
            self,
            "选择 JPG 输出文件夹"
        )

        if not output_folder:
            return

        progress = self.show_busy_progress("转换中", "正在将 PDF 转换为 JPG，请稍候...")

        try:
            self.convert_pdf_to_jpg(selected_pdf_path, output_folder)
            progress.close()

            QMessageBox.information(
                self,
                "转换成功",
                f"PDF 转 JPG 完成！\n\n输出文件夹：\n{output_folder}"
            )
        except Exception as e:
            progress.close()
            QMessageBox.critical(
                self,
                "错误",
                f"PDF 转 JPG 失败：\n{e}"
            )

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PDFToolMainWindow()
    window.show()
    sys.exit(app.exec())