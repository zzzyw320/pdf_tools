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


class SplitWindow(QWidget):
    def __init__(self, pdf_path):
        super().__init__()
        self.pdf_path = pdf_path
        self.pdf_name = os.path.basename(pdf_path)
        self.pdf_base_name = os.path.splitext(self.pdf_name)[0]
        self.page_count = self.get_pdf_page_count(pdf_path)

        # 保存“在哪一页后切分”
        self.split_points = set()

        self.setWindowTitle("拆分 PDF")
        self.resize(980, 760)

        self.init_ui()

    def get_pdf_page_count(self, pdf_path):
        """读取 PDF 页数"""
        try:
            reader = PdfReader(pdf_path)
            return len(reader.pages)
        except Exception as e:
            QMessageBox.critical(self, "错误", f"读取 PDF 页数失败：\n{e}")
            return 0

    def init_ui(self):
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)

        # 顶部文件信息
        self.file_label = QLabel(f"当前文件：{self.pdf_name}    （共 {self.page_count} 页）")
        self.file_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        main_layout.addWidget(self.file_label)

        # 模式选择
        mode_layout = QHBoxLayout()
        mode_layout.setSpacing(20)

        self.manual_radio = QRadioButton("手动选择")
        self.auto_radio = QRadioButton("自动选择")
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

        # 手动模式提示
        self.manual_tip_label = QLabel("手动模式：点击某一页下方按钮，表示“在这一页后切分”")
        self.manual_tip_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        main_layout.addWidget(self.manual_tip_label)

        # 自动模式输入区
        self.auto_input_container = QWidget()
        auto_input_layout = QHBoxLayout()
        auto_input_layout.setContentsMargins(0, 0, 0, 0)
        auto_input_layout.setSpacing(10)

        self.auto_label_left = QLabel("每")
        self.auto_page_input = QLineEdit()
        self.auto_page_input.setFixedWidth(100)
        self.auto_page_input.setPlaceholderText("输入页数")
        self.auto_page_input.setValidator(QIntValidator(1, 999999, self))
        self.auto_label_right = QLabel("页自动分割")

        auto_input_layout.addWidget(self.auto_label_left)
        auto_input_layout.addWidget(self.auto_page_input)
        auto_input_layout.addWidget(self.auto_label_right)
        auto_input_layout.addStretch()

        self.auto_input_container.setLayout(auto_input_layout)
        main_layout.addWidget(self.auto_input_container)

        # 页面滚动区
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)

        self.pages_container = QWidget()
        self.pages_layout = QGridLayout()
        self.pages_layout.setSpacing(15)
        self.pages_layout.setContentsMargins(10, 10, 10, 10)

        self.pages_container.setLayout(self.pages_layout)
        self.scroll_area.setWidget(self.pages_container)

        main_layout.addWidget(self.scroll_area)

        # 底部按钮
        bottom_layout = QHBoxLayout()
        bottom_layout.setSpacing(20)

        self.finish_button = QPushButton("完成")
        self.finish_button.setFixedHeight(45)
        self.finish_button.clicked.connect(self.finish_split)

        self.cancel_button = QPushButton("取消")
        self.cancel_button.setFixedHeight(45)
        self.cancel_button.clicked.connect(self.close)

        bottom_layout.addWidget(self.finish_button)
        bottom_layout.addWidget(self.cancel_button)

        main_layout.addLayout(bottom_layout)

        self.setLayout(main_layout)

        # 样式
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

    def render_pdf_page_thumbnail(self, page_index, zoom=0.35):
        """把指定 PDF 页面渲染成缩略图 QPixmap"""
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

    def populate_page_placeholders(self):
        """生成真实页面缩略图，并附加切分按钮"""
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
                preview_label.setText("缩略图加载失败")
                preview_label.setStyleSheet("border: none; font-size: 14px; color: red;")

            page_num_label = QLabel(f"第 {page_number} 页")
            page_num_label.setAlignment(Qt.AlignCenter)
            page_num_label.setStyleSheet("border: none; font-weight: bold;")

            split_button = QPushButton("在此页后切分")
            split_button.setFixedHeight(36)
            split_button.clicked.connect(
                lambda checked=False, page=page_number, btn=split_button: self.toggle_split_point(page, btn)
            )

            if page_number == self.page_count:
                split_button.setText("最后一页")
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
        """切换某一页后的切分状态"""
        if page_number in self.split_points:
            self.split_points.remove(page_number)
            button.setText("在此页后切分")
            button.setStyleSheet("")
        else:
            self.split_points.add(page_number)
            button.setText("已切分")
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
        """根据模式切换界面显示"""
        if self.manual_radio.isChecked():
            self.manual_tip_label.show()
            self.scroll_area.show()
            self.auto_input_container.hide()
        else:
            self.manual_tip_label.hide()
            self.scroll_area.hide()
            self.auto_input_container.show()

    def build_manual_ranges(self):
        """把切分点转换成页范围"""
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
        """根据每N页自动分割，生成页范围"""
        ranges = []
        start = 1

        while start <= self.page_count:
            end = min(start + pages_per_split - 1, self.page_count)
            ranges.append((start, end))
            start = end + 1

        return ranges

    def format_ranges_text(self, ranges):
        """把页范围格式化成文本"""
        result = []
        for i, (start, end) in enumerate(ranges, start=1):
            if start == end:
                result.append(f"第{i}段：第 {start} 页")
            else:
                result.append(f"第{i}段：第 {start} - {end} 页")
        return "\n".join(result)

    def export_ranges_to_pdfs(self, ranges, output_folder):
        """根据页范围导出多个 PDF"""
        reader = PdfReader(self.pdf_path)
        output_files = []

        for i, (start_page, end_page) in enumerate(ranges, start=1):
            writer = PdfWriter()

            # pypdf 页码从0开始
            for page_index in range(start_page - 1, end_page):
                writer.add_page(reader.pages[page_index])

            output_filename = f"{self.pdf_base_name}_part_{i}.pdf"
            output_path = os.path.join(output_folder, output_filename)

            with open(output_path, "wb") as f:
                writer.write(f)

            output_files.append(output_path)

        return output_files

    def finish_split(self):
        """点击完成后的逻辑"""
        if self.auto_radio.isChecked():
            text = self.auto_page_input.text().strip()

            if not text:
                QMessageBox.warning(self, "提示", "请输入自动分割的页数。")
                return

            try:
                pages_per_split = int(text)
            except ValueError:
                QMessageBox.warning(self, "提示", "请输入有效的阿拉伯数字。")
                return

            if pages_per_split <= 0:
                QMessageBox.warning(self, "提示", "页数必须大于 0。")
                return

            output_folder = QFileDialog.getExistingDirectory(
                self,
                "选择拆分输出文件夹"
            )

            if not output_folder:
                return

            progress = self.show_busy_progress("拆分中", "正在自动拆分 PDF，请稍候...")

            try:
                ranges = self.build_auto_ranges(pages_per_split)
                output_files = self.export_ranges_to_pdfs(ranges, output_folder)
                progress.close()

                QMessageBox.information(
                    self,
                    "拆分成功",
                    f"自动拆分完成！\n\n共生成 {len(output_files)} 个 PDF 文件。\n输出文件夹：\n{output_folder}"
                )
            except Exception as e:
                progress.close()
                QMessageBox.critical(
                    self,
                    "错误",
                    f"自动拆分失败：\n{e}"
                )

        else:
            ranges = self.build_manual_ranges()
            if not ranges:
                QMessageBox.warning(self, "提示", "当前无法生成手动拆分范围。")
                return

            output_folder = QFileDialog.getExistingDirectory(
                self,
                "选择拆分输出文件夹"
            )

            if not output_folder:
                return

            progress = self.show_busy_progress("拆分中", "正在手动拆分 PDF，请稍候...")

            try:
                output_files = self.export_ranges_to_pdfs(ranges, output_folder)
                progress.close()

                QMessageBox.information(
                    self,
                    "拆分成功",
                    f"手动拆分完成！\n\n共生成 {len(output_files)} 个 PDF 文件。\n输出文件夹：\n{output_folder}"
                )
            except Exception as e:
                progress.close()
                QMessageBox.critical(
                    self,
                    "错误",
                    f"手动拆分失败：\n{e}"
                )