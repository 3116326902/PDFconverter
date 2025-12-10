import sys
import os
import json
import tempfile
from datetime import datetime
from pathlib import Path

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QFrame, QLabel, QPushButton,
    QListWidget, QProgressBar, QVBoxLayout, QHBoxLayout, QGridLayout,
    QFileDialog, QMessageBox
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QFont

# 尝试导入转换库，缺失时提供友好提示
try:
    from pdf2docx import Converter
    import pdfplumber
    import openpyxl
    from PIL import Image
    import fitz  # PyMuPDF
    CONVERSION_ENABLED = True
except ImportError as e:
    CONVERSION_ENABLED = False
    MISSING_MODULE = str(e).split("'")[1]  # 获取缺失的模块名

# 转换线程（避免UI卡顿）
class ConversionThread(QThread):
    progress_update = pyqtSignal(int)
    finished_signal = pyqtSignal(bool, str)
    conversion_type = ""
    input_file = ""
    PDF_output_file = ""
    Word_output_file = ""
    Excel_output_file = ""


    def run(self):
        try:
            if not CONVERSION_ENABLED:
                raise Exception(f"缺少转换依赖库，请先安装：{MISSING_MODULE}")

            if self.conversion_type == "pdf2word":
                self.pdf_to_word()
            elif self.conversion_type == "pdf2excel":
                self.pdf_to_excel()
            elif self.conversion_type == "pdf2img":
                self.pdf_to_image()
            self.finished_signal.emit(True, f"转换完成：\n")
        except Exception as e:
            self.finished_signal.emit(False, f"转换失败：\n{str(e)}")

    def pdf_to_word(self):
        """PDF转Word"""
        cv = Converter(self.input_file)
        pdf_doc = fitz.open(self.input_file)
        total_pages = len(pdf_doc)

        # 分步转换（显示进度）
        cv.convert(self.Word_output_file, start=0, end=None)
        self.progress_update.emit(100)
        cv.close()
        pdf_doc.close()

    def pdf_to_excel(self):
        """PDF转Excel"""
        workbook = openpyxl.Workbook()
        worksheet = workbook.active
        worksheet.title = "PDF内容"

        with pdfplumber.open(self.input_file) as pdf:
            total_pages = len(pdf.pages)
            row = 1
            for i, page in enumerate(pdf.pages):
                try:
                    text = page.extract_text()
                    if text:
                        for line in text.split('\n'):
                            worksheet.cell(row=row, column=1, value=line)
                            row += 1
                    # 更新进度
                    progress = int((i + 1) / total_pages * 100)
                    self.progress_update.emit(progress)
                except Exception as e:
                    self.progress_update.emit(int((i + 1) / total_pages * 100))
                    continue

        workbook.save(self.PDF_output_file)
        self.progress_update.emit(100)

    def pdf_to_image(self):
        """PDF转图片（高分辨率）"""
        pdf_document = fitz.open(self.input_file)
        total_pages = len(pdf_document)

        # 创建输出目录（多页PDF）
        if total_pages > 1:
            img_dir = Path(self.PDF_output_file).parent / Path(self.PDF_output_file).stem
            img_dir.mkdir(exist_ok=True)

        for i, page in enumerate(pdf_document):
            # 设置高分辨率（dpi=300）
            pix = page.get_pixmap(dpi=300)
            if total_pages > 1:
                img_path = str(img_dir / f"第{i+1}页.png")
            else:
                img_path = self.PDF_output_file

            pix.save(img_path)
            progress = int((i + 1) / total_pages * 100)
            self.progress_update.emit(progress)

        pdf_document.close()
        self.progress_update.emit(100)


class PDFConverterGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.init_ui()

        # 初始化转换线程
        if CONVERSION_ENABLED:
            self.conversion_thread = ConversionThread()
            self.conversion_thread.progress_update.connect(self.update_progress)
            self.conversion_thread.finished_signal.connect(self.conversion_finished)
        else:
            # 依赖缺失时提示
            QMessageBox.warning(
                self,
                "功能受限",
                f"缺少必要的转换库：{MISSING_MODULE}\n\n请执行以下命令安装：\n"
                f"pip install {MISSING_MODULE} -i https://pypi.tuna.tsinghua.edu.cn/simple"
            )

    def init_ui(self):
        # 主窗口设置
        self.setWindowTitle("PDF转换器 - 多功能格式转换工具")
        self.setGeometry(100, 100, 1080, 720)
        self.setMinimumSize(720, 480)

        # 中心窗口
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # 主布局
        main_layout = QGridLayout(central_widget)
        main_layout.setSpacing(10)
        main_layout.setContentsMargins(10, 10, 10, 10)

        # 配置网格权重
        main_layout.setColumnStretch(0, 1)
        main_layout.setColumnStretch(1, 14)
        main_layout.setRowStretch(0, 1)
        main_layout.setRowStretch(1, 9)

        # 创建组件
        self.create_top_frame(main_layout)
        self.create_left_frame(main_layout)
        self.create_middle_frame(main_layout)


    def create_top_frame(self, parent_layout):
        """顶部标题栏"""
        top_frame = QFrame()
        top_frame.setStyleSheet("background-color: #D3D3D3;")
        parent_layout.addWidget(top_frame, 0, 0, 1, 2)

        top_layout = QVBoxLayout(top_frame)
        top_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        title_label = QLabel("PDF转换器 - 多功能格式转换工具")
        title_font = QFont("微软雅黑", 14, QFont.Weight.Bold)
        title_label.setFont(title_font)
        title_label.setStyleSheet("background-color: #D3D3D3;")
        top_layout.addWidget(title_label)

    def create_left_frame(self, parent_layout):
        """左侧功能栏"""
        left_frame = QFrame()
        left_frame.setStyleSheet("background-color: #D3D3D3;")
        parent_layout.addWidget(left_frame, 1, 0)

        left_layout = QVBoxLayout(left_frame)
        left_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        left_layout.setContentsMargins(10, 20, 10, 10)
        left_layout.setSpacing(10)

        # 功能标题
        func_label = QLabel("功能选择")
        func_font = QFont("微软雅黑", 14, QFont.Weight.Bold)
        func_label.setFont(func_font)
        func_label.setStyleSheet("background-color: #D3D3D3;")
        func_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        left_layout.addWidget(func_label)
        left_layout.addSpacing(10)

        # 按钮样式
        btn_style = """
            QPushButton {
                font-family: 微软雅黑;
                font-size: 12px;
                padding: 8px;
                border-radius: 4px;
                background-color: #f0f0f0;
            }
            QPushButton:hover {
                background-color: #e0e0e0;
            }
            QPushButton:pressed {
                background-color: #d0d0d0;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
        """

        # 功能按钮（依赖缺失时禁用）
        self.pdf2word_btn = QPushButton("PDF转Word")
        self.pdf2word_btn.setStyleSheet(btn_style)
        self.pdf2word_btn.clicked.connect(lambda: self.select_file("pdf2word"))
        self.pdf2word_btn.setEnabled(CONVERSION_ENABLED)
        left_layout.addWidget(self.pdf2word_btn)

        self.pdf2excel_btn = QPushButton("PDF转Excel")
        self.pdf2excel_btn.setStyleSheet(btn_style)
        self.pdf2excel_btn.clicked.connect(lambda: self.select_file("pdf2excel"))
        self.pdf2excel_btn.setEnabled(CONVERSION_ENABLED)
        left_layout.addWidget(self.pdf2excel_btn)

        self.pdf2img_btn = QPushButton("PDF转图片")
        self.pdf2img_btn.setStyleSheet(btn_style)
        self.pdf2img_btn.clicked.connect(lambda: self.select_file("pdf2img"))
        self.pdf2img_btn.setEnabled(CONVERSION_ENABLED)
        left_layout.addWidget(self.pdf2img_btn)

        left_layout.addStretch()

    def create_middle_frame(self, parent_layout):
        """中间主界面"""
        middle_frame = QFrame()
        middle_frame.setStyleSheet("background-color: #FFFFFF;")
        parent_layout.addWidget(middle_frame, 1, 1)

        middle_layout = QVBoxLayout(middle_frame)
        middle_layout.setContentsMargins(20, 20, 20, 20)
        middle_layout.setSpacing(20)

        # 最近文档标题
        recent_label = QLabel("最近文档")
        recent_font = QFont("微软雅黑", 14, QFont.Weight.Bold)
        recent_label.setFont(recent_font)
        recent_label.setStyleSheet("background-color: #FFFFFF;")
        middle_layout.addWidget(recent_label)

        # 选择文件按钮
        select_btn = QPushButton("选择文件")
        select_btn.setStyleSheet("""
            QPushButton {
                font-family: 微软雅黑;
                font-size: 12px;
                padding: 10px;
                border-radius: 4px;
                background-color: #4a86e8;
                color: white;
            }
            QPushButton:hover {
                background-color: #3d72d6;
            }
            QPushButton:disabled {
                background-color: #88a4d8;
                color: #dddddd;
            }
        """)
        select_btn.clicked.connect(lambda: self.select_file(""))
        select_btn.setEnabled(CONVERSION_ENABLED)
        middle_layout.addWidget(select_btn)


        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(False)
        middle_layout.addWidget(self.progress_bar)

    def select_file(self, conversion_type):
        """选择PDF文件"""
        if not conversion_type:
            QMessageBox.information(self, "提示", "请先选择左侧的转换类型（PDF转Word/Excel/图片）")
            return

        # 选择PDF文件
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择PDF文件",
            "",
            "PDF文件 (*.pdf);;所有文件 (*.*)"
        )

        if not file_path:
            return


        # 设置默认输出文件名
        base_name = os.path.splitext(file_path)[0]
        if conversion_type == "pdf2word":
            output_file = f"{base_name}.docx"
            file_filter = "Word文件 (*.docx)"
        elif conversion_type == "pdf2excel":
            output_file = f"{base_name}.xlsx"
            file_filter = "Excel文件 (*.xlsx)"
        elif conversion_type == "pdf2img":
            output_file = f"{base_name}.png"
            file_filter = "图片文件 (*.png)"
        else:
            return

        # 选择保存位置
        save_path, _ = QFileDialog.getSaveFileName(
            self,
            f"保存{conversion_type.replace('pdf2', '')}文件",
            output_file,
            file_filter
        )

        if save_path:
            self.start_conversion(conversion_type, file_path, save_path)

    def start_conversion(self, conversion_type, input_file, output_file):
        """启动转换"""
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)

        # 禁用按钮
        self.pdf2word_btn.setEnabled(False)
        self.pdf2excel_btn.setEnabled(False)
        self.pdf2img_btn.setEnabled(False)

        # 启动转换线程
        self.conversion_thread.conversion_type = conversion_type
        self.conversion_thread.input_file = input_file
        self.conversion_thread.output_file = output_file
        self.conversion_thread.start()

    def update_progress(self, value):
        """更新进度条"""
        self.progress_bar.setValue(value)

    def conversion_finished(self, success, message):
        """转换完成回调"""
        # 启用按钮
        self.pdf2word_btn.setEnabled(True)
        self.pdf2excel_btn.setEnabled(True)
        self.pdf2img_btn.setEnabled(True)

        # 显示结果
        if success:
            QMessageBox.information(self, "转换成功", message)
        else:
            QMessageBox.critical(self, "转换失败", message)

        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(False)

        # 限制最多10条





    def select_file_with_path(self, conversion_type, file_path):
        """使用已有路径转换"""
        base_name = os.path.splitext(file_path)[0]
        if conversion_type == "pdf2word":
            output_file = f"{base_name}.docx"
            file_filter = "Word文件 (*.docx)"
        elif conversion_type == "pdf2excel":
            output_file = f"{base_name}.xlsx"
            file_filter = "Excel文件 (*.xlsx)"
        elif conversion_type == "pdf2img":
            output_file = f"{base_name}.png"
            file_filter = "图片文件 (*.png)"
        else:
            return

        save_path, _ = QFileDialog.getSaveFileName(
            self,
            f"保存{conversion_type.replace('pdf2', '')}文件",
            output_file,
            file_filter
        )

        if save_path:
            self.start_conversion(conversion_type, file_path, save_path)


if __name__ == "__main__":
    # ===== 修复核心：移除PyQt6中不存在的高分屏属性 =====
    # PyQt6 已默认启用高分屏缩放，无需手动设置
    app = QApplication(sys.argv)

    # 设置全局字体
    font = QFont("微软雅黑", 10)
    app.setFont(font)

    # 启动主窗口
    window = PDFConverterGUI()
    window.show()

    sys.exit(app.exec())