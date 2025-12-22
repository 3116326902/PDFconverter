import sys
import os
from pathlib import Path

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QFrame, QLabel, QPushButton, QProgressBar, QVBoxLayout,
    QHBoxLayout, QGridLayout, QFileDialog, QMessageBox, QListWidget, QListWidgetItem
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QFont, QPixmap

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
        cv = Converter(self.input_file)
        pdf_doc = fitz.open(self.input_file)
        total_pages = len(pdf_doc)

        # 分步转换（显示进度）
        cv.convert(self.Word_output_file, start=0, end=None)
        self.progress_update.emit(100)
        cv.close()
        pdf_doc.close()

    def pdf_to_excel(self):
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
        top_frame = QFrame()
        top_frame.setStyleSheet("background-color: #3c3f41")
        parent_layout.addWidget(top_frame, 0, 0, 1, 2)
        top_layout = QVBoxLayout(top_frame)
        img_text_layout = QHBoxLayout()

        #加载图片
        img_label = QLabel(top_frame) #指定父控件
        img_label.setFixedSize(50, 50) #加载图片大小
        img = QPixmap("PDFconverter.ico")
        img = img.scaled(50, 50, Qt.AspectRatioMode.IgnoreAspectRatio, Qt.TransformationMode.SmoothTransformation) #设置图片大小,解除比例锁定
        img_label.setPixmap(img)
        img_label.setStyleSheet("border:0.5px solid #ffffff")

        #顶部文字
        title_label = QLabel("PDF转换器")
        title_font = QFont("微软雅黑", 14, QFont.Weight.Bold)
        title_label.setFont(title_font)

        img_text_layout.addWidget(img_label)
        img_text_layout.addWidget(title_label)

        v_layout = QVBoxLayout() #控制上下距离
        v_layout.addSpacing(5)
        h_layout = QHBoxLayout()
        h_layout.addLayout(img_text_layout) #控制左右距离
        v_layout.addLayout(h_layout)
        v_layout.addStretch(1) #底部拉伸

        top_layout.addLayout(v_layout)


    def create_left_frame(self, parent_layout):
        """左侧功能栏"""
        left_frame = QFrame()
        left_frame.setStyleSheet("background-color: #3c3f41")
        parent_layout.addWidget(left_frame, 1, 0)

        left_layout = QVBoxLayout(left_frame)
        left_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        left_layout.setContentsMargins(10, 20, 10, 10)
        left_layout.setSpacing(10)

        # 功能标题
        func_label = QLabel("功能选择")
        func_font = QFont("微软雅黑", 14, QFont.Weight.Bold)
        func_label.setFont(func_font)
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
        # PDF转Word
        self.pdf2word_btn = QPushButton("PDF转Word")
        self.pdf2word_btn.setStyleSheet(btn_style)
        self.pdf2word_btn.clicked.connect(lambda: self.switch_to_select_func("pdf2word"))
        self.pdf2word_btn.setEnabled(CONVERSION_ENABLED)
        left_layout.addWidget(self.pdf2word_btn)

        # PDF转Excel
        self.pdf2excel_btn = QPushButton("PDF转Excel")
        self.pdf2excel_btn.setStyleSheet(btn_style)
        self.pdf2excel_btn.clicked.connect(lambda: self.switch_to_select_func("pdf2excel"))
        self.pdf2excel_btn.setEnabled(CONVERSION_ENABLED)
        left_layout.addWidget(self.pdf2excel_btn)

        # Word转PDF
        self.word2pdf_btn = QPushButton("Word转PDF")
        self.word2pdf_btn.setStyleSheet(btn_style)
        self.word2pdf_btn.clicked.connect(lambda: self.switch_to_select_func("word2pdf"))
        self.word2pdf_btn.setEnabled(CONVERSION_ENABLED)
        left_layout.addWidget(self.word2pdf_btn)

        # Excel转PDF
        self.excel2pdf_btn = QPushButton("Excel转PDF")
        self.excel2pdf_btn.setStyleSheet(btn_style)
        self.excel2pdf_btn.clicked.connect(lambda: self.switch_to_select_func("excel2pdf"))
        self.excel2pdf_btn.setEnabled(CONVERSION_ENABLED)
        left_layout.addWidget(self.excel2pdf_btn)

        #底部拉伸，按钮置顶
        left_layout.addStretch()

    def create_middle_frame(self, parent_layout):
        """中间主界面"""
        middle_frame = QFrame()
        middle_frame.setStyleSheet("background-color: #2b2d30")
        parent_layout.addWidget(middle_frame, 1, 1)

        middle_layout = QVBoxLayout(middle_frame)
        middle_layout.setContentsMargins(20, 20, 20, 20)
        middle_layout.setSpacing(20)

        # 最近文档标题
        recent_label = QLabel("最近文档")
        recent_font = QFont("微软雅黑", 14, QFont.Weight.Bold)
        recent_label.setFont(recent_font)
        recent_label.setStyleSheet("background-color: #2b2d30")
        middle_layout.addWidget(recent_label)


    #跳转新窗口
    def switch_to_select_func(self, conversion_type):
        # 先检查新窗口是否已创建
        if hasattr(self, 'selectfunc') and self.selectfunc.isVisible():
            # 若已创建，直接激活并置顶
            self.selectfunc.activateWindow()
            self.selectfunc.raise_()
        else:
            # 若未创建，先打开新窗口
            self.selectfunc= SelectFunc(conversion_type, self)
            self.selectfunc.setParent(self)
            self.selectfunc.show()




class SelectFunc(QMainWindow):
    def __init__(self, conversion_type, main_windows):
        super().__init__()
        self.conversion_type = conversion_type
        self.main_window = main_windows  # 保存主窗口引用
        self.file_paths = []  # 修改：从单个文件路径改为列表，存储多选文件
        self.drag_pos = None  # 初始化拖动位置变量
        self.setStyleSheet("""
        QMainWindow {
            background-color: #1e1e1e
        }
        QWidget {
            color: #F8FAFC;
        }
        QListWidget {
            background-color: #2c2f31;
            border: 1px solid #444;
            border-radius: 5px;
            font-size: 14px;
            padding: 5px;
        }
        QListWidget::item {
            padding: 8px;
            border-bottom: 1px solid #3c3f41;
        }
        QListWidget::item:selected {
            background-color: #2196F3;
            color: white;
        }
    """)
        # 先初始化UI，再居中（否则获取不到窗口正确尺寸）
        self.init_ui()
        self.move_to_main_window_center()

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

    def move_to_main_window_center(self):
        # 获取主窗口的几何信息（位置+大小）
        main_geo = self.main_window.geometry()
        # 获取新窗口的大小
        self_geo = self.geometry()

        # 计算新窗口居中位置：主窗口中心 - 新窗口半宽/半高
        center_x = main_geo.x() + (main_geo.width() - self_geo.width()) // 2
        center_y = main_geo.y() + (main_geo.height() - self_geo.height()) // 2

        # 应用位置（仅改位置，不改大小）
        self.move(center_x, center_y)

    def init_ui(self):
        # 中心部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        # 主垂直布局
        layout = QVBoxLayout(central_widget)

        # 子网格布局
        second_layout = QGridLayout()
        second_layout.setSpacing(10)
        second_layout.setContentsMargins(10, 10, 10, 10)

        second_layout.setRowStretch(0, 1)
        second_layout.setRowStretch(1, 9)
        second_layout.setRowStretch(2, 3)

        # 将grid布局添加到主垂直布局
        layout.addLayout(second_layout)

        self.create_top_frame(second_layout)
        self.create_middle_frame(second_layout)
        self.create_bottom_frame(second_layout)

    def create_top_frame(self, parent_layout):
        top_frame = QFrame()
        top_frame.setStyleSheet("background-color: #3c3f41")
        parent_layout.addWidget(top_frame, 0, 0, 1, 2)
        top_layout = QVBoxLayout(top_frame)
        # 页面内容：根据转换类型动态显示标题
        title_text = f"✨ {self.get_conversion_title()}"
        title = QLabel(title_text)
        title.setStyleSheet("font-size: 20px; color: #2E86AB;")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        top_layout.addWidget(title)

    def create_middle_frame(self, parent_layout):
        middle_frame = QFrame()
        middle_frame.setStyleSheet("background-color: #3c3f41")
        # 跨列显示，避免布局错乱
        parent_layout.addWidget(middle_frame, 1, 0, 1, 2)

        middle_layout = QVBoxLayout(middle_frame)
        middle_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        middle_layout.setContentsMargins(10, 20, 10, 10)
        middle_layout.setSpacing(10)

        # 添加提示标签
        top_tool_layout = QHBoxLayout()
        list_tip_label = QLabel("已选择的文件：")
        list_tip_label.setStyleSheet("""
                font-size: 14px; 
                color: #2E86AB; 
                font-weight: bold;
            """)
        top_tool_layout.addWidget(list_tip_label)
        top_tool_layout.addStretch()  # 实现按钮右对齐

        # 删除按钮
        delete_btn = QPushButton("删除选中文件")
        delete_btn.clicked.connect(self.delete_selected_file)  # 绑定删除事件
        top_tool_layout.addWidget(delete_btn)
        middle_layout.addLayout(top_tool_layout)


        # 创建QListWidget用于展示多选文件列表
        self.file_list_widget = QListWidget()
        self.file_list_widget.setMinimumHeight(200)  # 设置最小高度，保证显示区域
        self.file_list_widget.setSelectionMode(QListWidget.SelectionMode.MultiSelection)# 支持批量选择


        delete_btn.setStyleSheet("""
            QPushButton {
                padding: 8px 16px;
                font-size: 14px;
                background-color: #f44336;
                color: white;
                border: none;
                border-radius: 6px;
            }
            QPushButton:hover {
                background-color: #d32f2f;
            }
            QPushButton:pressed {
                background-color: #b71c1c;
            }
        """)

        middle_layout.addWidget(self.file_list_widget)

    def create_bottom_frame(self, parent_layout):
        bottom_frame = QFrame()
        bottom_frame.setStyleSheet("background-color: #3c3f41")
        parent_layout.addWidget(bottom_frame, 2, 0, 1, 2)
        bottom_layout = QHBoxLayout(bottom_frame)
        bottom_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        bottom_layout.setContentsMargins(50, 20, 50, 10)
        bottom_layout.setSpacing(30)

        # 选择文件按钮
        select_btn = QPushButton("选择文件")
        select_btn.setStyleSheet("""
            QPushButton {
                padding: 10px 20px;
                font-size: 16px;
                background-color: #4CAF50;
                color: white;
                border: none;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        select_btn.clicked.connect(lambda: self.select_file(self.conversion_type))

        # 转换按钮
        converter_btn = QPushButton("开始转换")
        converter_btn.setStyleSheet("""
            QPushButton {
                padding: 10px 20px;
                font-size: 16px;
                background-color: #4CAF50;
                color: white;
                border: none;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        converter_btn.clicked.connect(lambda: self.converter_func(self.conversion_type))

        # 返回主窗口按钮
        back_btn = QPushButton("返回主窗口")
        back_btn.setStyleSheet("""
            QPushButton {
                padding: 10px 20px;
                font-size: 16px;
                background-color: #2196F3;
                color: white;
                border: none;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #0b7dda;
            }
        """)
        back_btn.clicked.connect(self.back_to_main)

        # 将按钮添加到底部布局
        bottom_layout.addWidget(select_btn)
        bottom_layout.addWidget(converter_btn)
        bottom_layout.addWidget(back_btn)

    # 根据转换类型获取标题
    def get_conversion_title(self):
        title_map = {
            "pdf2word": "PDF转Word 转换界面",
            "pdf2excel": "PDF转Excel 转换界面",
            "word2pdf": "word转PDF 转换界面",
            "excel2pdf": "excel转PDF 转换界面"
        }
        return title_map.get(self.conversion_type, "PDF转换界面")

    # 实现窗口拖动
    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.drag_pos = event.globalPosition().toPoint() - self.frameGeometry().topLeft()
            event.accept()

    def mouseMoveEvent(self, event):
        if event.buttons() == Qt.MouseButton.LeftButton and self.drag_pos is not None:
            self.move(event.globalPosition().toPoint() - self.drag_pos)
            event.accept()

    def select_file(self, conversion_type):
        if not conversion_type:
            QMessageBox.information(self, "提示", "请先选择左侧的转换类型（PDF转Word/Excel/图片）")
            return
        if conversion_type in ["pdf2word", "pdf2excel", "pdf2img"]:
            # PDF转其他格式：仅筛选PDF文件
            file_paths, _ = QFileDialog.getOpenFileNames(
                self,
                "选择PDF文件（可多选）",
                "",
                "PDF文件 (*.pdf);;所有文件 (*.*)"
            )
        elif conversion_type == "word2pdf":
            # Word转PDF：筛选docx/doc格式（新版+旧版Word文件）
            file_paths, _ = QFileDialog.getOpenFileNames(
                self,
                "选择Word文件（可多选）",
                "",
                "Word文件 (*.docx *.doc);;所有文件 (*.*)"
            )
        elif conversion_type == "excel2pdf":
            # 扩展：Excel转PDF：筛选xlsx/xls格式（新版+旧版Excel文件）
            file_paths, _ = QFileDialog.getOpenFileNames(
                self,
                "选择Excel文件（可多选）",
                "",
                "Excel文件 (*.xlsx *.xls);;所有文件 (*.*)"
            )
        elif conversion_type == "img2pdf":
            # 扩展：图片转PDF：筛选常见图片格式
            file_paths, _ = QFileDialog.getOpenFileNames(
                self,
                "选择图片文件（可多选）",
                "",
                "图片文件 (*.png *.jpg *.jpeg *.bmp);;所有文件 (*.*)"
            )


        if file_paths:
            self.file_paths = file_paths  # 保存多选文件路径到列表
            self.update_file_list_widget()  # 更新文件列表显示


    def update_file_list_widget(self):

        # 若有选中文件，逐个添加到列表
        if self.file_paths:
            for file_path in self.file_paths:
                # 获取文件名，同时显示完整路径可改为直接用file_path
                file_name = os.path.basename(file_path)
                list_item = QListWidgetItem(f"{file_name}")
                self.file_list_widget.addItem(list_item)
        else:
            # 若无选中文件，显示提示文字
            self.file_list_widget.addItem(QListWidgetItem("暂无选中文件"))

    def converter_func(self, conversion_type):
        # 先判断是否选择了文件
        if not self.file_paths:
            QMessageBox.warning(self, "警告", "请先选择要转换的PDF文件！")
            return

        if not conversion_type:
            QMessageBox.warning(self, "警告", "转换类型异常！")
            return

        # 批量处理每个选中的文件
        for file_path in self.file_paths:
            # 设置默认输出文件名
            file_name = os.path.basename(file_path)
            base_name = os.path.splitext(file_name)[0]

            if conversion_type == "pdf2word":
                output_file = f"{base_name}.docx"
                file_filter = "Word文件 (*.docx)"
            elif conversion_type == "pdf2excel":
                output_file = f"{base_name}.xlsx"
                file_filter = "Excel文件 (*.xlsx)"
            elif conversion_type == "pdf2img":
                output_file = f"{base_name}.png"
                file_filter = "图片文件 (*.png *.jpg)"
            else:
                QMessageBox.warning(self, "警告", f"不支持的转换类型：{conversion_type}")
                continue

            # 选择保存位置
            save_path = QFileDialog.getExistingDirectory(
                self,
                f"保存{conversion_type.replace('pdf2', '')}文件",
            )

            if save_path:
                self.start_conversion(conversion_type, file_path, save_path)



    def start_conversion(self, conversion_type, file_path, save_path):
        QMessageBox.information(self, "转换提示",
            f"正在转换：\n源文件：{os.path.basename(file_path)}\n目标文件：{os.path.basename(save_path)}\n转换类型：{conversion_type}")


        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        # 禁用按钮
        self.pdf2word_btn.setEnabled(False)
        self.pdf2excel_btn.setEnabled(False)
        self.pdf2img_btn.setEnabled(False)
        # 启动转换线程
        self.conversion_thread.conversion_type = conversion_type
        self.conversion_thread.input_file = file_path
        self.conversion_thread.output_file = save_path
        self.conversion_thread.start()



    def update_progress(self, value):
        """更新进度条"""
        self.progress_bar.setValue(value)



    def conversion_finished(self, success, message):
        # 显示结果
        if success:
            QMessageBox.information(self, "转换成功", message)
        else:
            QMessageBox.information(self, "转换失败", message)

        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(False)



    def delete_selected_file(self):

        # 获取选中项，无选中则提示
        selected_items = self.file_list_widget.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "提示", "请先选中要删除的文件！")
            return

        # 确认删除弹窗
        confirm = QMessageBox.question(
            self,
            "确认删除",
            f"是否确定删除选中的 {len(selected_items)} 个文件？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        if confirm != QMessageBox.StandardButton.Yes:
            return

        # 提取选中文件名，移除列表项
        selected_file_names = []
        for item in selected_items:
            file_name = item.text().replace("📄 ", "")
            selected_file_names.append(file_name)
            self.file_list_widget.takeItem(self.file_list_widget.row(item))

        # 同步更新self.file_paths数据
        new_file_paths = []
        for file_path in self.file_paths:
            base_name = os.path.basename(file_path)
            if base_name not in selected_file_names:
                new_file_paths.append(file_path)
        self.file_paths = new_file_paths

        QMessageBox.information(self, "成功", f"已成功删除 {len(selected_items)} 个文件！")

    # 返回主窗口的方法
    def back_to_main(self):
        self.close()


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