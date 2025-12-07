import tkinter as tk
from tkinter import ttk
import os

class PDFConverterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF转换器")
        self.root.geometry("1080x720")
        self.root.minsize(720, 480)
        self.root.resizable(True, True)

        # 初始化变量
        self.recent_files = []  # 最近文件列表
        self.selected_file = tk.StringVar()

        # 配置网格权重
        self._setup_grid()
        # 设置全局样式
        self._setup_styles()
        # 创建界面组件
        self._create_top_frame()
        self._create_left_frame()
        self._create_middle_frame()  # 简化的中间界面（无圆角）

    def _setup_grid(self):
        """配置主窗口网格权重"""
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_columnconfigure(1, weight=14)
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_rowconfigure(1, weight=9)

    def _setup_styles(self):
        """设置全局样式"""
        style = ttk.Style()
        # 配置Frame样式
        style.configure("Left.TFrame", background="#D3D3D3")
        style.configure("Top.TFrame", background="#D3D3D3")
        style.configure("Middle.TFrame", background="#FFFFFF")
        # 配置标签样式
        style.configure("Title.TLabel", font=("微软雅黑", 14), background="#D3D3D3")
        style.configure("Content.TLabel", font=("微软雅黑", 14), background="#FFFFFF")

    def _create_top_frame(self):
        """创建顶部界面"""
        top_frame = ttk.Frame(self.root, style="Top.TFrame")
        top_frame.grid(row=0, column=0, columnspan=2, sticky="nsew")
        # 顶部标题
        top_label = ttk.Label(top_frame, text="PDF转换器 - 多功能格式转换工具", style="Title.TLabel")
        top_label.pack(expand=True)

    def _create_left_frame(self):
        """创建左侧边栏"""
        left_frame = ttk.Frame(self.root, style="Left.TFrame")
        left_frame.grid(row=1, column=0, sticky="nsew")

        # 左侧标题
        left_label = ttk.Label(left_frame, text="功能选择", style="Title.TLabel")
        left_label.pack(pady=20)

        # 功能按钮
        btn_pdf2word = ttk.Button(left_frame, text="PDF转Word")
        btn_pdf2word.pack(pady=10, padx=10, fill=tk.X)

        btn_pdf2excel = ttk.Button(left_frame, text="PDF转Excel")
        btn_pdf2excel.pack(pady=10, padx=10, fill=tk.X)

        btn_pdf2img = ttk.Button(left_frame, text="PDF转图片")
        btn_pdf2img.pack(pady=10, padx=10, fill=tk.X)

    def _create_middle_frame(self):
        """创建无圆角的中间主界面"""
        # 直接创建原生Frame（移除所有画布相关逻辑）
        middle_frame = ttk.Frame(self.root, style="Middle.TFrame")
        middle_frame.grid(row=1, column=1, sticky="nsew", padx=10, pady=10)

        # 中间界面标题
        middle_title = ttk.Label(middle_frame, text="最近文档", style="Content.TLabel")
        middle_title.pack(pady=20)

        # 可扩展：添加最近文件列表、文件选择按钮等
        # 示例：添加一个文件选择按钮
        btn_select_file = ttk.Button(middle_frame, text="选择文件")
        btn_select_file.pack(pady=10)

        # 示例：添加空白列表框（用于显示最近文件）
        recent_list = tk.Listbox(middle_frame, font=("微软雅黑", 12), height=20)
        recent_list.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

# 程序入口
if __name__ == "__main__":
    root = tk.Tk()
    app = PDFConverterGUI(root)
    root.mainloop()