import tkinter as tk
from tkinter import messagebox  # 弹窗组件

# 定义主窗口类（面向对象写法，更易扩展）
class SimpleGUI:
    def __init__(self, root):
        # 初始化主窗口
        self.root = root
        self.root.title("简易信息查询工具")  # 窗口标题
        self.root.geometry("400x250")  # 窗口大小（宽x高）
        self.root.resizable(False, False)  # 禁止调整窗口大小

        # 1. 创建标签（提示文字）
        self.label = tk.Label(
            root, 
            text="请输入查询内容：", 
            font=("微软雅黑", 12)  # 字体和大小
        )
        self.label.pack(pady=10)  # 布局（垂直间距10像素）

        # 2. 创建输入框
        self.input_entry = tk.Entry(
            root, 
            font=("微软雅黑", 12), 
            width=30  # 输入框宽度
        )
        self.input_entry.pack(pady=5)

        # 3. 创建按钮框架（放两个按钮，横向排列）
        self.btn_frame = tk.Frame(root)
        self.btn_frame.pack(pady=10)

        # 查询按钮
        self.query_btn = tk.Button(
            self.btn_frame, 
            text="查询", 
            font=("微软雅黑", 10), 
            width=10,
            command=self.query_info  # 点击触发的函数
        )
        self.query_btn.grid(row=0, column=0, padx=5)  # 网格布局

        # 清空按钮
        self.clear_btn = tk.Button(
            self.btn_frame, 
            text="清空", 
            font=("微软雅黑", 10), 
            width=10,
            command=self.clear_input  # 点击触发的函数
        )
        self.clear_btn.grid(row=0, column=1, padx=5)

        # 4. 创建结果显示标签
        self.result_label = tk.Label(
            root, 
            text="查询结果：", 
            font=("微软雅黑", 12), 
            fg="blue"  # 文字颜色
        )
        self.result_label.pack(pady=10)

    # 定义查询功能
    def query_info(self):
        input_text = self.input_entry.get().strip()  # 获取输入框内容并去空格
        if not input_text:
            # 弹窗提示（警告）
            messagebox.showwarning("警告", "请输入查询内容！")
            return
        # 模拟查询逻辑（实际可替换为数据库/API查询）
        result = f"你查询的内容是：{input_text}\n查询时间：2025-12-07"
        self.result_label.config(text=f"查询结果：{result}")  # 更新标签内容

    # 定义清空功能
    def clear_input(self):
        self.input_entry.delete(0, tk.END)  # 清空输入框（从0到末尾）
        self.result_label.config(text="查询结果：")  # 重置结果标签
        messagebox.showinfo("提示", "已清空输入！")  # 弹窗提示（信息）

# 程序入口
if __name__ == "__main__":
    root = tk.Tk()  # 创建主窗口对象
    app = SimpleGUI(root)  # 实例化GUI类
    root.mainloop()  # 启动主循环（保持窗口显示）