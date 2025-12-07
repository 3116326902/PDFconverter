import tkinter as tk
from tkinter import ttk #tkinter 美化组件
from tkinter import messagebox  # 弹窗组件

# 定义主窗口类
class GUI:
    def __init__(self, root):
        # 初始化主窗口
        self.root = root
        self.root.title("PDF转换器")  # 窗口标题
        self.root.geometry("1080x720")  # 窗口大小
        self.root.minsize(720, 480) #规定最小窗口大小
        self.root.resizable(True, True) #允许窗口调整大小

        #配置内部元素比例
        self.root.grid_columnconfigure(0,weight=1)
        self.root.grid_columnconfigure(1,weight=14) #列比例

        self.root.grid_rowconfigure(0,weight=1)
        self.root.grid_rowconfigure(1,weight=9) #行比例

        #左侧边栏
        left_frame = ttk.Frame(root, style="Left.TFrame") #创建左侧边栏
        left_frame.grid(row=1, column=0, sticky="nsew")  #界面填充剩余空间

        #中间界面
        middle_frame = ttk.Frame(root, relief="raised", style="Middle.TFrame") #创建中间界面
        middle_frame.grid(row=1, column=1, sticky="nsew")  #界面填充剩余空间
        #中间边框圆角

        #顶部界面
        top_frame = ttk.Frame(root, style="Top.TFrame") #创建顶部界面
        top_frame.grid(row=0, column=0,columnspan=2, sticky="nsew")  #界面填充剩余空间

        #设置frame样式
        style = ttk.Style()
        style.configure("Left.TFrame", background="#D3D3D3")
        style.configure("Top.TFrame", background="#D3D3D3")
        style.configure("Middle.TFrame", background="#FFFFFF")

        #添加标签
        #左侧
        left_label = ttk.Label(left_frame, text="功能选择", font=("微软雅黑", 14))

        #中间
        middle_label = ttk.Label(middle_frame, text="最近文档", font=("微软雅黑", 14))


# 程序入口
if __name__ == "__main__":
    root = tk.Tk()  # 创建主窗口对象
    app = GUI(root)  # 创建GUI
    root.mainloop()  # 启动主循环