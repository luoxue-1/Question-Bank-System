import tkinter as tk
from gui import QuizApp

if __name__ == "__main__":
    # 创建主窗口
    root = tk.Tk()
    
    # 初始化应用
    app = QuizApp(root)
    
    # 启动主循环
    root.mainloop()
