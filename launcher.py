#!/usr/bin/env python3
"""学生管理工具主界面。"""
import os, subprocess, sys, tkinter as tk
from tkinter import messagebox

BASE = os.path.dirname(os.path.abspath(__file__))

def run_app(filename):
    path = os.path.join(BASE, filename)
    try:
        subprocess.Popen([sys.executable, path], cwd=BASE)
    except Exception as exc:
        messagebox.showerror("启动失败", str(exc))

root = tk.Tk()
root.title("学生管理系统")
root.geometry("620x430")
root.resizable(False, False)
tk.Label(root, text="学生管理系统", font=("Microsoft YaHei UI", 26, "bold"), pady=28).pack()
tk.Label(root, text="请选择要使用的功能", font=("Microsoft YaHei UI", 13), fg="#666").pack(pady=(0, 20))
box = tk.Frame(root); box.pack(fill="x", padx=90)
tk.Button(box, text="周末留宿管理系统", font=("Microsoft YaHei UI", 17, "bold"), height=3, bg="#4f9bd8", fg="white", activebackground="#357fbb", command=lambda: run_app("student_gui.py")).pack(fill="x", pady=10)
tk.Button(box, text="教室同学管理系统", font=("Microsoft YaHei UI", 17, "bold"), height=3, bg="#55b96b", fg="white", activebackground="#3f9a53", command=lambda: run_app("seat_manager.py")).pack(fill="x", pady=10)
tk.Label(root, text="所有数据均保存在本地，不需要联网", font=("Microsoft YaHei UI", 10), fg="#888").pack(side="bottom", pady=18)
root.mainloop()
