#!/usr/bin/env python3
"""学生管理工具主界面。"""
import os, subprocess, sys, tkinter as tk
from tkinter import messagebox

BASE = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))

def run_app(mode, root=None):
    try:
        command = [sys.executable, f"--{mode}"] if getattr(sys, "frozen", False) else [sys.executable, os.path.abspath(__file__), f"--{mode}"]
        subprocess.Popen(command, cwd=BASE)
        if root is not None and mode == "classroom": root.destroy()
    except Exception as exc:
        messagebox.showerror("启动失败", str(exc))

def main():
    if "--weekend" in sys.argv:
        from student_gui import App
        App().mainloop(); return
    if "--classroom" in sys.argv:
        from seat_manager import App
        App().mainloop(); return
    root = tk.Tk(); root.title("学生管理系统"); root.geometry("620x430"); root.resizable(False, False)
    tk.Label(root, text="学生管理系统", font=("Microsoft YaHei UI", 26, "bold"), pady=28).pack()
    tk.Label(root, text="请选择要使用的功能", font=("Microsoft YaHei UI", 13), fg="#666").pack(pady=(0, 20))
    box = tk.Frame(root); box.pack(fill="x", padx=90)
    tk.Button(box, text="周末留宿管理系统", font=("Microsoft YaHei UI", 17, "bold"), height=3, bg="#4f9bd8", fg="white", activebackground="#357fbb", command=lambda: run_app("weekend")).pack(fill="x", pady=10)
    tk.Button(box, text="教室同学管理系统", font=("Microsoft YaHei UI", 17, "bold"), height=3, bg="#55b96b", fg="white", activebackground="#3f9a53", command=lambda: run_app("classroom", root)).pack(fill="x", pady=10)
    tk.Label(root, text="所有数据均保存在本地，不需要联网", font=("Microsoft YaHei UI", 10), fg="#888").pack(side="bottom", pady=18)
    root.mainloop()

if __name__ == "__main__":
    main()
