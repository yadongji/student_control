"""隔离数据的界面检查，不读写实际学生记录。"""
import json
import os
import tempfile
import tkinter as tk
from unittest.mock import patch
import seat_manager as sm


def descendants(widget):
    for child in widget.winfo_children():
        yield child
        yield from descendants(child)


with tempfile.TemporaryDirectory() as folder:
    sm.FILE=os.path.join(folder,"seat_data.json")
    data={"students":{"测试学生":{"姓名":"测试学生"}},"seats":{"左1-1":"测试学生"},"speech":{},"sleep":{},"homework":{}}
    with patch.object(sm,"load_data",return_value=data):
        app=sm.App()
    try:
        app.withdraw()
        app.open_student("左1-1"); app.update()
        win=next(w for w in app.winfo_children() if isinstance(w,tk.Toplevel))
        for text in ("做板报一次","做板报一次","讲题一次","未打扫卫生一次"):
            next(w for w in descendants(win) if "text" in w.keys() and w.cget("text")==text).invoke()
        app.update()
        assert sm.discipline_score(app.data,"测试学生")==102
        with open(sm.FILE,encoding="utf-8") as f: stored=json.load(f)
        assert len(stored["activities"]["测试学生"])==4
        assert any("纪律评分：102" in str(w.cget("text")) for w in descendants(win) if "text" in w.keys())
        app.show_activity_history("测试学生",win); app.update()
        assert any(isinstance(w,sm.ttk.Treeview) and len(w.get_children())==4 for w in descendants(win))
        print("Activity UI, persistence and score refresh passed; real data untouched.")
    finally:
        app.destroy()
