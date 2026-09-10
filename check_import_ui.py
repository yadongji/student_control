import tempfile
from pathlib import Path
from unittest.mock import patch
import tkinter as tk
from tkinter import ttk
from openpyxl import Workbook
import seat_manager


def descendants(node):
    for child in node.winfo_children():
        yield child
        yield from descendants(child)


with tempfile.TemporaryDirectory() as temp:
    seat_manager.FILE = str(Path(temp)/"seat_data.json")
    seat_manager.save_data({"students": {"测试学生": {}}, "seats": {"左1-1": "测试学生"}})
    file = Path(temp)/"模拟月考.xlsx"
    wb = Workbook(); ws = wb.active
    ws.append(["姓名", "物理成绩", "班次", "校次", "化学", "班排名", "校排名"])
    ws.append(["测试学生", 80, 2, 18, 90, 1, 8]); wb.save(file); wb.close()
    app = seat_manager.App(); app.withdraw()
    try:
        with patch('classroom_features.filedialog.askopenfilename', return_value=str(file)), patch('classroom_features.simpledialog.askstring', return_value="模拟月考"), patch('classroom_features.messagebox.showinfo'), patch('classroom_features.messagebox.showerror') as error:
            app.import_grades(); app.update()
            preview = next(w for w in app.winfo_children() if isinstance(w, tk.Toplevel))
            tree = next(w for w in descendants(preview) if isinstance(w, ttk.Treeview))
            assert "物理班名次" in tree["columns"] and "化学校名次" in tree["columns"]
            button = next(w for w in descendants(preview) if isinstance(w, ttk.Button) and w.cget("text") == "确认导入本次考试")
            button.invoke(); app.update(); error.assert_not_called()
        app.grade_details("测试学生"); app.update()
        details = next(w for w in app.winfo_children() if isinstance(w, tk.Toplevel))
        tree = next(w for w in descendants(details) if isinstance(w, ttk.Treeview))
        physics = next(tree.item(i)["values"] for i in tree.get_children() if tree.item(i)["values"][0] == "物理")
        assert list(map(str, physics[1:4])) == ["80", "2", "18"]
        print("IMPORT_UI_PASS: repeated headers, preview values, persistence, separate class/school ranks in details; real data untouched")
    finally: app.destroy()
