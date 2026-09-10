"""仅使用临时模拟班级，绝不加载或改写真实 JSON。"""
import json
import tempfile
from pathlib import Path
from datetime import datetime, timedelta, date
from unittest.mock import patch
import tkinter as tk
from tkinter import ttk
import seat_manager
from discipline import discipline_score
from gradebook import commit_exam
from attendance import set_leave


def children(node):
    for child in node.winfo_children():
        yield child
        yield from children(child)


def run():
    with tempfile.TemporaryDirectory() as temp:
        seat_manager.FILE=str(Path(temp)/"seat_data.json")
        data={"students":{},"seats":{},"speech":{},"sleep":{},"homework":{}}
        index=0
        for zone,cols,rows in seat_manager.SEAT_LAYOUT:
            for r in range(1,rows+1):
                for c in range(1,cols+1):
                    index+=1; name=f"测试学生{index:02}"
                    data["students"][name]={"姓名":name,"性别":"男" if index%2 else "女","学号":str(index)}
                    data["seats"][f"{zone}{r}-{c}"]=name
                    data["speech"][name]={date.today().isoformat():(index%5)*10}
        set_leave(data,"测试学生02",(datetime.now()+timedelta(minutes=1)).isoformat())
        draft={"source":"模拟.xlsx","sheet":"Sheet","sha256":"simulation","records":{
            "测试学生01":{"语文.score":95,"数学.score":70,"语文.rank":5,"数学.rank":80},
            "测试学生02":{"语文.score":70,"数学.score":95,"语文.rank":80,"数学.rank":5}}}
        commit_exam(data,draft,"模拟月考")
        seat_manager.save_data(data)
        app=seat_manager.App(); app.update()
        assert len(app.buttons)==80
        for button in app.buttons.values():
            assert button.winfo_width()>20 and button.winfo_height()>20
            assert button.cget("highlightbackground")=="black"
        assert app.buttons["左1-2"].cget("state")=="disabled"
        assert "纪律分" in app.buttons["左1-1"].cget("text")
        before=len([w for w in app.winfo_children() if isinstance(w,tk.Toplevel)])
        app.open_student("左1-2")
        assert before==len([w for w in app.winfo_children() if isinstance(w,tk.Toplevel)])
        app.drag_start("左1-2"); assert app.drag_seat is None
        app.open_student("左1-1"); app.update()
        detail=next(w for w in app.winfo_children() if isinstance(w,tk.Toplevel))
        question=next(w for w in children(detail) if isinstance(w,ttk.Button) and str(w.cget("text")).startswith("问问题"))
        before=discipline_score(app.data,"测试学生01")
        question.invoke(); app.update()
        assert discipline_score(app.data,"测试学生01")==before+.5
        assert "disabled" in question.state()
        water=next(w for w in children(detail) if isinstance(w,ttk.Button) and str(w.cget("text")).startswith("搬水"))
        water.invoke(); app.update()
        assert discipline_score(app.data,"测试学生01")==round(before+.7,2)
        assert f"纪律分：{before+.7:g}" in app.buttons["左1-1"].cget("text")
        detail.destroy(); app.update()
        for fn in (app.discipline_ranking, app.manage_attendance, app.choose_grade_student,
                   lambda: app.grade_details("测试学生01"), app.preview_complementary):
            fn(); app.update()
            for win in list(app.winfo_children()):
                if isinstance(win,tk.Toplevel): win.destroy()
        app.data["students"]["测试学生02"]["请假截止"]=(datetime.now()-timedelta(seconds=1)).isoformat()
        app.leave_tick(); app.update()
        assert app.buttons["左1-2"].cget("state")=="normal"
        assert app.data["students"]["测试学生02"]["状态"]=="在校中"
        from PIL import ImageGrab
        app.update_idletasks()
        ImageGrab.grab(bbox=(app.winfo_rootx(),app.winfo_rooty(),app.winfo_rootx()+app.winfo_width(),app.winfo_rooty()+app.winfo_height())).save("upgrade_ui_check.jpg")
        app.destroy()
        restored=json.loads(Path(seat_manager.FILE).read_text(encoding="utf8"))
        assert "exams" in restored and restored["activities"]["测试学生01"]
        print("CLASSROOM_UI_PASS: 80 seats, borders, score refresh, disabled leave, question limit, new dialogs, expiry, persistence")


if __name__=="__main__":run()
