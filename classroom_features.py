"""班级新功能的 Tkinter 表现层；业务计算放在独立模块以便回归验证。"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from copy import deepcopy
from datetime import date, datetime, timedelta
from discipline import discipline_score
from attendance import on_leave, set_leave, return_to_school, expire_leaves
from gradebook import (parse_xlsx, commit_exam, student_exams, comparisons,
                       profile_text, complementary_seats, SUBJECTS, field_label)


def table(parent, columns, height=15):
    body = ttk.Frame(parent); body.pack(fill="both", expand=True, padx=10, pady=8)
    tree = ttk.Treeview(body, columns=columns, show="headings", height=height)
    for column in columns:
        tree.heading(column, text=column); tree.column(column, width=115, minwidth=80, anchor="center")
    y = ttk.Scrollbar(body, orient="vertical", command=tree.yview)
    x = ttk.Scrollbar(body, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=y.set, xscrollcommand=x.set)
    tree.grid(row=0, column=0, sticky="nsew"); y.grid(row=0, column=1, sticky="ns"); x.grid(row=1, column=0, sticky="ew")
    body.rowconfigure(0, weight=1); body.columnconfigure(0, weight=1)
    return tree


class ClassroomFeatures:
    def commit_changes(self, pending, parent=None):
        from seat_manager import save_data
        try: save_data(pending)
        except OSError as exc:
            messagebox.showerror("保存失败", str(exc), parent=parent or self); return False
        self.data = pending
        self.refresh_summary()
        return True

    def build_feature_toolbar(self, parent):
        row = ttk.Frame(parent); row.pack(fill="x", pady=2)
        for text, callback in (("排名按钮", self.discipline_ranking), ("成绩导入 XLSX", self.import_grades),
                               ("成绩详情", self.choose_grade_student), ("互补排座", self.preview_complementary),
                               ("请假管理", self.manage_attendance)):
            ttk.Button(row, text=text, command=callback).pack(side="left", padx=3)
        ttk.Label(row, text="≤85 红 / ≤90 橙 / <95 黄 / ≥95 绿", font=("Microsoft YaHei UI", 9)).pack(side="left", padx=8)

    def discipline_ranking(self):
        win = tk.Toplevel(self); win.title("纪律分排名"); win.geometry("660x550")
        ttk.Label(win, text="按历史全部记录计算，从高到低；同分并列", padding=10).pack()
        tree = table(win, ("排名", "姓名", "学号", "纪律分", "状态"))
        def refresh():
            tree.delete(*tree.get_children())
            previous, rank = None, 0
            rows = sorted(self.data["students"], key=lambda n: (-discipline_score(self.data, n), n))
            for index, name in enumerate(rows, 1):
                score = discipline_score(self.data, name)
                if score != previous: rank = index
                previous = score; student = self.data["students"][name]
                tree.insert("", "end", values=(rank, name, student.get("学号", ""), f"{score:g}", "请假中" if on_leave(student) else "在校中"))
        ttk.Button(win, text="刷新排名", command=refresh).pack(pady=5); refresh()

    def leave_tick(self):
        today = date.today().isoformat()
        if today != self._clock_day:
            if self.current_day == self._clock_day:
                self.current_day = today; self.day_var.set(today); self.refresh_summary()
            self._clock_day = today
        expired = any(s.get("状态") in ("请假", "请假中") and not on_leave(s)
                      for s in self.data.get("students", {}).values())
        if expired:
            pending = deepcopy(self.data)
            expire_leaves(pending)
            if not self.commit_changes(pending):
                # 日期已到仍可根据真实期限操作；失败时下次继续尝试持久化。
                self.paint()
        for child in self.winfo_children():
            callback = getattr(child, "refresh_daily_actions", None)
            if callback: callback()
        self.after(1000, self.leave_tick)

    def record_leave(self, name, parent=None):
        until = simpledialog.askstring("请假", "请输入请假截止时间（YYYY-MM-DD HH:MM）\n时间到后自动恢复在校中；可在请假管理提前销假。",
                                       initialvalue=(datetime.now()+timedelta(hours=2)).strftime("%Y-%m-%d %H:%M"), parent=parent or self)
        if until is None: return
        pending = deepcopy(self.data)
        try: set_leave(pending, name, until.strip())
        except (ValueError, TypeError) as exc:
            messagebox.showerror("请假时间错误", str(exc), parent=parent or self); return
        if self.commit_changes(pending, parent) and parent and parent is not self:
            parent.destroy()

    def manage_attendance(self):
        win = tk.Toplevel(self); win.title("请假管理"); win.geometry("850x540")
        ttk.Label(win, text="请假中的座位不可点击。在这里登记期限或提前销假。", padding=10).pack()
        tree = table(win, ("姓名", "状态", "开始时间", "截止时间"))
        tree.column("开始时间", width=210); tree.column("截止时间", width=210)
        def refresh():
            if not win.winfo_exists(): return
            selected = tree.selection()
            tree.delete(*tree.get_children())
            for name, student in sorted(self.data["students"].items()):
                tree.insert("", "end", iid=name, values=(name, "请假中" if on_leave(student) else "在校中", student.get("请假开始", ""), student.get("请假截止", "")))
            for name in selected:
                if tree.exists(name): tree.selection_add(name)
        def action(leave):
            selected = tree.selection()
            if len(selected) != 1: messagebox.showinfo("请选择", "请先选择一名学生", parent=win); return
            name = selected[0]
            if leave: self.record_leave(name)
            elif on_leave(self.data["students"][name]) and messagebox.askyesno("提前销假", f"确认 {name} 已返校？", parent=win):
                pending = deepcopy(self.data); return_to_school(pending, name); self.commit_changes(pending, win)
            refresh()
        bar = ttk.Frame(win); bar.pack(pady=8)
        for text, cmd in (("请假/修改期限", lambda: action(True)), ("提前销假", lambda: action(False)), ("刷新状态", refresh)):
            ttk.Button(bar, text=text, command=cmd).pack(side="left", padx=5)
        refresh()

    def import_grades(self):
        filename = filedialog.askopenfilename(title="导入成绩（姓名匹配）", filetypes=[("Excel 工作簿", "*.xlsx")], parent=self)
        if not filename: return
        try: draft = parse_xlsx(filename, self.data["students"])
        except Exception as exc: messagebox.showerror("成绩导入失败", str(exc), parent=self); return
        title = simpledialog.askstring("考试名称", "确认本次考试名称（先后顺序按导入时间）：", initialvalue=draft["title"], parent=self)
        if title is None: return
        win = tk.Toplevel(self); win.title("成绩导入预览：确认后才保存"); win.geometry("960x650"); win.grab_set()
        note = f"考试：{title}；匹配 {len(draft['records'])} 人；未匹配 {len(draft['unmatched'])} 人（不新建学生、不修改学号）。"
        ttk.Label(win, text=note, wraplength=920, padding=10).pack(fill="x")
        info = tk.Text(win, height=5, wrap="word")
        info.insert("1.0", "列对应关系：\n"+"\n".join(draft["column_mapping"])+"\n未匹配姓名："+"、".join(draft["unmatched"])+"\n忽略列："+"、".join(draft["unrecognized"])+"\n班名次与校名次独立保存，没有校名次时不生成校名次变化。空白/缺考不当成 0 分。")
        info.configure(state="disabled"); info.pack(fill="x", padx=10)
        fields = [k for k in draft["recognized"] if k.endswith((".score", ".rank", ".class_rank"))]
        tree = table(win, ("姓名", *(field_label(k) for k in fields)), 13)
        for name, record in draft["records"].items():
            tree.insert("", "end", values=(name, *(record.get(k, "—") for k in fields)))
        def commit():
            pending = deepcopy(self.data)
            try: commit_exam(pending, draft, title)
            except ValueError as exc: messagebox.showerror("不能导入", str(exc), parent=win); return
            if self.commit_changes(pending, win):
                win.destroy(); messagebox.showinfo("导入成功", "成绩已按姓名归档，可以打开成绩详情查看。", parent=self)
        ttk.Button(win, text="确认导入本次考试", command=commit).pack(pady=10)

    def choose_grade_student(self):
        win = tk.Toplevel(self); win.title("选择学生查看成绩详情"); win.geometry("560x530")
        query = tk.StringVar(); ttk.Entry(win, textvariable=query).pack(fill="x", padx=10, pady=8)
        tree = table(win, ("姓名", "学号", "考试记录数"))
        def refresh(*_):
            tree.delete(*tree.get_children())
            for name, student in sorted(self.data["students"].items()):
                if query.get() not in name and query.get() not in str(student.get("学号", "")): continue
                tree.insert("", "end", iid=name, values=(name, student.get("学号", ""), len(student_exams(self.data, name))))
        def show(*_):
            if tree.selection(): self.grade_details(tree.selection()[0])
        query.trace_add("write", refresh); tree.bind("<Double-1>", show)
        ttk.Button(win, text="成绩详情", command=show).pack(pady=8); refresh()

    def grade_details(self, name):
        exams = student_exams(self.data, name)
        win = tk.Toplevel(self); win.title(f"成绩详情 · {name}"); win.geometry("940x580"); win.grab_set()
        if not exams:
            ttk.Label(win, text="暂无成绩，请先从主界面导入 XLSX。", padding=25).pack(); return
        choice = ttk.Combobox(win, state="readonly", values=[f"{e['title']} · 录入 {e['imported_at'][:19]}" for e in exams], width=80)
        choice.pack(fill="x", padx=10, pady=10); choice.current(len(exams)-1)
        headline = ttk.Label(win, wraplength=900, padding=10); headline.pack(fill="x")
        tree = table(win, ("科目", "本次分数", "本次班名次", "本次校名次", "较上次校名次", "分数变化"), 7)
        note = ttk.Label(win, wraplength=900, padding=10); note.pack(fill="x")
        ttk.Button(win, text="智能评语（预留，尚未启用）", state="disabled").pack(pady=8)
        def change(*_):
            index = choice.current(); exam = exams[index]; current = exam["records"][name]
            previous = exams[index-1]["records"][name] if index else {}
            profile, basis = profile_text(exam, name)
            headline.configure(text=profile+"\n依据："+basis+"；仅作学习搭配参考，不代表学科能力定论。")
            tree.delete(*tree.get_children())
            for row in comparisons(current, previous):
                delta = row["rank_change"]
                trend = "无法比较" if delta is None else (f"进步 {delta} 名" if delta > 0 else f"退步 {-delta} 名" if delta < 0 else "持平")
                tree.insert("", "end", values=(row["subject"], row["score"] if row["score"] is not None else "未提供", row["class_rank"] if row["class_rank"] is not None else "未提供", row["rank"] if row["rank"] is not None else "未提供", trend, "无法比较" if row["score_change"] is None else f"{row['score_change']:+g}"))
            text = "首次成绩记录，无上次可比较。" if not index else "对比："+exams[index-1]["title"]+"（该学生上一条已录入成绩）。名次越小越好；原始分数受试卷难度和满分影响。"
            if current.get("cohort") != previous.get("cohort") and index: text += " 参考人数有变化或未完整提供，名次变化请谨慎解读。"
            note.configure(text=text+"\n智能评语先预留：以后可结合三次以上趋势、进步稳定性、教师备注给出可核对的建议；本版不生成 AI 评价，也不联网传输成绩。")
        choice.bind("<<ComboboxSelected>>", change); change()

    def preview_complementary(self):
        from seat_manager import SEAT_LAYOUT
        try: seats, exam, count = complementary_seats(self.data, SEAT_LAYOUT)
        except ValueError as exc: messagebox.showinfo("暂不能排座", str(exc), parent=self); return
        win = tk.Toplevel(self); win.title("互补排座预览（未修改座次）"); win.geometry("980x650"); win.grab_set()
        ttk.Label(win, text=f"依据最近录入：{exam['title']}；建议互补 {count} 对。\n同区同排相邻两座为一对，不跨过道；无成绩学生仍保留座位。启发式建议不保证最优，也未考虑身高、视力或人际因素。", wraplength=930, padding=10).pack()
        tree = table(win, ("座位", "姓名", "学科参考"), 20); tree.column("学科参考", width=610)
        for zone, cols, rows in SEAT_LAYOUT:
            for row in range(1, rows+1):
                for col in range(1, cols+1):
                    seat = f"{zone}{row}-{col}"; name = seats[seat]
                    tree.insert("", "end", values=(seat, name or "空座", profile_text(exam, name)[0] if name else ""))
        def apply():
            if not messagebox.askyesno("替换当前座位", "确认使用此方案？将保存一份原座位记录，成绩与纪律记录仍跟随姓名。", parent=win): return
            pending = deepcopy(self.data)
            pending.setdefault("seat_history", []).append({"at": datetime.now().isoformat(), "seats": dict(pending["seats"]), "reason": "互补排座"})
            pending["seats"] = seats
            if self.commit_changes(pending, win): win.destroy()
        ttk.Button(win, text="确认应用排座方案", command=apply).pack(pady=10)

    def detail_body(self, win):
        """详情可滚动，活动按钮增加时不会被窗口底部裁掉。"""
        win.geometry(f"900x{min(760, max(480, self.winfo_screenheight()-100))}")
        host = ttk.Frame(win); host.pack(fill="both", expand=True)
        canvas = tk.Canvas(host, highlightthickness=0)
        scroll = ttk.Scrollbar(host, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y"); canvas.pack(side="left", fill="both", expand=True)
        body = tk.Frame(canvas); item = canvas.create_window((0, 0), window=body, anchor="nw")
        body.bind("<Configure>", lambda _: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(item, width=e.width))
        win.bind("<MouseWheel>", lambda e: canvas.yview_scroll(-int(e.delta/120), "units"))
        return body
