#!/usr/bin/env python3
"""离线学生信息与周末留校登记系统（Tkinter 图形版）。"""
from __future__ import annotations

import json
import os
import re
import sys
import tkinter as tk
from datetime import date, timedelta
from tkinter import filedialog, messagebox, ttk, simpledialog

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Font, PatternFill
except ImportError:
    Workbook = load_workbook = None

BASE_DIR = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(BASE_DIR, "student_data.json")
STUDENT_FIELDS = ["学号", "姓名", "住宿情况", "性别", "宿舍号", "床位号", "班级"]
STAY_OPTIONS = ["走读留铺", "走读无铺", "住宿"]
PERIODS = ["周五晚留校", "周六上午留校", "周六下午留校", "周六晚上留校", "周日留校", "周五住宿", "周六住宿"]
PERIOD_MAP = {"周五晚留校": ["周五晚留校"], "周六上午留校": ["周六上午留校"], "周六下午留校": ["周六下午留校"], "周六晚上留校": ["周六晚上留校", "周六晚上留宿"], "周日留校": ["周日留校"], "周五住宿": ["周五晚留宿"], "周六住宿": ["周六晚上留宿"]}
ALIASES = {
    "学号": ["学号", "学生学号", "编号", "学生编号"],
    "姓名": ["姓名", "学生姓名", "名字"],
    "住宿情况": ["住宿情况", "住宿类型", "住宿状态", "是否住宿"],
    "性别": ["性别"], "宿舍号": ["宿舍号", "寝室号", "宿舍", "寝室"],
    "班级": ["班级", "行政班", "班别"],
}


def clean(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value).strip()


def yes_blank(value):
    return "是" if int(value or 0) else ""


def week_key(day=None):
    day = day or date.today()
    monday = day - timedelta(days=day.weekday())
    return monday.isoformat()


def week_title(key):
    start = date.fromisoformat(key)
    return f"{start:%Y-%m-%d} 至 {start + timedelta(days=6):%Y-%m-%d}"


def fill_class_from_id(student):
    """班级为空时，从学号中提取三位班级编码（如 2026107... -> 107）。"""
    sid = clean(student.get("学号"))
    if not clean(student.get("班级")) and len(sid) >= 7 and sid[:7].isdigit():
        student["班级"] = sid[4:7]
    return student


class Store:
    def __init__(self):
        self.data = {"students": [], "weeks": {}}
        self.load()

    def load(self):
        if os.path.exists(DATA_FILE):
            try:
                with open(DATA_FILE, encoding="utf-8") as f:
                    raw = json.load(f)
                if isinstance(raw, dict):
                    self.data.update(raw)
            except (OSError, json.JSONDecodeError) as exc:
                messagebox.showwarning("读取失败", f"数据文件无法读取：{exc}")
        self.data.setdefault("students", [])
        self.data.setdefault("weeks", {})

    def save(self):
        temp = DATA_FILE + ".tmp"
        with open(temp, "w", encoding="utf-8") as f:
            json.dump(self.data, f, ensure_ascii=False, indent=2)
        os.replace(temp, DATA_FILE)

    def student(self, sid):
        return next((s for s in self.data["students"] if s["学号"] == sid), None)

    def record(self, week, sid):
        records = self.data["weeks"].setdefault(week, {})
        return records.setdefault(sid, {p: 0 for p in sum(PERIOD_MAP.values(), [])})


class StudentDialog(tk.Toplevel):
    def __init__(self, parent, initial=None):
        super().__init__(parent)
        self.title("学生信息")
        self.resizable(False, False)
        self.result = None
        initial = initial or {}
        self.vars = {f: tk.StringVar(value=initial.get(f, "")) for f in STUDENT_FIELDS}
        for row, field in enumerate(STUDENT_FIELDS):
            ttk.Label(self, text=field, font=("Microsoft YaHei UI", 12)).grid(row=row, column=0, padx=12, pady=8, sticky="e")
            if field == "住宿情况":
                widget = ttk.Combobox(self, textvariable=self.vars[field], values=STAY_OPTIONS, state="readonly", width=25)
            else:
                widget = ttk.Entry(self, textvariable=self.vars[field], width=28)
            widget.grid(row=row, column=1, padx=12, pady=8)
        box = ttk.Frame(self); box.grid(row=len(STUDENT_FIELDS), columnspan=2, pady=14)
        ttk.Button(box, text="保存", command=self.submit).pack(side="left", padx=8)
        ttk.Button(box, text="取消", command=self.destroy).pack(side="left", padx=8)
        self.transient(parent); self.grab_set(); self.protocol("WM_DELETE_WINDOW", self.destroy)

    def submit(self):
        row = {k: v.get().strip() for k, v in self.vars.items()}
        if not row["学号"] or not row["姓名"]:
            messagebox.showwarning("信息不完整", "学号和姓名不能为空", parent=self); return
        if row["住宿情况"] not in STAY_OPTIONS:
            messagebox.showwarning("信息不完整", "请选择住宿情况", parent=self); return
        fill_class_from_id(row)
        self.result = row; self.destroy()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("学生信息与周末留校登记系统")
        self.geometry("1280x760")
        self.minsize(1050, 650)
        self.store = Store()
        self.current_sid = None
        self.week_var = tk.StringVar(value=week_key())
        self.search_var = tk.StringVar()
        self.class_var = tk.StringVar(value="全部班级")
        self.status_var = tk.StringVar(value="请选择左侧学生")
        self.notice_var = tk.StringVar(value="通知栏：暂无通知")
        self.period_buttons = {}
        self.sid_sort_reverse = False
        self.gender_sort_reverse = False
        self.sort_mode = "sid"
        self.build_ui()
        self.refresh_classes(); self.refresh_students()

    def build_ui(self):
        style = ttk.Style(self); style.configure("Treeview", rowheight=35, font=("Microsoft YaHei UI", 11)); style.configure("Treeview.Heading", font=("Microsoft YaHei UI", 11, "bold"))
        top = ttk.Frame(self, padding=10); top.pack(fill="x")
        ttk.Button(top, text="导入 XLSX", command=self.import_xlsx).pack(side="left", padx=4)
        ttk.Button(top, text="导出留校名单", command=self.export_xlsx).pack(side="left", padx=4)
        ttk.Button(top, text="导出住宿学生 XLSX", command=self.export_housing_xlsx).pack(side="left", padx=4)
        ttk.Button(top, text="重置本周信息", command=self.reset_current_week).pack(side="left", padx=4)
        ttk.Button(top, text="新增学生", command=self.add_student).pack(side="left", padx=4)
        ttk.Button(top, text="修改学生", command=self.edit_student).pack(side="left", padx=4)
        ttk.Button(top, text="删除学生", command=self.delete_student).pack(side="left", padx=4)
        ttk.Button(top, text="清空班级数据", command=self.clear_class_data).pack(side="left", padx=4)
        ttk.Label(top, text="  当前周（周一日期）：").pack(side="left")
        ttk.Entry(top, textvariable=self.week_var, width=12).pack(side="left")
        ttk.Button(top, text="切换周", command=self.change_week).pack(side="left", padx=4)

        notice = tk.Label(
            self, textvariable=self.notice_var, anchor="w", justify="left",
            bg="#fff3cd", fg="#8a3b00", padx=12, pady=7,
            font=("Microsoft YaHei UI", 10), wraplength=1220,
        )
        notice.pack(fill="x", padx=10, pady=(0, 8))

        body = ttk.PanedWindow(self, orient="horizontal"); body.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        left = ttk.Frame(body, padding=6); right = ttk.Frame(body, padding=12); body.add(left, weight=3); body.add(right, weight=2)
        filters = ttk.Frame(left); filters.pack(fill="x", pady=(0, 6))
        ttk.Label(filters, text="搜索学号/姓名：").pack(side="left")
        ent = ttk.Entry(filters, textvariable=self.search_var, width=22); ent.pack(side="left", padx=4); ent.bind("<KeyRelease>", lambda _e: self.refresh_students())
        ttk.Combobox(filters, textvariable=self.class_var, state="readonly", width=15).pack(side="left", padx=5)
        self.class_var.trace_add("write", lambda *_: self.refresh_students())
        cols = ("学号", "姓名", "班级", "住宿情况", "性别", "宿舍号", "床位号")
        self.tree = ttk.Treeview(left, columns=cols, show="headings", selectmode="browse")
        widths = (115, 100, 100, 115, 60, 90, 75)
        for c, w in zip(cols, widths):
            if c == "学号":
                self.tree.heading(c, text="学号 ↑", command=self.toggle_sid_sort)
            elif c == "性别":
                self.tree.heading(c, text="性别 男↑女↓", command=self.toggle_gender_sort)
            else:
                self.tree.heading(c, text=c)
            self.tree.column(c, width=w, anchor="center")
        scroll = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview); self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side="left", fill="both", expand=True); scroll.pack(side="right", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self.select_student)

        ttk.Label(right, textvariable=self.status_var, font=("Microsoft YaHei UI", 18, "bold"), wraplength=450).pack(pady=(10, 20))
        ttk.Label(right, text="点击按钮切换状态（绿色=留校，灰色=不留）", font=("Microsoft YaHei UI", 11)).pack(pady=(0, 10))
        for period in PERIODS:
            b = tk.Button(right, text=f"{period}：不留（0）", font=("Microsoft YaHei UI", 15, "bold"), height=2, bg="#d8d8d8", activebackground="#bbbbbb", command=lambda p=period: self.toggle_period(p))
            b.pack(fill="x", pady=7); self.period_buttons[period] = b
        ttk.Separator(right).pack(fill="x", pady=18)
        self.summary = ttk.Label(right, font=("Microsoft YaHei UI", 12), justify="left")
        self.summary.pack(anchor="w"); ttk.Button(right, text="刷新本周统计", command=self.refresh_summary).pack(anchor="w", pady=10)

    def filtered_students(self):
        q = self.search_var.get().strip().lower(); cls = self.class_var.get()
        rows = self.store.data["students"]
        return [s for s in rows if (not q or q in s.get("学号", "").lower() or q in s.get("姓名", "").lower()) and (cls == "全部班级" or s.get("班级") == cls)]

    def refresh_classes(self):
        values = ["全部班级"] + sorted({s.get("班级", "") for s in self.store.data["students"] if s.get("班级")})
        for widget in self.winfo_children():
            pass
        def walk(w):
            for child in w.winfo_children():
                if isinstance(child, ttk.Combobox) and str(child.cget("textvariable")) == str(self.class_var): child["values"] = values
                walk(child)
        walk(self)

    def refresh_students(self):
        selected = self.current_sid
        for item in self.tree.get_children(): self.tree.delete(item)
        def sid_key(student):
            sid = clean(student.get("学号", ""))
            # 数字学号按数值排序；混合字母的学号则使用自然排序键。
            return tuple((0, int(part)) if part.isdigit() else (1, part.lower()) for part in __import__("re").split(r"(\d+)", sid) if part)
        def gender_rank(student):
            gender = clean(student.get("性别", ""))
            return 0 if gender in ("男", "男生", "M", "m") else 1
        rows = self.filtered_students()
        if self.sort_mode == "gender":
            # 性别排序时，男女分组；每组内部固定按学号升序。
            rows = sorted(rows, key=lambda s: (gender_rank(s), sid_key(s)), reverse=self.gender_sort_reverse)
        else:
            rows = sorted(rows, key=sid_key, reverse=self.sid_sort_reverse)
        for s in rows:
            self.tree.insert("", "end", iid=s["学号"], values=[s.get(c, "") for c in self.tree["columns"]])
        if selected and self.tree.exists(selected): self.tree.selection_set(selected); self.tree.see(selected)
        self.refresh_summary()

    def toggle_sid_sort(self):
        if self.sort_mode != "sid":
            self.sort_mode = "sid"
            self.sid_sort_reverse = False
        else:
            self.sid_sort_reverse = not self.sid_sort_reverse
        self.tree.heading("学号", text="学号 ↓" if self.sid_sort_reverse else "学号 ↑")
        self.refresh_students()

    def toggle_gender_sort(self):
        if self.sort_mode != "gender":
            self.sort_mode = "gender"
            self.gender_sort_reverse = False
        else:
            self.gender_sort_reverse = not self.gender_sort_reverse
        self.tree.heading("性别", text="性别 女↑男↓" if self.gender_sort_reverse else "性别 男↑女↓")
        self.refresh_students()

    def select_student(self, _event=None):
        selected = self.tree.selection()
        if not selected: return
        self.current_sid = selected[0]; s = self.store.student(self.current_sid)
        self.status_var.set(f"{s['姓名']}  ·  {s['学号']}\n{s.get('班级','')}  {s.get('住宿情况','')}  宿舍 {s.get('宿舍号','—') or '—'}")
        self.paint_buttons()

    def valid_week(self):
        try:
            key = date.fromisoformat(self.week_var.get().strip())
            return week_key(key)
        except ValueError:
            messagebox.showwarning("日期错误", "请输入 YYYY-MM-DD 格式的日期"); return None

    def change_week(self):
        key = self.valid_week()
        if key: self.week_var.set(key); self.paint_buttons(); self.refresh_summary()

    def paint_buttons(self):
        key = self.valid_week()
        if not self.current_sid or not key: return
        record = self.store.record(key, self.current_sid)
        for p, b in self.period_buttons.items():
            value = 1 if any(int(record.get(k, 0)) for k in PERIOD_MAP[p]) else 0
            b.configure(text=f"{p}：{'留校（1）' if value else '不留（0）'}", bg="#55b96b" if value else "#d8d8d8", activebackground="#45a65a" if value else "#bbbbbb")

    def toggle_period(self, period):
        if not self.current_sid: messagebox.showinfo("提示", "请先在左侧选择一名学生"); return
        key = self.valid_week()
        if not key: return
        record = self.store.record(key, self.current_sid)
        keys = PERIOD_MAP[period]; value = 0 if any(int(record.get(k, 0)) for k in keys) else 1
        for key_name in keys: record[key_name] = value
        self.store.save(); self.paint_buttons(); self.refresh_summary()

    def reset_current_week(self):
        key = self.valid_week()
        if not key: return
        if not messagebox.askyesno("确认重置", f"确定清空 {week_title(key)} 的全部留校和住宿登记吗？\n学生公共信息不会改变。"):
            return
        self.store.data["weeks"].pop(key, None)
        self.store.save()
        self.paint_buttons(); self.refresh_summary()
        self.notice_var.set(f"通知栏：已重置 {week_title(key)} 的全部登记")

    def add_student(self):
        dialog = StudentDialog(self); self.wait_window(dialog)
        if not dialog.result: return
        if self.store.student(dialog.result["学号"]): messagebox.showwarning("重复学号", "该学号已存在"); return
        self.store.data["students"].append(dialog.result); self.store.save(); self.refresh_classes(); self.refresh_students()

    def edit_student(self):
        if not self.current_sid: messagebox.showinfo("提示", "请先选择学生"); return
        old = self.store.student(self.current_sid); dialog = StudentDialog(self, old); self.wait_window(dialog)
        if not dialog.result: return
        new_sid = dialog.result["学号"]
        if new_sid != self.current_sid and self.store.student(new_sid): messagebox.showwarning("重复学号", "该学号已存在"); return
        old.update(dialog.result)
        if new_sid != self.current_sid:
            for records in self.store.data["weeks"].values():
                if self.current_sid in records: records[new_sid] = records.pop(self.current_sid)
            self.current_sid = new_sid
        self.store.save(); self.refresh_classes(); self.refresh_students(); self.paint_buttons()

    def delete_student(self):
        if not self.current_sid: messagebox.showinfo("提示", "请先选择学生"); return
        s = self.store.student(self.current_sid)
        if not messagebox.askyesno("确认删除", f"确定删除 {s['姓名']}（{s['学号']}）吗？"): return
        self.store.data["students"].remove(s)
        for records in self.store.data["weeks"].values(): records.pop(self.current_sid, None)
        self.current_sid = None; self.store.save(); self.status_var.set("请选择左侧学生"); self.refresh_classes(); self.refresh_students()

    def clear_class_data(self):
        classes = sorted({clean(s.get("班级", "")) for s in self.store.data["students"] if clean(s.get("班级", ""))})
        if not classes:
            messagebox.showinfo("提示", "当前没有可清空的班级数据"); return
        class_name = simpledialog.askstring("清空班级数据", "请输入要清空的班级：\n" + "、".join(classes), parent=self)
        if not class_name: return
        class_name = class_name.strip()
        targets = [s for s in self.store.data["students"] if clean(s.get("班级", "")) == class_name]
        if not targets:
            messagebox.showwarning("未找到班级", f"没有找到班级“{class_name}”"); return
        if not messagebox.askyesno("第一次确认", f"将删除班级“{class_name}”的 {len(targets)} 名学生及全部周末登记。\n是否继续？", parent=self): return
        confirm = simpledialog.askstring("第二次确认", f"此操作不可撤销。请输入班级名称 {class_name} 确认删除：", parent=self)
        if not confirm or confirm.strip() != class_name:
            messagebox.showinfo("已取消", "班级名称不一致，未删除任何数据"); return
        ids = {s["学号"] for s in targets}
        self.store.data["students"] = [s for s in self.store.data["students"] if s.get("学号") not in ids]
        for records in self.store.data["weeks"].values():
            for sid in ids: records.pop(sid, None)
        self.current_sid = None; self.store.save(); self.status_var.set("请选择左侧学生")
        if self.class_var.get() == class_name: self.class_var.set("全部班级")
        self.refresh_classes(); self.refresh_students()
        self.notice_var.set(f"通知栏：已清空班级 {class_name}，共删除 {len(targets)} 名学生及其周末登记")

    def import_xlsx(self):
        if not load_workbook: messagebox.showerror("缺少组件", "请先运行：python -m pip install openpyxl"); return
        path = filedialog.askopenfilename(title="选择学生信息表", filetypes=[("Excel 工作簿", "*.xlsx")])
        if not path: return
        try:
            ws = load_workbook(path, read_only=True, data_only=True).active
            rows = list(ws.iter_rows(values_only=True))
            if not rows: raise ValueError("工作表为空")
            headers = [clean(x).replace(" ", "").replace("\n", "") for x in rows[0]]
            mapping = {}
            for field, names in ALIASES.items():
                for i, header in enumerate(headers):
                    if header in names: mapping[field] = i; break
            existing = {s["学号"]: s for s in self.store.data["students"]}
            added = updated = unchanged = skipped = 0
            conflicts = []
            filename_class = ""
            if "学号" not in mapping:
                # 无学号表：从文件名识别“107班”等班级编码，再按姓名匹配已有学生。
                match = re.search(r"(\d{1,3})班", os.path.basename(path))
                filename_class = match.group(1) if match else ""
                if "姓名" not in mapping or not filename_class:
                    raise ValueError("表格没有学号，且无法从文件名识别班级（文件名应包含如“107班”）")
            for values in rows[1:]:
                sid = clean(values[mapping["学号"]]) if "学号" in mapping and mapping["学号"] < len(values) else ""
                incoming = {f: "" for f in STUDENT_FIELDS}
                for field, index in mapping.items():
                    if index < len(values): incoming[field] = clean(values[index])
                if not sid:
                    name = incoming.get("姓名", "")
                    candidates = [s for s in existing.values() if s.get("姓名", "") == name and (s.get("班级", "") == filename_class or not s.get("班级"))]
                    if len(candidates) == 1:
                        sid = candidates[0]["学号"]
                    elif len(candidates) > 1:
                        conflicts.append(f"姓名 {name}（{filename_class}班）匹配到多个学生"); continue
                    else:
                        conflicts.append(f"姓名 {name}（{filename_class}班）未找到对应学号"); continue
                incoming["学号"] = sid
                if not incoming.get("班级"): incoming["班级"] = filename_class
                fill_class_from_id(incoming)
                old = existing.get(sid)
                if old is None:
                    if incoming.get("住宿情况") not in STAY_OPTIONS:
                        incoming["住宿情况"] = "走读无铺"
                    existing[sid] = incoming; added += 1
                    continue
                if clean(old.get("姓名")) != incoming.get("姓名"):
                    conflicts.append(
                        f"学号 {sid}：已有“{old.get('姓名', '')}”，导入表为“{incoming.get('姓名', '')}”"
                    )
                    continue
                merged = dict(old)
                for field in STUDENT_FIELDS:
                    if incoming.get(field): merged[field] = incoming[field]
                if all(clean(old.get(f)) == clean(merged.get(f)) for f in STUDENT_FIELDS):
                    unchanged += 1
                else:
                    existing[sid] = merged; updated += 1
            self.store.data["students"] = list(existing.values()); self.store.save(); self.refresh_classes(); self.refresh_students()
            if conflicts:
                shown = conflicts[:8]
                extra = f"；另有 {len(conflicts) - 8} 条" if len(conflicts) > 8 else ""
                self.notice_var.set("学号与姓名冲突，以下记录未导入：" + "；".join(shown) + extra)
            else:
                self.notice_var.set("通知栏：本次导入没有发现学号与姓名冲突")
            messagebox.showinfo(
                "导入完成",
                f"新增 {added} 人，更新 {updated} 人，完全相同跳过 {unchanged} 人，"
                f"姓名冲突拒绝 {len(conflicts)} 人，空学号跳过 {skipped} 行。\n"
                "姓名冲突详情已显示在顶部通知栏。",
            )
        except Exception as exc: messagebox.showerror("导入失败", str(exc))

    def refresh_summary(self):
        key = self.valid_week()
        if not key: return
        records = self.store.data["weeks"].get(key, {})
        counts = {p: sum(1 for r in records.values() if any(int(r.get(k, 0)) for k in PERIOD_MAP[p])) for p in PERIODS}
        lines = [f"本周：{week_title(key)}", f"学生总数：{len(self.store.data['students'])} 人"] + [f"{p}：{counts[p]} 人" for p in PERIODS]
        self.summary.configure(text="\n".join(lines))

    def export_xlsx(self):
        if not Workbook: messagebox.showerror("缺少组件", "请先安装 openpyxl"); return
        key = self.valid_week()
        if not key: return
        path = filedialog.asksaveasfilename(title="导出本周报表", defaultextension=".xlsx", initialfile=f"周末留校统计_{key}.xlsx", filetypes=[("Excel 工作簿", "*.xlsx")])
        if not path: return
        try:
            wb = Workbook(); detail = wb.active; detail.title = "本周登记明细"
            stay_periods = ["周五晚留校", "周六上午留校", "周六下午留校", "周六晚上留校", "周日留校"]
            headers = ["学号", "姓名"] + stay_periods; detail.append(headers)
            records = self.store.data["weeks"].get(key, {})
            for s in sorted(self.store.data["students"], key=lambda x: (x.get("班级", ""), x["学号"])):
                r = records.get(s["学号"], {})
                # 四个留校时段全部为 0 时，不列入留校名单。
                if not any(int(r.get(p, 0)) for p in stay_periods):
                    continue
                detail.append([s.get("学号", ""), s.get("姓名", "")] + [yes_blank(r.get(p, 0)) for p in stay_periods])
            summary = wb.create_sheet("人数统计"); summary.append(["统计周", week_title(key)]); summary.append(["时段", "留校人数"])
            for p in PERIODS: summary.append([p, sum(1 for r in records.values() if any(int(r.get(k, 0)) for k in PERIOD_MAP[p]))])
            for p in PERIODS:
                ws = wb.create_sheet(p.replace("留校", "").replace("在校", "")[:31]); ws.append(STUDENT_FIELDS)
                for s in self.store.data["students"]:
                    if any(int(records.get(s["学号"], {}).get(k, 0)) for k in PERIOD_MAP[p]): ws.append([s.get(f, "") for f in STUDENT_FIELDS])
            for ws in wb.worksheets:
                for cell in ws[1]: cell.font = Font(bold=True, color="FFFFFF"); cell.fill = PatternFill("solid", fgColor="4472C4")
                for col in ws.columns:
                    letter = col[0].column_letter; ws.column_dimensions[letter].width = min(max(len(clean(c.value)) for c in col) + 4, 28)
                    for c in col: c.alignment = Alignment(horizontal="center", vertical="center")
                ws.freeze_panes = "A2"; ws.auto_filter.ref = ws.dimensions
            wb.save(path); messagebox.showinfo("导出成功", f"报表已保存到：\n{path}")
        except Exception as exc: messagebox.showerror("导出失败", str(exc))

    def export_housing_xlsx(self):
        if not Workbook: messagebox.showerror("缺少组件", "请先安装 openpyxl"); return
        key = self.valid_week()
        if not key: return
        path = filedialog.asksaveasfilename(title="导出住宿学生", defaultextension=".xlsx", initialfile=f"住宿学生_{key}.xlsx", filetypes=[("Excel 工作簿", "*.xlsx")])
        if not path: return
        records = self.store.data["weeks"].get(key, {})
        selected = [s for s in self.store.data["students"] if int(records.get(s["学号"], {}).get("周五晚留宿", 0)) or int(records.get(s["学号"], {}).get("周六晚上留宿", 0))]
        wb = Workbook(); ws = wb.active; ws.title = "住宿学生"
        ws.append(["学号", "姓名", "宿舍号", "周五住宿", "周六住宿"])
        for s in selected:
            r = records.get(s["学号"], {}); ws.append([s.get("学号", ""), s.get("姓名", ""), s.get("宿舍号", ""), yes_blank(r.get("周五晚留宿", 0)), yes_blank(r.get("周六晚上留宿", 0))])
        wb.save(path); messagebox.showinfo("导出成功", f"住宿名单已保存到：\n{path}")


if __name__ == "__main__":
    App().mainloop()
