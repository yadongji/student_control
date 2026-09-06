#!/usr/bin/env python3
"""离线班级座位与课堂发言管理系统。"""
import json, os, re, random, shutil, tkinter as tk
from datetime import date, datetime
from tkinter import filedialog, messagebox, ttk, simpledialog
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    BaseTk = TkinterDnD.Tk
except ImportError:
    DND_FILES = None
    BaseTk = tk.Tk
try:
    from openpyxl import load_workbook
except ImportError:
    load_workbook = None
try:
    from PIL import Image, ImageEnhance, ImageOps
    import pytesseract
    for _tess in (r"C:\Program Files\Tesseract-OCR\tesseract.exe", r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"):
        if os.path.exists(_tess):
            pytesseract.pytesseract.tesseract_cmd = _tess
            break
except ImportError:
    Image = pytesseract = None

BASE = os.path.dirname(os.path.abspath(__file__))
FILE = os.path.join(BASE, "seat_data.json")
EVIDENCE_DIR = os.path.join(BASE, "evidence")
ROWS = 8
SEAT_LAYOUT = [("左", 3, 8), ("中", 4, 8), ("右", 3, 8)]
SUBJECTS = ["语文", "数学", "英语", "物理", "化学", "生物"]

def load_data():
    if os.path.exists(FILE):
        try:
            with open(FILE, encoding="utf-8") as f: data=json.load(f)
            if not isinstance(data, dict): data={}
            data.setdefault("students", {}); data.setdefault("seats", {}); data.setdefault("speech", {}); data.setdefault("homework", {}); data.setdefault("sleep", {}); data.setdefault("evidence", {})
            return data
        except Exception: pass
    return {"students": {}, "seats": {}, "speech": {}, "homework": {}, "sleep": {}, "evidence": {}}

def save_data(data):
    with open(FILE + ".tmp", "w", encoding="utf-8") as f: json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(FILE + ".tmp", FILE)

class App(BaseTk):
    def __init__(self):
        super().__init__(); self.title("班级座位与发言管理系统"); self.geometry("1400x820"); self.data=load_data(); self.buttons={}; self.current_day=date.today().isoformat(); self.drag_seat=None; self.search=tk.StringVar()
        self.build(); self.draw_seats(); self.refresh_summary()

    def build(self):
        top=ttk.Frame(self,padding=8); top.pack(fill="x")
        ttk.Button(top,text="导入座次表 XLSX",command=self.import_xlsx).pack(side="left",padx=4)
        ttk.Button(top,text="导出发言统计",command=self.export_report).pack(side="left",padx=4)
        ttk.Button(top,text="今日发言清零",command=self.reset_today).pack(side="left",padx=4)
        ttk.Button(top,text="发言排行榜",command=self.leaderboard).pack(side="left",padx=4)
        ttk.Button(top,text="随机排座",command=self.random_seats).pack(side="left",padx=4)
        ttk.Button(top,text="发布作业",command=self.publish_homework).pack(side="left",padx=4)
        ttk.Button(top,text="今日未交作业名单",command=self.missing_homework).pack(side="left",padx=4)
        ttk.Label(top,text="搜索：").pack(side="left",padx=(15,2)); e=ttk.Entry(top,textvariable=self.search,width=16); e.pack(side="left"); e.bind("<KeyRelease>",lambda _:self.paint())
        ttk.Label(top,text="日期：").pack(side="left",padx=(20,2)); self.day_var=tk.StringVar(value=self.current_day); ttk.Entry(top,textvariable=self.day_var,width=12).pack(side="left")
        ttk.Button(top,text="切换日期",command=self.change_day).pack(side="left",padx=4)
        self.summary=ttk.Label(top,text="",font=("Microsoft YaHei UI",11)); self.summary.pack(side="right",padx=10)
        note=ttk.Label(self,text="点击学生按钮查看详情；点击“记录一次发言”会增加当天次数。学生信息按姓名保存，与座位独立。",foreground="#555",padding=(10,0,10,8)); note.pack(anchor="w")
        self.board=ttk.Frame(self,padding=10); self.board.pack(fill="both",expand=True)

    def draw_seats(self):
        for w in self.board.winfo_children(): w.destroy()
        self.buttons={}; col=0
        for name, width, height in SEAT_LAYOUT:
            frame=ttk.LabelFrame(self.board,text=f"{name}区",padding=6); frame.grid(row=0,column=col,columnspan=width,padx=14 if col else 0,sticky="nsew")
            for c in range(width): frame.columnconfigure(c,weight=1)
            for r in range(height):
                frame.rowconfigure(r,weight=1)
                for c in range(width):
                    seat=f"{name}{r+1}-{c+1}"; b=tk.Button(frame,font=("Microsoft YaHei UI",10,"bold"),width=11,height=3,command=lambda s=seat:self.open_student(s)); b.grid(row=r,column=c,padx=3,pady=3,sticky="nsew"); b.bind("<ButtonPress-1>",lambda e,s=seat:self.drag_start(s)); b.bind("<ButtonRelease-1>",lambda e,s=seat:self.drag_end(s)); self.buttons[seat]=b
            col+=width

    def student_for(self, seat):
        name=self.data["seats"].get(seat,""); return name,self.data["students"].get(name,{})
    def paint(self):
        for seat,b in self.buttons.items():
            name,s=self.student_for(seat); count=int(self.data["speech"].get(name,{}).get(self.current_day,0)); gender=s.get("性别",""); q=self.search.get().strip().lower(); visible=not q or q in name.lower() or q in str(s.get("学号","")).lower()
            status=s.get("状态",""); label=f"{name or '空座'}\n{'💬 '+str(count) if name else ''}{'\n'+status if status else ''}"
            b.configure(text=label,bg="#b9dcff" if gender in ("男","男生","M") else ("#ffc4d6" if gender else "#e6e6e6"),activebackground="#8fc4fa",state="normal" if visible else "disabled")
    def refresh_summary(self):
        total=sum(int(v.get(self.current_day,0)) for v in self.data["speech"].values()); self.summary.configure(text=f"{self.current_day} 发言总次数：{total}    学生：{len(self.data['students'])} 人"); self.paint()
    def change_day(self):
        try: date.fromisoformat(self.day_var.get().strip())
        except ValueError: messagebox.showwarning("日期错误","日期请输入 YYYY-MM-DD"); return
        self.current_day=self.day_var.get().strip(); self.refresh_summary()
    def open_student(self, seat):
        name,s=self.student_for(seat)
        if not name: messagebox.showinfo("空座","该座位尚未安排学生"); return
        win=tk.Toplevel(self); win.title(f"学生详情 - {name}"); win.geometry("900x650"); win.grab_set()
        total=sum(int(v) for v in self.data["speech"].get(name,{}).values()); recent=self.data["speech"].get(name,{})
        homework=self.data.setdefault("homework",{}).setdefault(self.current_day,{})
        missing_total=sum(len(info.get("missing",[])) for day in self.data.get("homework",{}).values() for info in day.values() if name in info.get("missing",[])); score=max(0,100-total*.5-int(s.get("迟到",0))-missing_total*2)
        sleep_recent=self.data.setdefault("sleep",{}).get(name,{}); sleep_total=sum(int(v) for v in sleep_recent.values())
        text=f"姓名：{name}\n学号：{s.get('学号','')}\n性别：{s.get('性别','')}\n状态：{s.get('状态','正常')}\n纪律评分：{score:g}\n迟到次数：{s.get('迟到',0)}\n备注：{s.get('备注','')}\n座位：{seat}\n\n当天说话：{recent.get(self.current_day,0)} 次\n历史说话次数：***\n当天瞌睡：{sleep_recent.get(self.current_day,0)} 次　历史瞌睡：{sleep_total} 次\n最近记录：\n"+"\n".join(f"{d}：{n} 次" for d,n in sorted(recent.items(),reverse=True)[:7])
        content=tk.Frame(win); content.pack(fill="both",expand=True); left=tk.Frame(content,width=450); left.pack(side="left",fill="both",expand=True); right=tk.LabelFrame(content,text="本周作业",font=("Microsoft YaHei UI",13),padx=15,pady=10); right.pack(side="right",fill="both",expand=True,padx=10,pady=10)
        info_label=ttk.Label(left,text=text,justify="left",font=("Microsoft YaHei UI",13),padding=18); info_label.pack(anchor="nw")
        hw_buttons={}
        for sub in SUBJECTS:
            if sub not in homework:
                continue
            info=homework[sub]; submitted=name not in info.get("missing",[]); row=tk.Frame(right); row.pack(fill="x",pady=5); tk.Label(row,text=f"{sub}作业",font=("Microsoft YaHei UI",11),anchor="w").pack(side="left",fill="x",expand=True)
            btn=tk.Button(row,text="已交" if submitted else "未交",width=8,bg="#55b96b" if submitted else "#f19a9a"); btn.pack(side="right"); hw_buttons[sub]=btn
            btn.configure(command=lambda su=sub,b=btn:self.toggle_homework(name,su,b))
        eye=tk.BooleanVar(value=False); eye_btn=tk.Button(left,text="🙈",font=("Segoe UI Emoji",10),width=2,relief="flat")
        def eye_toggle():
            eye.set(not eye.get()); eye_btn.configure(text="👁" if eye.get() else "🙈")
            info_label.configure(text=text.replace("历史说话次数：***", f"历史说话次数：{total} 次") if eye.get() else text)
        eye_btn.configure(command=eye_toggle); eye_btn.place(in_=left,x=165,y=238)
        actions=tk.Frame(left); actions.pack(anchor="w",padx=18,pady=8)
        ttk.Button(actions,text="说话一次",command=lambda:self.add_speech(name,win)).pack(side="left",padx=3)
        ttk.Button(actions,text="迟到一次",command=lambda:self.add_late(name,win)).pack(side="left",padx=3)
        ttk.Button(actions,text="瞌睡一次",command=lambda:self.add_sleep(name,win)).pack(side="left",padx=3)
        evidence_box=tk.LabelFrame(left,text="图片证据（拖入图片即记录一次；也可点击选择）",padx=8,pady=8); evidence_box.pack(fill="x",padx=18,pady=8)
        speech_drop=tk.Label(evidence_box,text="拖入说话证据图片",bg="#dceeff",height=3,cursor="hand2"); speech_drop.pack(side="left",fill="both",expand=True,padx=4)
        sleep_drop=tk.Label(evidence_box,text="拖入瞌睡证据图片",bg="#fff0c9",height=3,cursor="hand2"); sleep_drop.pack(side="left",fill="both",expand=True,padx=4)
        speech_drop.bind("<Button-1>",lambda _e:self.choose_evidence(name,"说话",win)); sleep_drop.bind("<Button-1>",lambda _e:self.choose_evidence(name,"瞌睡",win))
        if DND_FILES:
            for widget,event_type in ((speech_drop,"说话"),(sleep_drop,"瞌睡")):
                widget.drop_target_register(DND_FILES); widget.dnd_bind("<<Drop>>",lambda e,t=event_type:self.drop_evidence(e,name,t,win))

    def edit_extra(self,name,win):
        s=self.data["students"].setdefault(name,{"姓名":name}); s["状态"]=simpledialog.askstring("状态","输入：正常 / 缺勤 / 请假",initialvalue=s.get("状态","正常"),parent=win) or s.get("状态","正常"); s["组别"]=simpledialog.askstring("小组","小组名称（可留空）",initialvalue=s.get("组别",""),parent=win) or s.get("组别",""); s["评分"]=simpledialog.askinteger("纪律评分","输入 0-100 分",initialvalue=int(s.get("评分",0)),minvalue=0,maxvalue=100,parent=win) or 0; s["备注"]=simpledialog.askstring("备注","纪律备注",initialvalue=s.get("备注",""),parent=win) or ""; save_data(self.data); win.destroy(); self.paint(); self.open_student(next(seat for seat,n in self.data["seats"].items() if n==name))

    def drag_start(self,seat): self.drag_seat=seat
    def drag_end(self,seat):
        if self.drag_seat and self.drag_seat!=seat:
            self.data["seats"][self.drag_seat],self.data["seats"][seat]=self.data["seats"].get(seat,""),self.data["seats"].get(self.drag_seat,""); save_data(self.data); self.paint()
        self.drag_seat=None

    def random_seats(self):
        if not messagebox.askyesno("确认","随机安排所有学生座位？学生发言记录不会改变"): return
        seats=list(self.buttons); names=list(self.data["students"]); random.shuffle(names); self.data["seats"]={s:(names[i] if i<len(names) else "") for i,s in enumerate(seats)}; save_data(self.data); self.paint()

    def leaderboard(self):
        win=tk.Toplevel(self); win.title("发言次数排行榜"); win.geometry("500x550"); tree=ttk.Treeview(win,columns=("排名","姓名","学号","总次数"),show="headings")
        for c in tree["columns"]: tree.heading(c,text=c); tree.column(c,width=110,anchor="center")
        rows=sorted(((n,s.get("学号",""),sum(int(v) for v in self.data["speech"].get(n,{}).values())) for n,s in self.data["students"].items()),key=lambda x:x[2],reverse=True)
        for i,(n,s,total) in enumerate(rows,1): tree.insert("","end",values=(i,n,s,total))
        tree.pack(fill="both",expand=True,padx=10,pady=10)
    def add_speech(self,name,win=None):
        rec=self.data["speech"].setdefault(name,{}); rec[self.current_day]=int(rec.get(self.current_day,0))+1; save_data(self.data); self.refresh_summary()
        if win: win.destroy(); self.open_student(next(seat for seat,n in self.data["seats"].items() if n==name))
    def add_late(self,name,win=None):
        s=self.data["students"].setdefault(name,{"姓名":name}); s["迟到"]=int(s.get("迟到",0))+1; save_data(self.data)
        if win: win.destroy(); self.open_student(next(seat for seat,n in self.data["seats"].items() if n==name))
    def add_sleep(self,name,win=None):
        rec=self.data.setdefault("sleep",{}).setdefault(name,{}); rec[self.current_day]=int(rec.get(self.current_day,0))+1; save_data(self.data)
        if win: win.destroy(); self.open_student(next(seat for seat,n in self.data["seats"].items() if n==name))
    def choose_evidence(self,name,event_type,win):
        path=filedialog.askopenfilename(parent=win,title=f"选择{event_type}证据图片",filetypes=[("图片","*.jpg;*.jpeg;*.png;*.bmp;*.webp")])
        if path:self.save_evidence(name,event_type,path,win)
    def drop_evidence(self,event,name,event_type,win):
        paths=list(self.tk.splitlist(event.data))
        if paths:self.save_evidence(name,event_type,paths[0],win)
    def save_evidence(self,name,event_type,path,win):
        if not os.path.isfile(path) or os.path.splitext(path)[1].lower() not in (".jpg",".jpeg",".png",".bmp",".webp"):
            messagebox.showwarning("文件不正确","请拖入 JPG、PNG、BMP 或 WEBP 图片",parent=win); return
        safe_name=re.sub(r'[\\/:*?"<>|]',"_",name); folder=os.path.join(EVIDENCE_DIR,safe_name,self.current_day); os.makedirs(folder,exist_ok=True)
        stamp=datetime.now().strftime("%H%M%S_%f"); target=os.path.join(folder,f"{event_type}_{stamp}{os.path.splitext(path)[1].lower()}"); shutil.copy2(path,target)
        relative=os.path.relpath(target,BASE); records=self.data.setdefault("evidence",{}).setdefault(name,[]); records.append({"日期":self.current_day,"类型":event_type,"图片":relative})
        if event_type=="说话": rec=self.data["speech"].setdefault(name,{}); rec[self.current_day]=int(rec.get(self.current_day,0))+1
        else: rec=self.data.setdefault("sleep",{}).setdefault(name,{}); rec[self.current_day]=int(rec.get(self.current_day,0))+1
        save_data(self.data); self.refresh_summary(); win.destroy(); self.open_student(next(seat for seat,n in self.data["seats"].items() if n==name))
    def publish_homework(self):
        win=tk.Toplevel(self); win.title("发布作业"); win.resizable(False,False); tk.Label(win,text="请选择要发布的科目",font=("Microsoft YaHei UI",13),padx=20,pady=12).pack()
        box=tk.Frame(win); box.pack(padx=15,pady=10)
        for i,sub in enumerate(SUBJECTS):
            tk.Button(box,text=sub,width=12,height=2,font=("Microsoft YaHei UI",12),command=lambda s=sub:self.publish_subject(s,win)).grid(row=i//3,column=i%3,padx=5,pady=5)
    def publish_subject(self,sub,win):
        info=self.data.setdefault("homework",{}).setdefault(self.current_day,{}).setdefault(sub,{"published":0,"missing":[]}); info["published"]=len(self.data["students"]); info["missing"]=[]; save_data(self.data); win.destroy(); self.homework_summary(sub)
    def toggle_homework(self,name,sub,button):
        info=self.data.setdefault("homework",{}).setdefault(self.current_day,{}).setdefault(sub,{"published":len(self.data["students"]),"missing":[]}); missing=info.setdefault("missing",[])
        if name in missing: missing.remove(name); button.configure(text="已交",bg="#55b96b")
        else: missing.append(name); button.configure(text="未交",bg="#f19a9a")
        save_data(self.data); self.refresh_summary()
    def homework_summary(self, subject=None):
        win=tk.Toplevel(self); win.title("作业统计"); win.geometry("560x430")
        for sub in SUBJECTS:
            info=self.data.setdefault("homework",{}).setdefault(self.current_day,{}).get(sub)
            if not info: continue
            missing=info.get("missing",[]); published=int(info.get("published",len(self.data["students"]))); row=tk.Frame(win); row.pack(fill="x",padx=12,pady=4); tk.Label(row,text=f"{sub}：已交{published-len(missing)}人  未交{len(missing)}人",font=("Microsoft YaHei UI",11),anchor="w").pack(side="left",fill="x",expand=True); tk.Button(row,text="查看未交",command=lambda s=sub:self.show_missing(s)).pack(side="right",padx=3); tk.Button(row,text="删除作业",command=lambda s=sub,w=win:self.delete_homework(s,w)).pack(side="right",padx=3)
    def show_missing(self,sub):
        info=self.data.get("homework",{}).get(self.current_day,{}).get(sub,{"missing":[]}); names=info.get("missing",[]); messagebox.showinfo(f"{sub}未交名单", "\n".join(names) if names else "暂无未交学生", parent=self)
    def delete_homework(self, sub, win=None):
        if not messagebox.askyesno("确认删除", f"确定删除 {self.current_day} 的{sub}作业吗？", parent=win or self): return
        day=self.data.get("homework",{}).get(self.current_day,{})
        day.pop(sub,None)
        if not day: self.data.get("homework",{}).pop(self.current_day,None)
        save_data(self.data)
        if win: win.destroy()
        self.refresh_summary()
        self.homework_summary()
    def missing_homework(self):
        win=tk.Toplevel(self); win.title("今日未交作业名单"); win.geometry("620x500")
        tree=ttk.Treeview(win,columns=("科目","姓名","学号","历史未交"),show="headings")
        for c in tree["columns"]: tree.heading(c,text=c); tree.column(c,width=140,anchor="center")
        today=self.data.setdefault("homework",{}).get(self.current_day,{})
        for sub,info in today.items():
            for name,s in self.data["students"].items():
                if name in info.get("missing",[]):
                    history=sum(1 for d,all_sub in self.data.get("homework",{}).items() if d!=self.current_day and sub in all_sub and name in all_sub[sub].get("missing",[]))
                    tree.insert("","end",values=(sub,name,s.get("学号",""),history if history else ""))
        tree.pack(fill="both",expand=True,padx=10,pady=10)
    def homework_entry(self,name,win=None):
        sub=simpledialog.askstring("作业登记","科目（语文/数学/英语/物理/化学/生物）",parent=win or self)
        if sub not in SUBJECTS:return
        info=self.data.setdefault("homework",{}).setdefault(self.current_day,{}).setdefault(sub,{"published":len(self.data["students"]),"submitted":[]}); submitted=info.setdefault("submitted",[])
        if name in submitted: submitted.remove(name); result="已标记为未交"
        else: submitted.append(name); result="已标记为已交"
        save_data(self.data); messagebox.showinfo("作业登记",result,parent=win or self)
        if win: win.destroy(); self.open_student(next(seat for seat,n in self.data["seats"].items() if n==name))
    def reset_today(self):
        if not messagebox.askyesno("确认","清零当天所有学生的发言次数？历史日期不会改变"): return
        for rec in self.data["speech"].values(): rec.pop(self.current_day,None)
        save_data(self.data); self.refresh_summary()
    def import_xlsx(self):
        if not load_workbook: messagebox.showerror("缺少组件","请安装 openpyxl"); return
        path=filedialog.askopenfilename(filetypes=[("Excel 工作簿","*.xlsx")]);
        if not path:return
        try:
            rows=list(load_workbook(path,read_only=True,data_only=True).active.values)
            if not rows: raise ValueError("表格为空")
            headers=[str(x or "").strip() for x in rows[0]]; name_i=next((i for i,h in enumerate(headers) if h in ("姓名","学生姓名","名字")),None); sid_i=next((i for i,h in enumerate(headers) if h in ("学号","学生学号","编号")),None); gender_i=next((i for i,h in enumerate(headers) if h=="性别"),None); seat_i=next((i for i,h in enumerate(headers) if h in ("座位","座位号","位置")),None)
            if name_i is None or seat_i is None:
                # 兼容矩阵座次表：每个非空单元格视为姓名，按列/行生成座位号。
                max_row, max_col = len(rows), max(len(r) for r in rows)
                # 座次表约定为左3列、中4列、右3列，区域之间各留1列过道；允许末尾多一行空白。
                valid_shape = max_col == 12 and max_row in (7, 8) and all(
                    not any(str(rows[r][c] or "").strip() for r in range(max_row))
                    for c in (3, 8) if c < max_col
                )
                if not valid_shape:
                    raise ValueError(
                        f"座次表格式不正确：检测到 {max_row} 行、{max_col} 列。\n"
                        "矩阵格式应为左3列 × 7/8行、中4列 × 7/8行、右3列 × 7/8行，区域之间留空列。\n"
                        "请检查是否多了标题行、少了空过道列，或左右区域列数不一致。"
                    )
                # 该类座次表通常是 A:D、F:H、J:L 三个区域；对应中区4列、左区3列、右区3列。
                blocks = [(0, 3, "左"), (4, 8, "中"), (9, 12, "右")]
                # 导入前清空本次表格覆盖的座位，Excel 空白格就严格对应为空座。
                for start, end, zone in blocks:
                    max_cols = end - start
                    max_rows = min(8, max_row)
                    for rr in range(1, max_rows + 1):
                        for cc in range(1, max_cols + 1):
                            self.data["seats"][f"{zone}{rr}-{cc}"] = ""
                for r, row in enumerate(rows, 1):
                    for start, end, zone in blocks:
                        for c in range(start, min(end, len(row))):
                            value = row[c]
                            name = str(value or "").strip()
                            col = c - start + 1; seat = f"{zone}{r}-{col}"
                            if name:
                                self.data["students"].setdefault(name, {"姓名": name}); self.data["seats"][seat] = name
                save_data(self.data); self.draw_seats(); self.refresh_summary(); messagebox.showinfo("导入完成", "已按左/中/右矩阵区域刷新座位"); return
            for row in rows[1:]:
                if name_i>=len(row) or seat_i>=len(row):continue
                name=str(row[name_i] or "").strip(); seat=str(row[seat_i] or "").strip()
                if not name or not seat:continue
                old=self.data["students"].get(name,{"姓名":name}); old["姓名"]=name
                if sid_i is not None and sid_i<len(row):old["学号"]=str(row[sid_i] or "").strip()
                if gender_i is not None and gender_i<len(row):old["性别"]=str(row[gender_i] or "").strip()
                self.data["students"][name]=old; self.data["seats"][seat]=name
            save_data(self.data); self.draw_seats(); self.refresh_summary(); messagebox.showinfo("导入完成","座次和学生信息已刷新")
        except Exception as e: messagebox.showerror("导入失败",str(e))

    def import_image(self):
        if not Image or not pytesseract:
            messagebox.showerror("缺少组件", "图片识别需要安装 Pillow 和 pytesseract。\n另外需要安装 Tesseract OCR 引擎并加入 PATH。")
            return
        path = filedialog.askopenfilename(title="选择座次表图片", filetypes=[("图片", "*.jpg;*.jpeg;*.png;*.bmp")])
        if not path: return
        try:
            image = Image.open(path).convert("L")
            image = ImageOps.autocontrast(image)
            image = ImageEnhance.Contrast(image).enhance(1.6)
            image = image.resize((image.width * 2, image.height * 2))
            info = pytesseract.image_to_data(image, lang="chi_sim", config="--psm 6", output_type=pytesseract.Output.DICT)
            width, height = image.size; found = []
            for i, raw in enumerate(info["text"]):
                name = re.sub(r"[^\u4e00-\u9fff·]", "", raw or "")
                if len(name) < 2 or len(name) > 5 or float(info["conf"][i]) < 15: continue
                x, y, w, h = (int(info[k][i]) for k in ("left", "top", "width", "height")); cx=x+w/2; cy=y+h/2
                # 图片中三个区域按横向比例映射，过道区域自动跳过。
                if cx < width * .30: zone, cols, start = "左", 3, 0
                elif cx < width * .70: zone, cols, start = "中", 4, width * .30
                else: zone, cols, start = "右", 3, width * .70
                row = min(7, max(0, int((cy / height) * 8))); col = min(cols-1, max(0, int(((cx-start) / (width*(.30 if zone != "中" else .40))) * cols)))
                seat=f"{zone}{row+1}-{col+1}"; found.append((seat,name))
            if not found: raise ValueError("没有识别到清晰的中文姓名")
            preview="\n".join(f"{seat}：{name}" for seat,name in found[:40]); more=f"\n……另有 {len(found)-40} 个" if len(found)>40 else ""
            if not messagebox.askyesno("识别结果确认", f"共识别到 {len(found)} 个姓名：\n\n{preview}{more}\n\n确认写入座位系统吗？"):
                return
            for seat,name in found:
                self.data["students"].setdefault(name,{"姓名":name}); self.data["seats"][seat]=name
            save_data(self.data); self.draw_seats(); self.refresh_summary(); messagebox.showinfo("导入完成", "图片识别结果已写入，可在界面中拖动修正。")
        except Exception as exc: messagebox.showerror("图片识别失败", str(exc))
    def export_report(self):
        path=filedialog.asksaveasfilename(defaultextension=".xlsx",initialfile=f"发言统计_{self.current_day}.xlsx",filetypes=[("Excel 工作簿","*.xlsx")]);
        if not path:return
        try:
            from openpyxl import Workbook
            wb=Workbook(); ws=wb.active; ws.title="发言统计"; ws.append(["学号","姓名","性别","座位","当天次数","历史总次数"])
            for seat,name in self.data["seats"].items():
                s=self.data["students"].get(name,{}); rec=self.data["speech"].get(name,{}); ws.append([s.get("学号",""),name,s.get("性别",""),seat,int(rec.get(self.current_day,0)),sum(int(x) for x in rec.values())])
            wb.save(path); messagebox.showinfo("导出成功",path)
        except Exception as e: messagebox.showerror("导出失败",str(e))

if __name__=="__main__": App().mainloop()
