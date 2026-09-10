"""成绩只按姓名归档。校排名来自输入列，不以班级顺序代替。"""
import hashlib
import math
import re
from datetime import datetime
from pathlib import Path
from uuid import uuid4

SUBJECTS = ("语文", "数学", "英语", "物理", "化学", "生物")
ALIASES = {"语文": ("语文", "语"), "数学": ("数学", "数"), "英语": ("英语", "外语", "英", "外"),
           "物理": ("物理", "物"), "化学": ("化学", "化"), "生物": ("生物", "生")}
CLASS_RANK_NAMES = {"班名次", "班次", "班排名", "班排", "班名", "班级名次", "班级排名"}
SCHOOL_RANK_NAMES = {"校名次", "校排名", "校次", "校排", "校名", "学校名次", "学校排名", "年级名次", "年级排名", "年级次"}
VALUE_SUFFIXES = (".score", ".rank", ".class_rank")


def normalize(value):
    return re.sub(r"[\s（）()_\-：:]", "", str(value or ""))


def column_key(value):
    token = normalize(value)
    if token in ("姓名", "学生姓名", "学生名字", "名字"): return "name"
    if token in ("学号", "学生学号", "考号", "考试号"): return "number"
    if token in ("校人数", "年级人数", "参考人数", "全校参考人数"): return "cohort"
    if token in ("总分", "总成绩", "总分成绩", "成绩总分"): return "总分.score"
    # 不带科目的排名必须由整行列顺序确定，不能默认当作总分校排名。
    for subject, aliases in {**ALIASES, "总分": ("总分", "总成绩", "总")}.items():
        for alias in aliases:
            if not token.startswith(alias): continue
            suffix = token[len(alias):]
            if suffix in ("", "成绩", "分数", "得分", "分"): return subject + ".score"
            if suffix.startswith("成绩"): suffix = suffix[2:]
            if suffix in CLASS_RANK_NAMES: return subject + ".class_rank"
            if suffix in SCHOOL_RANK_NAMES: return subject + ".rank"
    return None


def resolve_headers(headers):
    """未写科目的排名列属于左侧最近的科目分数列；空白列不改变位置。"""
    current = None
    keys = []
    for position, value in enumerate(headers, 1):
        token, key = normalize(value), column_key(value)
        if key and key.endswith(".score"):
            current = key.split(".")[0]
        elif token in CLASS_RANK_NAMES or token in SCHOOL_RANK_NAMES:
            if current is None:
                raise ValueError(f"第 {position} 列“{value}”前没有可识别的科目/总分列，请写明例如“物理校排名”或“总分班名次”")
            key = current + (".class_rank" if token in CLASS_RANK_NAMES else ".rank")
        elif token in {s + suffix for s in ("政治", "历史", "地理", "技术", "体育") for suffix in ("", "成绩", "分数")}:
            # 不把不支持学科之后的排名误写入上一科。
            current = None
        keys.append(key)
    return keys


def field_label(key):
    if "." not in key: return {"name": "姓名", "number": "学号", "cohort": "参考人数"}.get(key, key)
    subject, kind = key.split(".", 1)
    return subject + {"score": "分数", "rank": "校名次", "class_rank": "班名次"}[kind]


def numeric(value, rank=False):
    if value is None or str(value).strip() in ("", "-", "—", "缺考", "未考", "免考"): return None
    if isinstance(value, bool): raise ValueError("不接受布尔值")
    try: number = float(str(value).replace(",", "").strip())
    except ValueError: raise ValueError(f"不是有效数字：{value}")
    if not math.isfinite(number) or number < 0 or (rank and (number < 1 or number != int(number))):
        raise ValueError(f"{'名次必须为正整数' if rank else '成绩必须是非负有限数字'}：{value}")
    return int(number) if number.is_integer() else number


def parse_xlsx(filename, roster):
    from openpyxl import load_workbook
    wb = load_workbook(filename, read_only=True, data_only=True)
    try:
        candidates = []
        for sheet in wb:
            rows = list(sheet.values)
            for index, row in enumerate(rows[:15]):
                if not any(column_key(v) == "name" for v in row): continue
                keys = resolve_headers(row)
                if any(k and k.endswith(VALUE_SUFFIXES) for k in keys):
                    candidates.append((sheet.title, rows, index, keys)); break
        if len(candidates) != 1:
            raise ValueError("请保留一个含成绩的工作表，使用单行列名：学生姓名、语文、语文校排名……总分、总分校排名。" if candidates else "没有找到姓名和成绩表头。支持姓名/学生姓名；多级表头请先整理为单行。")
        sheet, rows, header, keys = candidates[0]
        recognized = [k for k in keys if k]
        duplicates = sorted({k for k in recognized if recognized.count(k) > 1})
        if duplicates: raise ValueError("表头重复映射到同一字段：" + "、".join(field_label(k) for k in duplicates) + "。不同科目后的同名排名列可以重复，同一科同一种排名只能保留一列")
        names = {}
        for name in roster:
            names.setdefault(normalize(name), []).append(name)
        records, unmatched, seen = {}, [], set()
        for line, row in enumerate(rows[header + 1:], header + 2):
            values = dict(zip(keys, row))
            raw_name = str(values.get("name") or "").strip()
            if not raw_name: continue
            name_key = normalize(raw_name)
            if name_key in seen: raise ValueError(f"第 {line} 行姓名重复：{raw_name}。同名学生请先使用不同的姓名标识")
            seen.add(name_key)
            matches = names.get(name_key, [])
            if len(matches) > 1: raise ValueError(f"姓名匹配不唯一：{raw_name}")
            if not matches: unmatched.append(raw_name); continue
            record = {}
            for key, value in values.items():
                if key and (key.endswith(VALUE_SUFFIXES) or key == "cohort"):
                    try: number = numeric(value, key.endswith((".rank", ".class_rank")) or key == "cohort")
                    except ValueError as exc: raise ValueError(f"第 {line} 行 {raw_name} / {key}：{exc}") from exc
                    if number is not None: record[key] = number
            if not any(k.endswith(VALUE_SUFFIXES) for k in record):
                raise ValueError(f"第 {line} 行 {raw_name} 没有有效成绩或名次（公式列请在 Excel 中计算并保存）")
            if any(k.endswith(".rank") and v > record.get("cohort", float("inf")) for k, v in record.items()):
                raise ValueError(f"第 {line} 行校名次超过参考人数")
            records[matches[0]] = record
        if not records: raise ValueError("没有能够匹配当前学生名单的成绩")
        return {"title": Path(filename).stem, "source": Path(filename).name, "sheet": sheet,
                "sha256": hashlib.sha256(Path(filename).read_bytes()).hexdigest(),
                "records": records, "unmatched": unmatched,
                "recognized": recognized,
                "column_mapping": [f"第 {i} 列 {v} → {field_label(k)}" for i, (v, k) in enumerate(zip(rows[header], keys), 1) if k],
                "unrecognized": [str(v) for v, k in zip(rows[header], keys) if v and not k]}
    finally: wb.close()


def commit_exam(data, draft, title, now=None):
    if not title.strip(): raise ValueError("考试名称不能为空")
    if any(e.get("sha256") == draft["sha256"] for e in data.get("exams", [])):
        raise ValueError("这个文件已经导入，不能重复计为一次考试")
    exam = {k: draft[k] for k in ("source", "sheet", "sha256", "records")}
    exam.update(id=uuid4().hex, title=title.strip(), imported_at=(now or datetime.now()).isoformat(timespec="microseconds"),
                sequence=max((e.get("sequence", 0) for e in data.get("exams", [])), default=0) + 1)
    data.setdefault("exams", []).append(exam)
    return exam


def student_exams(data, name):
    return sorted((e for e in data.get("exams", []) if name in e.get("records", {})),
                  key=lambda e: (e["imported_at"], e.get("sequence", 0)))


def comparisons(current, previous):
    results = []
    for subject in (*SUBJECTS, "总分"):
        score, rank = current.get(subject + ".score"), current.get(subject + ".rank")
        old_score, old_rank = previous.get(subject + ".score"), previous.get(subject + ".rank")
        results.append({"subject": subject, "score": score, "rank": rank, "class_rank": current.get(subject + ".class_rank"),
                        "score_change": None if score is None or old_score is None else round(score - old_score, 2),
                        "rank_change": None if rank is None or old_rank is None else old_rank - rank})
    return results


def subject_profile(exam, name):
    """校排名齐全时用校排；否则统一使用本次导入样本分位，不跨科比原始分。"""
    records = exam["records"]
    row = records.get(name, {})
    ranks = {s: row[s + ".rank"] for s in SUBJECTS if s + ".rank" in row}
    if len(ranks) >= 2:
        # 同一学生同一场考试的相对校排名，不冒充全校百分位。
        ordered = sorted(set(ranks.values()))
        if len(ordered) == 1: return {}, "各科校名次相同，暂无明显优势差异"
        return {s: 1 - ordered.index(r) / (len(ordered) - 1) for s, r in ranks.items()}, "本次各科校名次（相对比较）"
    strength = {}
    for subject in SUBJECTS:
        key = subject + ".score"
        if key not in row: continue
        sample = [r[key] for r in records.values() if key in r]
        if len(sample) < 2 or len(set(sample)) < 2: continue
        strength[subject] = (sum(v < row[key] for v in sample) + .5 * (sum(v == row[key] for v in sample) - 1)) / (len(sample) - 1)
    return strength, "本次导入样本的各科分位（非校排名）"


def profile_text(exam, name):
    values, basis = subject_profile(exam, name)
    if len(values) < 2 or max(values.values()) == min(values.values()): return "数据不足或各科表现接近，暂不判断优势/待提升学科。", basis
    strong = "、".join(s for s, v in values.items() if v == max(values.values()))
    weak = "、".join(s for s, v in values.items() if v == min(values.values()))
    return f"相对优势：{strong}；待提升：{weak}", basis


def pairing_value(a, b):
    shared = sorted(set(a) & set(b))
    if len(shared) < 2: return 0.0
    av, bv = sum(a[s] for s in shared)/len(shared), sum(b[s] for s in shared)/len(shared)
    # 只有不同方向的相对强项才是互补，不把全面强弱差距当成互补。
    return sum(max(0, -(a[s]-av)*(b[s]-bv)) for s in shared)


def complementary_seats(data, layout):
    exams = sorted(data.get("exams", []), key=lambda e: (e["imported_at"], e.get("sequence", 0)))
    if not exams: raise ValueError("请先导入成绩")
    exam = exams[-1]
    names = list(data.get("students", {}))
    capacity = sum(cols * rows for _, cols, rows in layout)
    if len(names) > capacity: raise ValueError("学生人数超过座位数，不能排座")
    profiles = {n: subject_profile(exam, n)[0] for n in names}
    if sum(len(v) >= 2 for v in profiles.values()) < 2: raise ValueError("最近一次考试中可用于比较的学科不足，请补充成绩")
    edges = sorted(((pairing_value(profiles[a], profiles[b]), a, b)
                    for i, a in enumerate(names) for b in names[i+1:]), key=lambda x: (-x[0], x[1], x[2]))
    remaining, pairs = set(names), []
    for value, a, b in edges:
        if value <= 0: break
        if a in remaining and b in remaining:
            pairs.append((a, b)); remaining.remove(a); remaining.remove(b)
    # 相邻只在同一排同一区内定义，不跨过道。三列区第三列为单座。
    slots, singles = [], []
    for zone, cols, rows in layout:
        for row in range(1, rows+1):
            for col in range(1, cols, 2): slots.append((f"{zone}{row}-{col}", f"{zone}{row}-{col+1}"))
            if cols % 2: singles.append(f"{zone}{row}-{cols}")
    assigned = {}
    for seats, pair in zip(slots, pairs): assigned.update(zip(seats, pair))
    placed = set(assigned.values())
    unused = [s for pair in slots for s in pair if s not in assigned] + singles
    for seat, name in zip(unused, sorted(set(names)-placed)): assigned[seat] = name
    for zone, cols, rows in layout:
        for row in range(1, rows+1):
            for col in range(1, cols+1): assigned.setdefault(f"{zone}{row}-{col}", "")
    return assigned, exam, min(len(pairs), len(slots))
