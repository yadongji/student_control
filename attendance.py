"""请假期限使用真实时间，不依赖界面中选择的统计日期。"""
from datetime import datetime


def on_leave(student, now=None):
    if student.get("状态") not in ("请假中", "请假"):
        return False
    try:
        return datetime.fromisoformat(student["请假截止"]) > (now or datetime.now())
    except (KeyError, TypeError, ValueError):
        return True  # 旧的无期限请假只能从请假管理中手动销假


def set_leave(data, name, until, now=None):
    now = now or datetime.now()
    end = datetime.fromisoformat(until)
    if end <= now:
        raise ValueError("请假截止时间必须晚于当前时间")
    student = data["students"][name]
    student.update({"状态": "请假中", "请假开始": now.isoformat(timespec="seconds"),
                    "请假截止": end.isoformat(timespec="seconds")})
    data.setdefault("leave_history", {}).setdefault(name, []).append(
        {"action": "请假", "from": student["请假开始"], "until": student["请假截止"]})


def return_to_school(data, name, now=None, automatic=False):
    student = data["students"][name]
    student["状态"] = "在校中"
    data.setdefault("leave_history", {}).setdefault(name, []).append(
        {"action": "到期返校" if automatic else "提前销假", "at": (now or datetime.now()).isoformat(timespec="seconds")})


def expire_leaves(data, now=None):
    now = now or datetime.now()
    expired = []
    for name, student in data.get("students", {}).items():
        if student.get("状态") in ("请假中", "请假") and not on_leave(student, now):
            return_to_school(data, name, now, automatic=True)
            expired.append(name)
    return expired
