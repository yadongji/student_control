"""学生纪律分与班级活动记录；兼容原有 JSON 字段。"""
from datetime import date, datetime
from uuid import uuid4

ACTIVITIES = {"做板报": 1, "讲题": 1, "问问题": .5, "搬水": .2, "未打扫卫生": -1, "逃操": -1}


def activity_counts(data, name):
    counts = dict.fromkeys(ACTIVITIES, 0)
    for event in data.get("activities", {}).get(name, []):
        kind = event.get("type")
        if kind in counts:
            counts[kind] += 1
    return counts


def discipline_score(data, name):
    speech = sum(int(n) for n in data.get("speech", {}).get(name, {}).values())
    sleep = sum(int(n) for n in data.get("sleep", {}).get(name, {}).values())
    late = int(data.get("students", {}).get(name, {}).get("迟到", 0))
    missing = sum(1 for day in data.get("homework", {}).values()
                  for task in day.values() if name in task.get("missing", []))
    counts = activity_counts(data, name)
    return round(max(0, 100 - speech * .5 - sleep * .5 - late - missing
               + sum(counts[kind] * delta for kind, delta in ACTIVITIES.items())), 2)


def score_color(score):
    """85 分属于红色，90 分属于橙色；95 分及以上绿色。"""
    if score < 80: return "#dd5555"
    if score <= 85: return "#f07878"
    if score <= 90: return "#f4ae65"
    if score < 95: return "#edda75"
    return "#80cd90"


def question_recorded(data, name, day=None):
    day = day or date.today().isoformat()
    return any(e.get("type") == "问问题" and e.get("date") == day
               for e in data.get("activities", {}).get(name, []))


def record_activity(data, name, kind, day):
    if name not in data.get("students", {}):
        raise ValueError("学生不存在")
    if kind not in ACTIVITIES:
        raise ValueError("未知记录类型")
    date.fromisoformat(day)
    if kind == "问问题":
        if day != date.today().isoformat():
            raise ValueError("问问题只能登记今天，不能通过切换日期重复加分")
        if question_recorded(data, name, day):
            raise ValueError("今天问问题已加过分，每人每天最多一次")
    event = {"id": uuid4().hex, "date": day, "type": kind,
             "points": ACTIVITIES[kind], "recorded_at": datetime.now().isoformat(timespec="seconds")}
    data.setdefault("activities", {}).setdefault(name, []).append(event)
    return event
