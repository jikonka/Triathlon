#!/usr/bin/env python3
"""
生成每周 Garmin FIT 训练计划文件
每周日运行，为下周（周一-周日）生成7个 .fit 文件
可直接导入 Garmin Connect

用法：python3 generate_weekly_fit.py --week 2026-03-23
"""

import time
import os
import sys
from datetime import date, timedelta

try:
    from fit_tool.fit_file_builder import FitFileBuilder
    from fit_tool.profile.messages.file_id_message import FileIdMessage
    from fit_tool.profile.messages.workout_message import WorkoutMessage
    from fit_tool.profile.messages.workout_step_message import WorkoutStepMessage
    from fit_tool.profile.profile_type import (
        FileType, Manufacturer, Sport,
        WorkoutStepDuration, WorkoutStepTarget, Intensity
    )
except ImportError:
    print("请安装: pip install fit-tool")
    sys.exit(1)

OUTPUT_DIR = "/root/.openclaw/workspace/training/weekly_fit"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ── 心率区间（基于HRmax=185） ──────────────────────────────
HR_ZONES = {
    "Z1": (0,   111),
    "Z2": (111, 130),
    "Z3": (130, 148),
    "Z4": (148, 167),
    "Z5": (167, 185),
    "Z2-Z3": (111, 148),
    "Z3-Z4": (130, 167),
}

SPORT_MAP = {
    "run":   Sport.RUNNING,
    "swim":  Sport.SWIMMING,
    "bike":  Sport.CYCLING,
    "brick": Sport.CYCLING,
}

def ms(seconds):
    """seconds -> milliseconds"""
    return int(seconds * 1000)

def min_to_ms(minutes):
    return ms(minutes * 60)

def build_fit(sport_type, workout_name, steps, out_path):
    """
    steps: list of dicts with keys:
      name, duration_min, hr_zone (str like 'Z2'), repeat (optional)
      OR for intervals: {type:'interval', reps:N, work:{...}, rest:{...}}
    """
    builder = FitFileBuilder()

    fid = FileIdMessage()
    fid.type = FileType.WORKOUT
    fid.manufacturer = Manufacturer.GARMIN.value
    fid.time_created = int(time.time() * 1000)
    builder.add(fid)

    # Count total steps (expanding intervals)
    total = 0
    for s in steps:
        if s.get("type") == "interval":
            total += s["reps"] * 2  # work + rest per rep
        else:
            total += 1

    wk = WorkoutMessage()
    wk.sport = SPORT_MAP.get(sport_type, Sport.RUNNING)
    wk.num_valid_steps = total
    wk.workout_name = workout_name[:16]  # Garmin max 16 chars
    builder.add(wk)

    step_idx = 0
    for s in steps:
        if s.get("type") == "interval":
            for rep in range(s["reps"]):
                # Work step
                w = s["work"]
                ws = WorkoutStepMessage()
                ws.message_index = step_idx
                ws.workout_step_name = f"{w.get('name','Work')}"[:16]
                ws.intensity = Intensity.ACTIVE
                ws.duration_type = WorkoutStepDuration.TIME
                ws.duration_time = min_to_ms(w["duration_min"])
                lo, hi = HR_ZONES.get(w["hr_zone"], (130, 167))
                ws.target_type = WorkoutStepTarget.HEART_RATE
                ws.custom_target_value_low = lo
                ws.custom_target_value_high = hi
                builder.add(ws)
                step_idx += 1

                # Rest step
                r = s["rest"]
                rs = WorkoutStepMessage()
                rs.message_index = step_idx
                rs.workout_step_name = "Rest"
                rs.intensity = Intensity.REST
                rs.duration_type = WorkoutStepDuration.TIME
                rs.duration_time = min_to_ms(r["duration_min"])
                rs.target_type = WorkoutStepTarget.OPEN
                builder.add(rs)
                step_idx += 1
        else:
            ws = WorkoutStepMessage()
            ws.message_index = step_idx
            ws.workout_step_name = s.get("name", "Step")[:16]
            intensity_str = s.get("intensity", "active")
            if intensity_str == "rest":
                ws.intensity = Intensity.REST
                ws.duration_type = WorkoutStepDuration.OPEN
                ws.target_type = WorkoutStepTarget.OPEN
            else:
                ws.intensity = Intensity.ACTIVE
                ws.duration_type = WorkoutStepDuration.TIME
                ws.duration_time = min_to_ms(s["duration_min"])
                lo, hi = HR_ZONES.get(s.get("hr_zone", "Z2"), (111, 130))
                ws.target_type = WorkoutStepTarget.HEART_RATE
                ws.custom_target_value_low = lo
                ws.custom_target_value_high = hi
            builder.add(ws)
            step_idx += 1

    fit = builder.build()
    fit.to_file(out_path)
    print(f"  ✅ {os.path.basename(out_path)}")


def generate_week(week_monday: date, week_plan: list):
    """
    week_plan: list of 7 dicts (Mon-Sun), each with:
      sport, name, steps (or None for rest)
    """
    day_names = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
    cn_names  = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]
    generated = []

    for i, day in enumerate(week_plan):
        d = week_monday + timedelta(days=i)
        label = f"{d.strftime('%Y-%m-%d')}_{day_names[i]}_{cn_names[i]}"

        if day is None or day.get("sport") == "rest":
            print(f"  ⏸  {label} — 休息日")
            continue

        fname = f"{label}_{day['name'][:20].replace(' ','_')}.fit"
        fpath = os.path.join(OUTPUT_DIR, fname)
        build_fit(day["sport"], day["name"], day["steps"], fpath)
        generated.append(fpath)

    return generated


# ════════════════════════════════════════════════════════════
# W1 训练计划定义（2026-03-23 开始）
# 每周更新此处
# ════════════════════════════════════════════════════════════

W1_PLAN = [
    # 周一：节奏跑 6km
    {
        "sport": "run", "name": "节奏跑6km",
        "steps": [
            {"name": "热身慢跑",    "duration_min": 8,  "hr_zone": "Z2"},
            {"name": "节奏跑Z3",    "duration_min": 24, "hr_zone": "Z3"},   # ~4km@6:20-6:40
            {"name": "放松慢跑",    "duration_min": 8,  "hr_zone": "Z2"},
        ]
    },
    # 周二：游泳 4×200m + 200m计时
    {
        "sport": "swim", "name": "游泳间歇4x200",
        "steps": [
            {"name": "热身",        "duration_min": 5,  "hr_zone": "Z2"},
            {"type": "interval", "reps": 4,
             "work": {"name": "200m快速",  "duration_min": 4,   "hr_zone": "Z3"},
             "rest": {"duration_min": 0.5}},
            {"name": "计时200m",    "duration_min": 4,  "hr_zone": "Z3-Z4"},
            {"name": "放松",        "duration_min": 3,  "hr_zone": "Z2"},
        ]
    },
    # 周三：阈值间歇 3×1km
    {
        "sport": "run", "name": "阈值间歇3x1km",
        "steps": [
            {"name": "热身慢跑",    "duration_min": 10, "hr_zone": "Z2"},
            {"type": "interval", "reps": 3,
             "work": {"name": "1km阈值",   "duration_min": 6,   "hr_zone": "Z4"},
             "rest": {"duration_min": 1.5}},
            {"name": "放松慢跑",    "duration_min": 5,  "hr_zone": "Z2"},
        ]
    },
    # 周四：休息
    None,
    # 周五：轻松跑 6km
    {
        "sport": "run", "name": "轻松跑6km",
        "steps": [
            {"name": "轻松跑Z2",    "duration_min": 42, "hr_zone": "Z2"},
        ]
    },
    # 周六：Brick 骑30km→跑5km
    {
        "sport": "brick", "name": "Brick骑30跑5",
        "steps": [
            {"name": "热身骑行",    "duration_min": 10, "hr_zone": "Z2"},
            {"name": "强度骑行Z3",  "duration_min": 60, "hr_zone": "Z3-Z4"},  # ~30km
            {"name": "换跑鞋过渡",  "duration_min": 2,  "hr_zone": "Z2"},
            {"name": "砖训跑步",    "duration_min": 35, "hr_zone": "Z3"},     # ~5km
        ]
    },
    # 周日：750m基准计时
    {
        "sport": "swim", "name": "750m基准计时",
        "steps": [
            {"name": "热身",        "duration_min": 5,  "hr_zone": "Z2"},
            {"name": "750m全力游",  "duration_min": 23, "hr_zone": "Z3-Z4"},
            {"name": "放松",        "duration_min": 3,  "hr_zone": "Z2"},
        ]
    },
]

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--week", default="2026-03-23",
                        help="本周周一日期，格式 YYYY-MM-DD")
    args = parser.parse_args()

    monday = date.fromisoformat(args.week)
    print(f"\n生成训练计划 FIT 文件：{monday} ～ {monday+timedelta(days=6)}")
    print(f"输出目录：{OUTPUT_DIR}\n")

    files = generate_week(monday, W1_PLAN)
    print(f"\n共生成 {len(files)} 个训练文件")
    print("请在 Garmin Connect App → 训练计划 → 导入 .fit 文件")
