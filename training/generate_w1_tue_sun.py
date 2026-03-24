#!/usr/bin/env python3
"""
生成本周（2026-03-24 周二 到 2026-03-29 周日）Garmin workout FIT 文件
最大心率 185bpm，乳酸阈值 172bpm

心率区间：
  Z2: 111-130  轻松有氧
  Z3: 130-148  有氧强化
  Z4: 148-167  阈值/节奏
  Z5: 167-185  无氧

用法：python3 generate_w1_tue_sun.py
"""

import os
import sys
import time
from datetime import date

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
    print("fit-tool 未安装，正在安装…")
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "fit-tool"])
    from fit_tool.fit_file_builder import FitFileBuilder
    from fit_tool.profile.messages.file_id_message import FileIdMessage
    from fit_tool.profile.messages.workout_message import WorkoutMessage
    from fit_tool.profile.messages.workout_step_message import WorkoutStepMessage
    from fit_tool.profile.profile_type import (
        FileType, Manufacturer, Sport,
        WorkoutStepDuration, WorkoutStepTarget, Intensity
    )

OUTPUT_DIR = "/root/.openclaw/workspace/training/weekly_fit"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ── 心率区间（HRmax=185） ──────────────────────────────────
HR_ZONES = {
    "Z2":      (111, 130),
    "Z3":      (130, 148),
    "Z4":      (148, 167),
    "Z5":      (167, 185),
    "Z2-Z3":   (111, 148),
    "Z3-Z4":   (130, 167),
    "Z4_pace": (155, 165),   # 节奏跑乳酸阈值附近
    "Z4_brick":(150, 160),   # 骑行冲刺 / 砖式
    "Z4_run":  (155, 165),   # 砖式跑稳定段
    "Z3_bike": (140, 152),   # 骑行有氧
    "Z3_bike2":(140, 150),   # 骑行节奏
    "Z4_bike": (150, 160),   # 骑行强化
    "Z3_swim": (130, 140),   # 游泳有氧
    "Z34_swim":(140, 155),   # 游泳间歇
    "Z4_adapt":(150, 160),   # 砖式前段适应
}

SPORT_MAP = {
    "RUNNING":  Sport.RUNNING,
    "SWIMMING": Sport.SWIMMING,
    "CYCLING":  Sport.CYCLING,
}


def min_to_ms(minutes: float) -> int:
    """分钟 → 毫秒"""
    return int(minutes * 60 * 1000)


def count_steps(steps: list) -> int:
    """统计展开后的步骤数（interval 展开为 reps×2）"""
    total = 0
    for s in steps:
        if s.get("type") == "interval":
            total += s["reps"] * 2
        else:
            total += 1
    return total


def build_fit(sport_key: str, workout_name: str, steps: list, out_path: str):
    """
    构建并写出 FIT 文件。

    steps 格式（两种）：
      普通步骤：
        {"name": str, "duration_min": float, "hr_zone": str, "intensity": "active"|"rest"}
      间歇展开：
        {"type": "interval", "reps": int,
         "work": {"name": str, "duration_min": float, "hr_zone": str},
         "rest": {"name": str, "duration_min": float}}
    """
    builder = FitFileBuilder()

    # File ID
    fid = FileIdMessage()
    fid.type = FileType.WORKOUT
    fid.manufacturer = Manufacturer.GARMIN.value
    fid.time_created = int(time.time() * 1000)
    builder.add(fid)

    # Workout header
    wk = WorkoutMessage()
    wk.sport = SPORT_MAP[sport_key]
    wk.num_valid_steps = count_steps(steps)
    wk.workout_name = workout_name[:16]
    builder.add(wk)

    idx = 0
    for s in steps:
        if s.get("type") == "interval":
            for _ in range(s["reps"]):
                # Work
                w = s["work"]
                lo, hi = HR_ZONES[w["hr_zone"]]
                ws = WorkoutStepMessage()
                ws.message_index = idx
                ws.workout_step_name = w.get("name", "Work")[:16]
                ws.intensity = Intensity.ACTIVE
                ws.duration_type = WorkoutStepDuration.TIME
                ws.duration_time = min_to_ms(w["duration_min"])
                ws.target_type = WorkoutStepTarget.HEART_RATE
                ws.custom_target_value_low = lo
                ws.custom_target_value_high = hi
                builder.add(ws)
                idx += 1

                # Rest
                r = s["rest"]
                rs = WorkoutStepMessage()
                rs.message_index = idx
                rs.workout_step_name = r.get("name", "Rest")[:16]
                rs.intensity = Intensity.REST
                rs.duration_type = WorkoutStepDuration.TIME
                rs.duration_time = min_to_ms(r["duration_min"])
                rs.target_type = WorkoutStepTarget.OPEN
                builder.add(rs)
                idx += 1
        else:
            # 普通步骤
            is_rest = s.get("intensity", "active") == "rest"
            lo, hi = HR_ZONES.get(s.get("hr_zone", "Z2"), (111, 130))

            ws = WorkoutStepMessage()
            ws.message_index = idx
            ws.workout_step_name = s.get("name", "Step")[:16]

            if is_rest:
                ws.intensity = Intensity.REST
                ws.duration_type = WorkoutStepDuration.OPEN
                ws.target_type = WorkoutStepTarget.OPEN
            else:
                ws.intensity = Intensity.ACTIVE
                ws.duration_type = WorkoutStepDuration.TIME
                ws.duration_time = min_to_ms(s["duration_min"])
                ws.target_type = WorkoutStepTarget.HEART_RATE
                ws.custom_target_value_low = lo
                ws.custom_target_value_high = hi

            builder.add(ws)
            idx += 1

    fit = builder.build()
    fit.to_file(out_path)
    print(f"  ✅  {os.path.basename(out_path)}")
    return out_path


# ═══════════════════════════════════════════════════════════
# 训练定义
# ═══════════════════════════════════════════════════════════

WORKOUTS = [
    # ── 周二 3/24：节奏跑 6km ──────────────────────────────
    {
        "filename": "2026-03-24_Tue_节奏跑.fit",
        "sport": "RUNNING",
        "name": "节奏跑6km",
        "steps": [
            {"name": "热身慢跑",   "duration_min": 8,  "hr_zone": "Z2"},
            {"name": "节奏跑Z4",   "duration_min": 22, "hr_zone": "Z4_pace"},
            {"name": "放松慢跑",   "duration_min": 6,  "hr_zone": "Z2"},
        ],
    },

    # ── 周三 3/25：游泳有氧耐力 1000m ──────────────────────
    {
        "filename": "2026-03-25_Wed_游泳耐力.fit",
        "sport": "SWIMMING",
        "name": "游泳耐力1000m",
        "steps": [
            {"name": "热身",      "duration_min": 5,  "hr_zone": "Z2"},
            # 4×100m：每组3min工作 + 20s(≈0.33min)休息
            {
                "type": "interval", "reps": 4,
                "work": {"name": "100m快泳",  "duration_min": 3,    "hr_zone": "Z3_swim"},
                "rest": {"name": "20s休息",   "duration_min": 0.334},
            },
            {"name": "300m连续",  "duration_min": 9,  "hr_zone": "Z3_swim"},
            {"name": "冷身",      "duration_min": 3,  "hr_zone": "Z2"},
        ],
    },

    # ── 周四 3/26：轻松有氧跑 6km ──────────────────────────
    {
        "filename": "2026-03-26_Thu_轻松跑.fit",
        "sport": "RUNNING",
        "name": "轻松跑6km",
        "steps": [
            {"name": "轻松跑Z2-Z3", "duration_min": 42, "hr_zone": "Z2-Z3"},
        ],
    },

    # ── 周五 3/27：游泳间歇强度 900m ───────────────────────
    {
        "filename": "2026-03-27_Fri_游泳间歇.fit",
        "sport": "SWIMMING",
        "name": "游泳间歇900m",
        "steps": [
            {"name": "热身",      "duration_min": 4,  "hr_zone": "Z2"},
            # 6×75m：每组2.5min工作 + 15s(≈0.25min)休息
            {
                "type": "interval", "reps": 6,
                "work": {"name": "75m间歇",  "duration_min": 2.5,  "hr_zone": "Z34_swim"},
                "rest": {"name": "15s休息",  "duration_min": 0.25},
            },
            {"name": "200m连续",  "duration_min": 6,  "hr_zone": "Z3_swim"},
            {"name": "冷身",      "duration_min": 3,  "hr_zone": "Z2"},
        ],
    },

    # ── 周六 3/28：节奏骑行 35km ───────────────────────────
    {
        "filename": "2026-03-28_Sat_节奏骑行.fit",
        "sport": "CYCLING",
        "name": "节奏骑行35km",
        "steps": [
            {"name": "热身",        "duration_min": 10, "hr_zone": "Z2"},
            {"name": "稳定节奏骑行", "duration_min": 50, "hr_zone": "Z3_bike2"},
            # 2×12min冲刺 + 3min Z2恢复（展开）
            {"name": "冲刺1 Z4",   "duration_min": 12, "hr_zone": "Z4_bike"},
            {"name": "恢复1 Z2",   "duration_min": 3,  "hr_zone": "Z2"},
            {"name": "冲刺2 Z4",   "duration_min": 12, "hr_zone": "Z4_bike"},
            {"name": "冷身",        "duration_min": 8,  "hr_zone": "Z2"},
        ],
    },

    # ── 周日 3/29：砖式训练 骑20km ─────────────────────────
    {
        "filename": "2026-03-29_Sun_砖式骑行.fit",
        "sport": "CYCLING",
        "name": "砖式骑20km",
        "steps": [
            {"name": "热身",        "duration_min": 8,  "hr_zone": "Z2"},
            {"name": "比赛强度骑行", "duration_min": 42, "hr_zone": "Z3_bike"},
            {"name": "最后冲刺",    "duration_min": 5,  "hr_zone": "Z4_brick"},
        ],
    },

    # ── 周日 3/29：砖式训练 跑5km ──────────────────────────
    {
        "filename": "2026-03-29_Sun_砖式跑步.fit",
        "sport": "RUNNING",
        "name": "砖式跑5km",
        "steps": [
            {"name": "前段适应Z4",  "duration_min": 10, "hr_zone": "Z4_adapt"},
            {"name": "稳定段Z4",   "duration_min": 22, "hr_zone": "Z4_run"},
        ],
    },
]


# ═══════════════════════════════════════════════════════════
# 主程序
# ═══════════════════════════════════════════════════════════

if __name__ == "__main__":
    print(f"\n{'='*56}")
    print(f"  生成本周训练 FIT 文件（2026-03-24 周二 ～ 2026-03-29 周日）")
    print(f"  输出目录：{OUTPUT_DIR}")
    print(f"{'='*56}\n")

    generated = []
    for wk in WORKOUTS:
        out_path = os.path.join(OUTPUT_DIR, wk["filename"])
        build_fit(wk["sport"], wk["name"], wk["steps"], out_path)
        generated.append(out_path)

    print(f"\n共生成 {len(generated)} 个 FIT 文件：")
    for f in generated:
        size = os.path.getsize(f)
        print(f"  {os.path.basename(f)}  ({size} bytes)")

    print("\n📲 导入方法：Garmin Connect App → 更多 → 训练与计划 → 训练 → 导入")
    print("   或 Garmin Connect Web → Training → Workouts → Import Workout\n")
