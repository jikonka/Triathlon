#!/usr/bin/env python3
"""
Garmin FIT Workout 文件生成器
基于 Garmin 官方 FIT SDK 文档: https://developer.garmin.com/fit/cookbook/encoding-workout-files/

关键规则（来自官方文档）:
1. 文件结构: FileId → Workout → WorkoutStep × N
2. 自定义心率 BPM 必须 +100 偏移 (e.g., 135bpm → 235, 148bpm → 248)
3. 时间单位: duration_value 为毫秒; duration_time 子字段为秒
4. 距离单位: duration_value 为厘米; duration_distance 子字段为米
5. 游泳: sub_sport=LapSwimming, pool_length 必填, 步骤用 Distance 而非 Time
6. Repeat 步骤: duration_type=RepeatUntilStepsCrit(6), duration_value=起始step_index
7. 心率区间: target_type=HeartRate(1), target_value=区间编号(1-5);
   自定义心率: target_value=0, custom_target_low=bpm+100, custom_target_high=bpm+100
"""

import os, time, sys
from datetime import date, timedelta

try:
    from fit_tool.fit_file_builder import FitFileBuilder
    from fit_tool.profile.messages.file_id_message import FileIdMessage
    from fit_tool.profile.messages.workout_message import WorkoutMessage
    from fit_tool.profile.messages.workout_step_message import WorkoutStepMessage
    from fit_tool.profile.profile_type import (
        FileType, Manufacturer, Sport, SubSport, DisplayMeasure,
        WorkoutStepDuration, WorkoutStepTarget, Intensity
    )
except ImportError:
    print("请安装: pip install fit-tool")
    sys.exit(1)

OUTPUT_DIR = "/root/.openclaw/workspace/training/weekly_fit"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ── 心率区间 Garmin 官方 Zone 编号（基于 Tactix 7 设置）──────────
# Zone 1: <111, Zone 2: 111-130, Zone 3: 130-148, Zone 4: 148-167, Zone 5: 167-185
# 使用自定义绝对 BPM 值（更精确，不依赖设备区间设置）

def hr_offset(bpm):
    """官方规定：自定义心率绝对值必须 +100"""
    return bpm + 100

# ── 工具函数 ────────────────────────────────────────────────────

def _file_id():
    msg = FileIdMessage()
    msg.type = FileType.WORKOUT
    msg.manufacturer = Manufacturer.GARMIN.value
    msg.time_created = int(time.time() * 1000)
    return msg

def _workout_msg(name, sport, sub_sport=None, num_steps=0,
                  pool_length=None, pool_length_unit=None):
    """
    官方要求: Workout 消息是文件第二条消息
    game: pool_length 单位米, pool_length_unit: DisplayMeasure.Metric=0
    """
    msg = WorkoutMessage()
    msg.sport = sport
    if sub_sport is not None:
        msg.sub_sport = sub_sport
    msg.num_valid_steps = num_steps
    msg.workout_name = name[:16]
    if pool_length is not None:
        msg.pool_length = pool_length       # meters (float)
    if pool_length_unit is not None:
        msg.pool_length_unit = pool_length_unit
    return msg

def step_active_time(idx, name, duration_sec, hr_low_bpm, hr_high_bpm,
                      intensity=None):
    """
    时间步骤（跑步/骑行用）
    duration_value 单位：毫秒（官方：Time → ms）
    心率自定义值：BPM + 100 偏移
    """
    msg = WorkoutStepMessage()
    msg.message_index = idx
    msg.workout_step_name = name[:16]
    msg.intensity = intensity if intensity is not None else Intensity.ACTIVE
    msg.duration_type = WorkoutStepDuration.TIME
    msg.duration_value = int(duration_sec * 1000)   # 秒 → 毫秒
    msg.target_type = WorkoutStepTarget.HEART_RATE
    msg.target_value = 0                             # 0 = 使用自定义值
    msg.custom_target_value_low  = hr_offset(hr_low_bpm)
    msg.custom_target_value_high = hr_offset(hr_high_bpm)
    return msg

def step_active_dist(idx, name, distance_m, hr_low_bpm=None, hr_high_bpm=None,
                      intensity=None):
    """
    距离步骤（游泳用）
    duration_value 单位：厘米（官方：Distance → cm）
    """
    msg = WorkoutStepMessage()
    msg.message_index = idx
    msg.workout_step_name = name[:16]
    msg.intensity = intensity if intensity is not None else Intensity.ACTIVE
    msg.duration_type = WorkoutStepDuration.DISTANCE
    msg.duration_value = int(distance_m * 100)   # 米 → 厘米
    if hr_low_bpm and hr_high_bpm:
        msg.target_type = WorkoutStepTarget.HEART_RATE
        msg.target_value = 0
        msg.custom_target_value_low  = hr_offset(hr_low_bpm)
        msg.custom_target_value_high = hr_offset(hr_high_bpm)
    else:
        msg.target_type = WorkoutStepTarget.OPEN
    return msg

def step_rest_time(idx, duration_sec):
    """固定时间休息步骤"""
    msg = WorkoutStepMessage()
    msg.message_index = idx
    msg.workout_step_name = "Rest"
    msg.intensity = Intensity.REST
    msg.duration_type = WorkoutStepDuration.TIME
    msg.duration_value = int(duration_sec * 1000)
    msg.target_type = WorkoutStepTarget.OPEN
    return msg

def step_rest_open(idx):
    """开放式休息（按 Lap 键继续）"""
    msg = WorkoutStepMessage()
    msg.message_index = idx
    msg.workout_step_name = "Rest"
    msg.intensity = Intensity.REST
    msg.duration_type = WorkoutStepDuration.OPEN
    msg.target_type = WorkoutStepTarget.OPEN
    return msg

def step_repeat(idx, repeat_from_idx, repetitions):
    """
    Repeat 步骤（官方: duration_type=RepeatUntilStepsCrit=6）
    repeat_from_idx: 重复块起始 step 的 message_index
    """
    msg = WorkoutStepMessage()
    msg.message_index = idx
    msg.duration_type = WorkoutStepDuration.REPEAT_UNTIL_STEPS_CMPLT
    msg.duration_value = repeat_from_idx   # 重复从哪个 step 开始
    msg.target_type = WorkoutStepTarget.OPEN
    msg.target_value = repetitions
    return msg

def build_fit_file(sport, sub_sport, workout_name, steps,
                    pool_length=None, pool_length_unit=None):
    """构建并返回 FitFile 对象"""
    builder = FitFileBuilder()

    # 1. File ID（必须是第一条消息）
    builder.add(_file_id())

    # 2. Workout（必须是第二条消息）
    num = len(steps)
    wk = _workout_msg(workout_name, sport, sub_sport, num,
                      pool_length, pool_length_unit)
    builder.add(wk)

    # 3. WorkoutStep 消息（按顺序）
    for s in steps:
        builder.add(s)

    return builder.build()

def save(fit_file, path):
    fit_file.to_file(path)
    size = os.path.getsize(path)
    print(f"  ✅ {os.path.basename(path)} ({size} bytes)")
    return path

# ════════════════════════════════════════════════════════════════
# W1 每日训练文件（2026-03-23 起）
# ════════════════════════════════════════════════════════════════

def w1_mon_tempo_run():
    """周一：节奏跑 6km — 热身Z2 8min + 节奏Z3 24min + 放松Z2 8min"""
    steps = []
    steps.append(step_active_time(0, "热身Z2",   8*60,  111, 130, Intensity.WARMUP))
    steps.append(step_active_time(1, "节奏跑Z3", 24*60,  130, 148, Intensity.ACTIVE))
    steps.append(step_active_time(2, "放松Z2",   8*60,  111, 130, Intensity.COOLDOWN))
    return build_fit_file(Sport.RUNNING, None, "节奏跑6km", steps)

def w1_tue_swim_intervals():
    """周二：游泳 4×200m间歇 + 200m计时
    游泳用距离步骤，pool_length=25m
    """
    steps = []
    # 热身 200m
    steps.append(step_active_dist(0, "热身200m", 200, 111, 130, Intensity.WARMUP))
    steps.append(step_rest_open(1))

    # 4×200m间歇（用repeat step）
    # step 2: 200m主动
    steps.append(step_active_dist(2, "200m Z3", 200, 130, 148))
    # step 3: 30s休息
    steps.append(step_rest_time(3, 30))
    # step 4: repeat 步骤，从 step 2 开始，重复 4 次
    steps.append(step_repeat(4, repeat_from_idx=2, repetitions=4))

    # 200m全力计时
    steps.append(step_active_dist(5, "200m计时", 200, 130, 167))
    steps.append(step_rest_open(6))

    # 放松 100m
    steps.append(step_active_dist(7, "放松100m", 100, 111, 130, Intensity.COOLDOWN))

    return build_fit_file(
        Sport.SWIMMING, SubSport.LAP_SWIMMING, "游泳间歇4x200",
        steps, pool_length=25.0, pool_length_unit=DisplayMeasure.METRIC
    )

def w1_wed_threshold():
    """周三：阈值间歇 3×1km @Z4 — 热身10min + 3×(1km@Z4 + 90s休息) + 放松5min"""
    steps = []
    steps.append(step_active_time(0, "热身",      10*60, 111, 130, Intensity.WARMUP))

    # step 1: 1km阈值 （用时间近似：Z4配速5:45~6:05/km ≈ 6min）
    steps.append(step_active_time(1, "1km阈值Z4", 6*60,  148, 167))
    # step 2: 90s慢跑恢复
    steps.append(step_rest_time(2, 90))
    # step 3: repeat，从step1开始，重复3次
    steps.append(step_repeat(3, repeat_from_idx=1, repetitions=3))

    steps.append(step_active_time(4, "放松",       5*60, 111, 130, Intensity.COOLDOWN))
    return build_fit_file(Sport.RUNNING, None, "阈值间歇3x1km", steps)

def w1_fri_easy_run():
    """周五：轻松跑 6km @Z2 — 全程42min"""
    steps = []
    steps.append(step_active_time(0, "轻松跑Z2", 42*60, 111, 130))
    return build_fit_file(Sport.RUNNING, None, "轻松跑6km", steps)

def w1_sat_brick():
    """周六：Brick — 骑30km@Z3-Z4 + T1 + 跑5km@Z3
    骑行: 热身10min@Z2 + 强度60min@Z3-Z4 + 冲刺10min@Z4
    跑步: 35min@Z3
    注：FIT格式骑行部分用 Cycling，T1后的跑步实际上手动切换项目
    这里输出骑行workout；跑步作为单独FIT文件
    """
    steps = []
    steps.append(step_active_time(0, "热身骑行",  10*60, 111, 130, Intensity.WARMUP))
    steps.append(step_active_time(1, "强度Z3-Z4", 60*60, 130, 155))
    steps.append(step_active_time(2, "冲刺Z4",    10*60, 148, 160))
    steps.append(step_active_time(3, "放松",       5*60, 111, 130, Intensity.COOLDOWN))
    return build_fit_file(Sport.CYCLING, None, "Brick骑行30km", steps)

def w1_sat_brick_run():
    """周六Brick跑步部分（骑后换项跑）"""
    steps = []
    steps.append(step_active_time(0, "砖训跑Z3", 35*60, 130, 148))
    return build_fit_file(Sport.RUNNING, None, "Brick跑步5km", steps)

def w1_sun_swim_750():
    """周日：750m基准计时 — 热身 + 750m全力 + 放松"""
    steps = []
    steps.append(step_active_dist(0, "热身200m",  200,  111, 130, Intensity.WARMUP))
    steps.append(step_rest_open(1))
    steps.append(step_active_dist(2, "750m计时",  750,  130, 167))
    steps.append(step_rest_open(3))
    steps.append(step_active_dist(4, "放松100m",  100,  111, 130, Intensity.COOLDOWN))
    return build_fit_file(
        Sport.SWIMMING, SubSport.LAP_SWIMMING, "750m基准计时",
        steps, pool_length=25.0, pool_length_unit=DisplayMeasure.METRIC
    )

# ════════════════════════════════════════════════════════════════
# 主函数：生成 W1 所有文件
# ════════════════════════════════════════════════════════════════

def generate_w1():
    monday = date(2026, 3, 23)
    print(f"\n生成 W1 训练FIT文件 ({monday} ～ {monday+timedelta(days=6)})")
    print(f"输出目录: {OUTPUT_DIR}\n")

    files = [
        (monday+timedelta(0), "Mon_周一", w1_mon_tempo_run()),
        (monday+timedelta(1), "Tue_周二", w1_tue_swim_intervals()),
        (monday+timedelta(2), "Wed_周三", w1_wed_threshold()),
        # 周四休息
        (monday+timedelta(4), "Fri_周五", w1_fri_easy_run()),
        (monday+timedelta(5), "Sat_周六_骑行", w1_sat_brick()),
        (monday+timedelta(5), "Sat_周六_跑步", w1_sat_brick_run()),
        (monday+timedelta(6), "Sun_周日", w1_sun_swim_750()),
    ]

    saved = []
    for d, label, fit in files:
        fname = f"{d}_{label}.fit"
        path = os.path.join(OUTPUT_DIR, fname)
        save(fit, path)
        saved.append(path)

    print(f"\n共生成 {len(saved)} 个文件")
    print("导入方式: Garmin Connect → 训练计划 → 导入 .fit 文件")
    print("或上传至: https://connect.garmin.cn/app/import-data")
    return saved

if __name__ == "__main__":
    generate_w1()
