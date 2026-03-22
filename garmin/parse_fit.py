#!/usr/bin/env python3
"""
Garmin FIT 文件批量解析脚本 v2
- 支持 multisport：每个 session 单独一行记录
- transition session 单独标记，不计入三项运动统计
解析 /root/.openclaw/workspace/garmin/raw/ 下所有 .fit 文件
输出 activities_summary.csv 和 weekly_stats.md
"""

import os
import csv
import glob
from datetime import datetime, timezone, timedelta
from collections import defaultdict

try:
    from fitparse import FitFile
except ImportError:
    print("请先安装 fitparse: python3 -m pip install fitparse")
    exit(1)

RAW_DIR = "/root/.openclaw/workspace/garmin/raw"
OUTPUT_CSV = "/root/.openclaw/workspace/garmin/activities_summary.csv"
OUTPUT_WEEKLY = "/root/.openclaw/workspace/garmin/weekly_stats.md"
TZ_BEIJING = timezone(timedelta(hours=8))

SPORT_MAP = {
    "running": "跑步",
    "cycling": "骑行",
    "swimming": "游泳",
    "walking": "步行",
    "training": "训练",
    "transition": "换项",
    "multisport": "铁三",
    "generic": "其他",
    "alpine_skiing": "高山滑雪",
    "snowboarding": "单板滑雪",
}

def to_beijing(dt):
    if dt is None:
        return None
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(TZ_BEIJING)

def seconds_to_minutes(s):
    if s is None:
        return ""
    return round(s / 60, 1)

def speed_to_pace(speed_ms, sport):
    """speed in m/s -> min/km string"""
    if not speed_ms or speed_ms <= 0:
        return ""
    if sport not in ("跑步", "游泳"):
        return ""
    pace_sec_per_km = 1000 / speed_ms
    minutes = int(pace_sec_per_km // 60)
    seconds = int(pace_sec_per_km % 60)
    return f"{minutes}:{seconds:02d}"

def parse_fit_file(filepath):
    """
    解析一个 FIT 文件，返回一个或多个 session 记录列表。
    multisport 文件会返回多条记录（每个子项目一条）。
    """
    results = []
    filename = os.path.basename(filepath)

    try:
        fitfile = FitFile(filepath)
        sessions = list(fitfile.get_messages("session"))
        sports_list = list(fitfile.get_messages("sport"))

        # Build sport map by index
        sport_by_index = {}
        for i, srec in enumerate(sports_list):
            sdata = {f.name: f.value for f in srec.fields}
            sport_by_index[i] = sdata

        is_multisport = len(sessions) > 1

        for idx, session_rec in enumerate(sessions):
            data = {f.name: f.value for f in session_rec.fields}

            row = {
                "filename": filename,
                "multisport": "是" if is_multisport else "否",
                "session_index": idx + 1 if is_multisport else "",
                "date": "",
                "start_time": "",
                "sport": "",
                "sub_sport": "",
                "total_elapsed_time": "",
                "total_timer_time": "",
                "total_distance": "",
                "avg_heart_rate": "",
                "max_heart_rate": "",
                "avg_speed": "",
                "avg_pace": "",
                "total_calories": "",
                "total_ascent": "",
                "avg_cadence": "",
                "training_load": "",
                "aerobic_training_effect": "",
                "anaerobic_training_effect": "",
                "normalized_power": "",
                "avg_power": "",
                "pool_length": "",
                "total_strokes": "",
            }

            # Time
            start_raw = data.get("start_time")
            if start_raw:
                bj = to_beijing(start_raw)
                row["date"] = bj.strftime("%Y-%m-%d")
                row["start_time"] = bj.strftime("%H:%M")

            # Sport: prefer session's own sport field
            sport_raw = str(data.get("sport", "")).lower()
            sub_raw = str(data.get("sub_sport", "") or "")
            # fallback to sport message by index
            if not sport_raw or sport_raw in ("none", "generic") and idx in sport_by_index:
                sdata = sport_by_index.get(idx, {})
                sport_raw = str(sdata.get("sport", sport_raw)).lower()
                sub_raw = str(sdata.get("sub_sport", sub_raw) or "")

            row["sport"] = SPORT_MAP.get(sport_raw, sport_raw or "其他")
            row["sub_sport"] = sub_raw if sub_raw not in ("None", "", "generic") else ""

            # Duration
            row["total_elapsed_time"] = seconds_to_minutes(data.get("total_elapsed_time"))
            row["total_timer_time"] = seconds_to_minutes(data.get("total_timer_time"))

            # Distance
            dist = data.get("total_distance")
            if dist is not None:
                dist_f = float(dist)
                sport = row["sport"]
                if sport == "游泳":
                    row["total_distance"] = round(dist_f, 0)  # meters
                else:
                    # Fix: some treadmill sessions store in meters (>500 but no speed)
                    avg_speed = data.get("avg_speed")
                    if dist_f > 500 and (not avg_speed or float(avg_speed or 0) == 0):
                        row["total_distance"] = round(dist_f / 1000, 2)
                    else:
                        row["total_distance"] = round(dist_f / 1000, 2)

            # Heart rate
            row["avg_heart_rate"] = data.get("avg_heart_rate") or ""
            row["max_heart_rate"] = data.get("max_heart_rate") or ""

            # Speed & pace
            avg_speed = data.get("avg_speed")
            if avg_speed:
                row["avg_speed"] = round(float(avg_speed) * 3.6, 2)
                row["avg_pace"] = speed_to_pace(float(avg_speed), row["sport"])
            else:
                # Calculate pace from distance and time if possible
                if row["total_distance"] and row["total_timer_time"] and row["sport"] in ("跑步", "游泳"):
                    try:
                        dist_km = float(row["total_distance"])
                        time_min = float(row["total_timer_time"])
                        if dist_km > 0 and time_min > 0:
                            pace_sec = (time_min * 60) / dist_km
                            mins = int(pace_sec // 60)
                            secs = int(pace_sec % 60)
                            row["avg_pace"] = f"{mins}:{secs:02d}"
                    except:
                        pass

            # Other metrics
            row["total_calories"] = data.get("total_calories") or ""
            row["total_ascent"] = data.get("total_ascent") or ""
            row["avg_cadence"] = data.get("avg_cadence") or ""
            row["normalized_power"] = data.get("normalized_power") or ""
            row["avg_power"] = data.get("avg_power") or ""
            row["pool_length"] = data.get("pool_length") or ""
            row["total_strokes"] = data.get("total_strokes") or ""
            row["aerobic_training_effect"] = data.get("total_training_effect") or ""
            row["anaerobic_training_effect"] = data.get("total_anaerobic_training_effect") or ""
            row["training_load"] = data.get("training_load_peak") or ""

            results.append(row)

    except Exception as e:
        results.append({
            "filename": filename,
            "multisport": "",
            "session_index": "",
            "date": "",
            "start_time": "",
            "sport": f"ERROR: {e}",
            **{k: "" for k in ["sub_sport","total_elapsed_time","total_timer_time",
                               "total_distance","avg_heart_rate","max_heart_rate",
                               "avg_speed","avg_pace","total_calories","total_ascent",
                               "avg_cadence","training_load","aerobic_training_effect",
                               "anaerobic_training_effect","normalized_power","avg_power",
                               "pool_length","total_strokes"]}
        })

    return results


def main():
    fit_files = sorted([
        f for f in glob.glob(os.path.join(RAW_DIR, "**", "*.fit"), recursive=True)
        if "__MACOSX" not in f
    ])

    print(f"找到 {len(fit_files)} 个 FIT 文件，开始解析（含 multisport 子项目）...")

    all_rows = []
    errors = []
    sport_counts = defaultdict(int)
    file_count = 0

    for i, fp in enumerate(fit_files):
        if (i + 1) % 20 == 0:
            print(f"  进度: {i+1}/{len(fit_files)}")
        rows = parse_fit_file(fp)
        file_count += 1
        for row in rows:
            if row["sport"].startswith("ERROR"):
                errors.append((fp, row["sport"]))
            else:
                sport_counts[row["sport"]] += 1
            all_rows.append(row)

    # Sort by date then start_time then session_index
    all_rows.sort(key=lambda x: (x["date"], x["start_time"], str(x["session_index"])))

    # Write CSV
    fieldnames = [
        "filename", "multisport", "session_index",
        "date", "start_time", "sport", "sub_sport",
        "total_elapsed_time", "total_timer_time", "total_distance",
        "avg_heart_rate", "max_heart_rate", "avg_speed", "avg_pace",
        "total_calories", "total_ascent", "avg_cadence", "training_load",
        "aerobic_training_effect", "anaerobic_training_effect",
        "normalized_power", "avg_power", "pool_length", "total_strokes"
    ]

    with open(OUTPUT_CSV, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(all_rows)

    print(f"\n✅ CSV 已保存: {OUTPUT_CSV}")
    print(f"   FIT文件: {file_count} 个，总记录: {len(all_rows)} 条，解析错误: {len(errors)} 个")
    print("   运动类型分布（含multisport子项目）:")
    for sport, cnt in sorted(sport_counts.items(), key=lambda x: -x[1]):
        print(f"     {sport}: {cnt} 次")

    if errors:
        print("\n⚠️  解析失败:")
        for fp, err in errors:
            print(f"   {os.path.basename(fp)}: {err}")

    generate_weekly_stats(all_rows)
    print(f"\n✅ 周统计已保存: {OUTPUT_WEEKLY}")


def generate_weekly_stats(all_rows):
    from datetime import date

    def get_week_start(date_str):
        if not date_str:
            return None
        d = datetime.strptime(date_str, "%Y-%m-%d").date()
        return d - timedelta(days=d.weekday())

    # Exclude transition sessions from stats
    stat_rows = [r for r in all_rows if r["sport"] not in ("换项", "其他", "ERROR") and not str(r["sport"]).startswith("ERROR")]

    weeks = defaultdict(lambda: {
        "total_time_min": 0.0,
        "sport_time": defaultdict(float),
        "sport_count": defaultdict(int),
        "heart_rates": [],
        "run_dist_km": 0.0,
        "bike_dist_km": 0.0,
        "swim_dist_m": 0.0,
    })

    for a in stat_rows:
        week_start = get_week_start(a["date"])
        if week_start is None:
            continue
        w = weeks[week_start]
        time_min = float(a["total_timer_time"]) if a["total_timer_time"] else 0
        w["total_time_min"] += time_min
        sport = a["sport"]
        w["sport_time"][sport] += time_min
        w["sport_count"][sport] += 1
        if a["avg_heart_rate"]:
            try:
                w["heart_rates"].append(float(a["avg_heart_rate"]))
            except:
                pass
        if a["total_distance"]:
            try:
                dist = float(a["total_distance"])
                if sport == "跑步":
                    w["run_dist_km"] += dist
                elif sport == "骑行":
                    w["bike_dist_km"] += dist
                elif sport == "游泳":
                    w["swim_dist_m"] += dist
            except:
                pass

    lines = ["# Garmin 周训练统计\n", f"_生成时间：{datetime.now(TZ_BEIJING).strftime('%Y-%m-%d %H:%M')}_\n"]

    for week_start in sorted(weeks.keys()):
        week_end = week_start + timedelta(days=6)
        w = weeks[week_start]
        total_sessions = sum(w["sport_count"].values())
        total_time = w["total_time_min"]
        avg_hr = round(sum(w["heart_rates"]) / len(w["heart_rates"]), 0) if w["heart_rates"] else "-"

        lines.append(f"\n## {week_start} ～ {week_end}")
        lines.append(f"- **总训练次数：** {total_sessions} 次")
        lines.append(f"- **总训练时长：** {round(total_time, 0)} 分钟（{round(total_time/60, 1)} 小时）")
        lines.append(f"- **平均心率：** {avg_hr} bpm")
        if w["run_dist_km"] > 0:
            lines.append(f"- **跑步距离：** {round(w['run_dist_km'], 2)} km")
        if w["bike_dist_km"] > 0:
            lines.append(f"- **骑行距离：** {round(w['bike_dist_km'], 2)} km")
        if w["swim_dist_m"] > 0:
            lines.append(f"- **游泳距离：** {round(w['swim_dist_m'], 0)} m")
        lines.append("- **各项目明细：**")
        for sport, cnt in sorted(w["sport_count"].items(), key=lambda x: -x[1]):
            time_h = round(w["sport_time"][sport] / 60, 1)
            lines.append(f"  - {sport}：{cnt} 次，{round(w['sport_time'][sport], 0)} 分钟（{time_h}h）")

    with open(OUTPUT_WEEKLY, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))


if __name__ == "__main__":
    main()
