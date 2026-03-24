# Training Files

## 文件说明

- `current_plan.md` — 当前有效训练计划（周视图）
- `weekly_log.md` — 每周执行记录与调整日志
- `phase_overview.md` — 备赛阶段目标与里程碑

## 维护原则

1. 训练计划由 AI 主动维护，聊天中的任何修改即时同步到 current_plan.md
2. Zikun 提供的 Garmin 数据解析后存入 ../garmin/ 目录
3. 每次对话前读取此目录，确保计划连续一致
