# Contributing

## 工作流

本 skill 遵循**双轨原则**：AI/人决定"引哪篇、引在哪"，脚本只做格式化。两者互不干扰。

- **外部工具（AI / 人工）**：分析正文内容 → 产出映射表（`--smap`）或写入 `[N]` 手打标记（`--replace`）
- **脚本（refcite.py）**：消费映射表 → 生成 REF 域 → 加书签 → 重排参考文献

## 修改脚本

1. 修改 `scripts/refcite.py`
2. 用真实 DOCX 文档跑一次 `--replace --verify`，确认输出路径正确、校验 PASS
3. 在 `TEST.md` 新增测试记录

### 注意

- 不引入新依赖。当前仅依赖 `python-docx`
- `clean_refs()` 只删域结构 run，不碰显示文本 run
- 同段多个 `[N]` 处理顺序是从右到左（避免位置偏移）
- 校验函数 `verify_output()` 检测到 `end` 后停止窗口扫描

## 修改文档

- `SKILL.md`：入口层，指向 `docs/` 目录下的详细文档
- `DUAL_TRACK.md`：双轨 / 场景E 的完整说明
- `TEST.md`：所有测试记录

## 提 PR

1. fork 本仓库
2. 创建特性分支
3. 提交改动
4. 开 PR，描述改了什么、测试了哪些文档
