# ref-citation-embedder

> 自动将参考文献嵌入 Word 论文正文（DOCX），生成可点击跳转的规范引用标记（REF 域代码）。
> 适用于本科/硕士论文、期刊论文、文献综述、学位论文的引用格式化。

---

## 使用方式

| 方式 | 起点 | 命令 |
|------|------|------|
| **场景D** | 用户已有完整手打 `[N]` | `--replace --keep-numbers` |
| **双轨** | 无标记，AI 产映射表 → 脚本消费 | 详见下方 |

双轨是核心工作流：AI 分析正文 → 产出映射表 → 脚本格式化。映射表有两种载体：

| 载体 | 命令 | 说明 |
|------|------|------|
| `--smap` 参数 | `--smap "5:1:。"` | 映射表通过命令行传入 |
| 正文 `[N]` 标记（场景E） | `--replace` | AI 将映射表写入正文，脚本替换 `[N]` 标记 |

两种载体本质相同，场景E 的优势是在用户有部分标记的情况下处理效果更好

## 安装

```bash
pip install python-docx
```

Python 3.8+，依赖见 `requirements.txt`。

## 快速开始

```bash
# 场景D：正文已有完整手打标记
python scripts/refcite.py 论文.docx --replace --keep-numbers --verify

# 双轨（--smap）：AI 产出句级映射表
python scripts/refcite.py 论文.docx --smap "5:1:。,12:2:。" --verify

# 双轨·场景E（[N] 载体）：AI 写全套标记后替换
python scripts/refcite.py 论文.docx --replace --verify
```

## 文档

| 文件 | 说明 |
|------|------|
| `SKILL.md` | 入口，包含场景说明、硬性规则、常见误区 |
| `docs/DUAL_TRACK.md` | 双轨用法、场景E 变体、AI 对接指南 |
| `docs/WORKFLOW.md` | 脚本执行流程（Step 0-7） |
| `TEST.md` | 测试历史、Bug 清单 |
| `CONTRIBUTING.md` | 贡献指南 |

## 与 Zotero 等文献管理工具的区别

Zotero 需要用户在写作过程中插入字段、维护文献库，适合从零开始的论文。ref-citation-embedder 面向**已经写完的论文**：正文已有 `[N]` 标记，文末已有参考文献列表，脚本自动完成格式化，不需要用户额外学习或改变写作习惯。

AI 双轨模式下，用户只需提供文档，AI 分析正文后自动写入标记，脚本执行替换——全流程自动化，不需要用户在工具之间切换。

## 安全边界

| 脚本行为 | 说明 |
|---------|------|
| 不改原文语义 | 只插入 REF 域、重排参考文献、加书签 |
| 强制备份 | 运行前创建 `.bak.docx` |
| 另存输出 | 不修改原文件，输出为 `*_已引用.docx` |
| 附录自动排除 | 参考文献后的「附录」「致谢」等内容不纳入处理 |

## 许可证

MIT
