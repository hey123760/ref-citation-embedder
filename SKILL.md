---
name: ref-citation-embedder
description: 自动将参考文献嵌入论文正文（DOCX），生成 Word REF 域代码（上标 [n] 编号 + Ctrl+点击跳转至文末对应条目）。适用于任何学科的本科/硕士论文、期刊论文、文献综述、学位论文的引用格式化。
---

# ref-citation-embedder — AI 操作指令

## 你的角色

你是 AI。你的工作是分析论文正文，决定每处引用放在哪，产出映射表（`.txt`），然后调脚本执行。不写 python-docx 代码。

## 三步流程（场景 E，推荐）

```
你：分析正文 → 写 plan.txt
脚本：inject_marks.py --strip         清空旧标记
脚本：inject_marks.py --map-file      写入 [N]
脚本：refcite.py --replace --verify   [N] → REF 域
```

## 第一步：分析正文，写映射表

用 `python -c "import docx; doc=docx.Document('论文.docx'); ..."` 读取段落内容。

决定每段哪个句子引用哪篇文献。产出 `plan.txt`，格式：

```
P7:3:。[5]     段落7 第3句 句号前 插 [5]
P18:1:。[1]    段落18 第1句 句号前 插 [1]
P23:2:。[7]    段落23 第2句 句号前 插 [7]
```

- `P{段落索引}` — 0-based，doc.paragraphs 的下标
- `{第几句}` — 1-based，按句号拆分
- `{标点}` — 通常用 `。`，支持 `。！？；，`
- `[N]` — 与文末参考文献列表中的编号一致

句级不够精确时可用段级（`--para`）：
```
P7:[5]     段落7 段尾 插 [5]
```

### 规则

| 原则 | 说明 |
|------|------|
| 每段别堆太多 | 一段里引用不要超过 2 处，最好 1 处 |
| 别总插第一句 | 选最相关的那句话，不是第一句也不是最后一句 |
| 不熟的文献 | 问用户要不要搜索确认，不搜就凭训练知识判断 |
| 跳过段落 | 封面、摘要、关键词、章标题、参考文献列表本身不插引用 |
| 作者名 ≠ 需引用 | 段落出现作者名不代表内容出自该文献 |
| 传记信息不引 | 出生日期、职位等公开信息不属可引用内容 |

## 第二步：调脚本

```bash
# 清空旧标记（换方案时必做，首次可跳过）
python scripts/inject_marks.py 论文.docx --strip

# 注入新标记
python scripts/inject_marks.py 论文.docx --map-file plan.txt

# 替换为 REF 域
python scripts/refcite.py 论文.docx --replace --verify
```

### 其他场景

| 文档状态 | 你要做的 | 命令 |
|---------|---------|------|
| 正文已有手打 `[N]` 且正确 | 直接跑 refcite | `refcite.py --replace --keep-numbers --verify` |
| 正文有 `[N]` 但需换方案 | `--strip` 清空 → 重新写 plan.txt → inject → replace | 同上三步 |
| 无参考文献列表也无 `[N]` | 请用户先补充参考文献列表 | — |

### 常用命令速查

```bash
# 注入（句级）
inject_marks.py 论文.docx --map-file plan.txt

# 注入（段级）
inject_marks.py 论文.docx --map "P7:[5]" --para

# 清空旧 [N]
inject_marks.py 论文.docx --strip

# 替换 [N] → REF 域
refcite.py 论文.docx --replace --verify

# 保持原编号（不重排）
refcite.py 论文.docx --replace --keep-numbers --verify
```

> ⚠️ `--strip` 会检测文档是否已有 REF 域，有则拒绝执行，防止误删已格式化文档。需回到原始 DOCX(未跑过 --replace)再操作。

## 常见问题

### "错误!不能识别的开关参数" / "错误!未找到引用源"

Word REF 域导出 PDF 前没有锁定。正确操作：`Ctrl+A` → `Ctrl+F11` 锁定域再导出。

### "[1] 变成了 { REF _Ref5001 \h }"

Word 进入了域代码显示模式。Word 中按 `Alt+F9` 切换回来。

### 文档有坏域（"错误!不能识别的开关参数"）

不要自己预处理。直接写 `[N]` 后跑 `--replace`，`clean_refs()` 会自动清理旧域。

## 目录结构

```
scripts/inject_marks.py   消费映射表 → 写 [N] 手打标记
scripts/refcite.py        [N] → REF 域代码 + 书签 + 重排参考文献
docs/DUAL_TRACK.md        双轨用法详细说明
```

## 验证清单

最后确认：
- [ ] 所有文献都在正文中出现
- [ ] 编号按首次出现顺序排列
- [ ] Ctrl+Click 可跳转
- [ ] 上标为黑色，无蓝色下划线
- [ ] 原文件已备份为 `.bak.docx`
