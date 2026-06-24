# 工作流程（Step 0-7）

> 本文档说明 `refcite.py` 的执行流程，代码片段为原理示意，**不承诺与脚本实时一致**。
> 日常使用优先调用 `scripts/refcite.py`，各步骤的具体实现以该脚本为准。

---

## Step 0：备份原文件

强制。使用固定名称 `{原文件名}.bak.docx`，仅首次创建，后续跳过。

```python
backup_path = os.path.join(dir_name, f'{base_name}.bak.docx')
if os.path.exists(backup_path) and os.path.getsize(backup_path) == os.path.getsize(docx_path):
    return backup_path  # 跳过
shutil.copy2(docx_path, backup_path)
assert os.path.getsize(docx_path) == os.path.getsize(backup_path)
```

三个文件平级存在：
- `{原名}.docx` — 原始文件（从未被修改）
- `{原名}.bak.docx` — 备份
- `{原名}_已引用.docx` — 输出

## Step 0a：清理旧 REF 域

移除文档中已有的 `w:fldChar` 和 `w:instrText[REF]` 元素，避免重复堆积。

> 仅清理 XML 域结构，不影响 `w:t` 中的纯文本 `[N]` 标记。模式 D 不受影响。

## Step 1：读取与解析文献

从文末「参考文献」标题之后提取文献条目。支持 `【参考文献】`、`References`、`Bibliography` 等标题。

## Step 2：段落过滤

跳过以下段落：
- 封面信息（学号、姓名、成绩等）
- 摘要、关键词
- 章节标题
- 参考文献列表本身
- 含脚注的段落
- 空段落、短段落（< 15 字）

## ### 模式 D（--replace）：原位替换手打标记

扫描正文中的 `[N]` 标记，直接替换为 REF 域，不做任何匹配。

```bash
python refcite.py 论文.docx --replace
```

流程：
1. 遍历文档正文，用正则 `\[\d+\]` 查找所有手打标记
2. 提取编号，映射到参考文献列表
3. 将每个 `[N]` 文本在 `<w:t>` 中移除，在原位置插入 5-run REF 域
4. 如 `[N]` 后有后续文本，创建新 run 承接

---

Step 3：自动匹配 + 三段混合定位

### 匹配策略（模式③）

```
对每个有效段落，对每条参考文献：
  分数 = 段落前40字含关键词×3 + 段落后半含关键词×2
         + 主题词重叠×2 + 出现次数×1
```

### 插入点定位（三段混合策略）

> 对所有匹配模式均生效。模式①/②/③的区别只在于选段落，插入点定位统一由 skill 处理。

**第一关：学术写作模式匹配**

| 模式 | 正则 | 插入位置 |
|------|------|---------|
| 书名号 | `《.+?》` | 吸附到后面首个 `，` 或 `。` 前 |
| 学者名+观点动词 | `认为\|指出\|提出\|主张\|强调` | 同上 |
| 引用句式 | `正如.+所言\|研究表明\|据报道` | 同上 |

搜索范围：触发位置后 **30 字符内**，超出则回落段末。

**第二关：句级关键词评分**

第一关未命中时，按句号拆分为句子，取关键词重叠度最高的句子末尾插入。

**第三关：段落末尾兜底**

前两关都未命中时，REF 插在段落末尾。

## Step 4：插入 REF 域

每个引用由 5 个连续 XML run 组成：

```
Run 1: fldChar begin
Run 2: instrText → REF {书签名} \r \h \* MERGEFORMAT
Run 3: fldChar separate
Run 4: display text → [{编号}]
Run 5: fldChar end
```

参数：
- 字号：`sz=24`（12pt）
- 书签名：`_Ref{5000 + 项目编号}`
- 颜色：黑色，无超链接样式
- `\r` 开关：显示书签段落的编号
- `\h` 开关：Ctrl+Click 可跳转

## Step 5：项目编号

按正文首次出现顺序分配编号，不是按文献原始顺序。

```
第1次出现 → [1]
第2次出现 → [2]
...
```

## Step 6：加书签 + 重排参考文献

- 预存元素引用 → 从 body 摘出 → 按项目编号重排 → 加 bookmarkStart/End
- 关键陷阱：摘插后 `doc.paragraphs` 引用悬空，必须预存元素再操作

## Step 7：保存

```python
doc.save(f'{base}_已引用.docx')
```
