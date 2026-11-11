---
name: courseware-to-obsidian
description: >
  把课件整理成结构统一的 Obsidian 讲义笔记。支持 PPTX / PPT / PDF / DOCX / DOC /
  Markdown / 纯文本 / 图片：一门课一个文件夹，一份课件一个 `NN-` 前缀的文件夹，
  一个目录页 + 每节一个笔记，含表格、LaTeX、必要截图、双链、课程总目录与关系图配色。
  Use when 用户要把课件、讲义、PPT、PDF、Word、Markdown 或图片导入 Obsidian，
  或提到课件笔记、课程笔记、讲义笔记、知识库、courseware-to-obsidian、"课件转 Markdown"。
  不需要外部 LLM API：由 Claude 自己读课件、写笔记。
---

# 课件 → Obsidian 讲义

一份课件 = 一个 `<NN>-<课件标题>/` 文件夹 + 一个同名目录页 + 每节一个 `.md`。脚本只做机械活（抽取文字、渲染页面、建课程总目录、写关系图配色）；**笔记内容由你直接写**，不调用任何外部模型。

先读两个参考文件，严格遵守：

- `~/.claude/skills/courseware-to-obsidian/references/note-format.md`：文件命名、frontmatter、正文结构和双链格式。
- `~/.claude/skills/courseware-to-obsidian/references/writing-standard.md`：专业性、术语、公式与写作深度要求。

## 需要的信息

只问缺的，一次问完：

1. **课件路径**：`.pptx / .ppt / .pdf / .docx / .doc / .md / .txt / 图片`，可多个。
2. **Obsidian 仓库根目录**：含 `.obsidian/` 的绝对路径。
3. **课程名**：决定 `<vault>/<课程>/`；可从路径或用户说法推断。
4. **课件顺序**：给出处理顺序；没给就按用户列举顺序，再从文件名自然排序。

环境检查：

```bash
python -c "import pptx, pymupdf, docx; print('ok')"
```

失败则提示 `python -m pip install python-pptx pymupdf python-docx`，不要自己乱装。

## 1. 抽取课件文字

```bash
python ~/.claude/skills/courseware-to-obsidian/scripts/courseware_tool.py extract "<课件路径>"
```

命令会打印一个临时 txt 的路径，用 Read 读取。页数很多（>60）时先只读标题做大纲：

```bash
python ~/.claude/skills/courseware-to-obsidian/scripts/courseware_tool.py extract "<课件路径>" --titles
```

写某一节时再按页取原文：

```bash
python ~/.claude/skills/courseware-to-obsidian/scripts/courseware_tool.py extract "<课件路径>" --pages 12-18
```

各格式的“页”：

| 输入 | 页的含义 | 截图 |
|------|----------|------|
| pptx / ppt / pdf | 幻灯片页 / PDF 物理页 | 可渲染 |
| docx / doc | 转成 PDF 后的物理页 | 可渲染；无转换器时退化为按标题切块，不能截图 |
| md / txt | 每个 `#` / `##` 标题块 | 不截图 |
| 图片 | 每张图 1 页 | 直接用 Read 看原图；需要时复制原图到 attachments |

图片课件不要抽取文字后凭文件名猜内容：**必须用 Read 工具查看原图**再写笔记。纯图片/扫描页同理。

## 2. 规划大纲

- 沿用课件**自带的小节编号**（`1.1`、`1.2`、`2.1`…）；没有编号才按主题切。
- 一份完整章节通常切 8~14 节；短课件按自然段落切，不要为凑数硬拆。
- 多页成块的复习/习题单独成节。
- 产出内部大纲：`编号 | 标题 | 页码范围 | 本节重点`。标题直接用于文件名（如 `1.3-热力学概论.md`）。
- 课件夹名取课件主标题；标题看不清就用源文件名。
- **课件夹顺序**：扫描 `<课程>/` 下已有 `NN-` 文件夹，新课件从最大编号 +1 开始，两位数字（`01-`、`02-`…）。不要用 `lec` 前缀；用户指定顺序时按用户顺序编号。

## 3. 逐节写笔记

目标路径：`<vault>/<课程>/<NN>-<课件标题>/<编号>-<标题>.md`

先建目录：

```bash
mkdir -p "<vault>/<课程>/<NN>-<课件标题>/attachments"
```

每节：

1. 只使用该节页码范围里的内容，**不编造**课件没讲的结论。
2. 按 `note-format.md` 写成 `# 标题` + 多个 `### 子主题`，表格/列表/公式优先，不写大段散文。
3. 题目必带原题：例题、习题、自测题的题干、条件、选项、题图完整抄录，再给解答；长题用 `> [!example]` / `> [!question]`。
4. 过程、对比、算例类重难点主动自绘图解（表格/代码块/ASCII），不要只放截图。
5. 例题、整页题目、无法用文字重建的图 → 在对应位置单独一行写 `【课件截图：第N页】`（每节最多 3 张，页码必须在本节范围内）。**不要截**装饰图、二维码/签到码、校徽/logo 模板、与正文无关的纯照片。

选截图前先生成接触表，一次看清多页再决定：

```bash
python ~/.claude/skills/courseware-to-obsidian/scripts/courseware_tool.py contact "<课件路径>" "3,7-12"
```

命令会打印临时 PNG 路径，用 Read 看这一张；再对入选页执行 `render`。

6. 为每个截图占位渲染 PNG（同一页只渲染一次）：

```bash
python ~/.claude/skills/courseware-to-obsidian/scripts/courseware_tool.py render "<课件路径>" 12 "<vault>/<课程>/<NN>-<课件标题>/attachments/<源文件名主干>_p12.png"
```

成功则把占位替换为：

```markdown
![课件原图 · 第12页](attachments/<源文件名主干>_p12.png)
```

图片课件需要保留原图时，用 `render "<图片路径>" 1 "<目标.png>"`（会复制原图）。
md / txt 不生成截图；渲染失败（没有 LibreOffice/Office）就删掉占位，不要留下标记。

7. 文件末尾加上一节/下一节导航（仓库相对双链 + 可读别名）。
8. 每写完 2~3 节就落盘一次，不要憋到最后；写完不要回读文件。

## 4. 写课件目录页

`<vault>/<课程>/<NN>-<课件标题>/<NN>-<课件标题>.md` —— folder-note，名字与文件夹相同。包含 frontmatter、`# 📗 课件标题`、返回课程总目录的链接、各节双链目录。格式见 `references/note-format.md`。

用户要求时才在目录页追加「本章小结 + 自测题」；完整的练习题集另写成节文件（见「可选：自测题 / 练习题」）。默认不写。

## 5. 生成/刷新课程总目录 + 关系图

所有课件写完后执行一次：

```bash
python ~/.claude/skills/courseware-to-obsidian/scripts/courseware_tool.py finalize "<vault>/<课程>"
```

它会写 `<课程> 课程总目录.md`，并把课程配色与图谱显示参数合并进 `<vault>/.obsidian/graph.json`（保留用户已有分组）。改完配色需要重开 Obsidian 才生效。

## 6. 自检

```bash
python ~/.claude/skills/courseware-to-obsidian/scripts/courseware_tool.py check "<vault>/<课程>"
```

必须修到 `problems: 0`：目录页齐全、无残留截图占位、双链和 `#标题` 锚点都能解析、图片都存在。最后报告：课件夹路径、节数、截图数、总目录路径。

## 可选：本章小结

- 默认不写。
- 用户提出后，在课件目录页末尾追加 `## 本章小结`。
- 小结最多 8 条，只用本课件讲过的结论。
- 格式见 `references/note-format.md`。

## 可选：自测题 / 练习题

- 格式见 `references/note-format.md` 的 **Exercise note format**。
- 仍写成课件夹内的节文件，编号接在课件现有节后，如 `6.5-保角映射练习题.md`。
- frontmatter 用 `type: exercise`、`层级/练习`，课程标签与课件完全一致。
- 结构固定：导语 → `## 考点地图` 表 → 字母分组（`## A 组　考点`；每题 `### A1` 完整题面 + `> [!success]- 答案`）→ `## 考前速查`。
- 组按难度递增；每组末 `---`；答案全部折叠。
- 增删题目或组后，同步更新「共 N 题」与 `考点地图`。
- 末尾照常加仓库相对双链导航。

## 边界

- **不调用外部 LLM/API**，也不启动任何 GUI/CLI。
- 双链一律写**仓库相对路径 + 别名**：`[[物理化学/01-热力学第一定律/1.3-热力学概论|1.3 热力学概论]]`。裸 `[[焓]]` 会把不同课程的同名小节串在一起。
- 每个文件的课程标签必须完全一致：`课程/<课程名去掉空格和标点>`。
- 重复转换同一课件：覆盖该课件夹下本 skill 生成的 `数字.数字-*.md` 和同名目录页；保留 `attachments/` 与用户自己放的文件。不要碰其它课件夹。
- 多份课件按顺序处理，最后统一 `finalize` 一次。
- 图片/扫描页没有可抽取文字：必须用 Read 看原图，不要凭封面猜内容。
- 旧 `lec1-xxx` 文件夹保持原样；新导入只使用 `NN-` 前缀，不强制迁移旧笔记。
