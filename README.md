# courseware-to-obsidian

把课件整理成 Obsidian 讲义笔记的 Claude Code skill。

支持 PPTX / PPT / PDF / DOCX / DOC / Markdown / 纯文本 / 图片；统一文件结构：

```
<vault>/<课程>/
├─ <课程> 课程总目录.md
└─ 01-<课件标题>/
   ├─ 01-<课件标题>.md
   ├─ 1.1-<节标题>.md
   └─ attachments/
```

## 安装

```bash
git clone https://github.com/peagothedreamer/courseware-to-obsidian.git
ln -s "$PWD/courseware-to-obsidian/skills/courseware-to-obsidian" \
      "$HOME/.claude/skills/courseware-to-obsidian"
```

Windows 未启用开发者模式时，用目录联接：

```cmd
git clone https://github.com/peagothedreamer/courseware-to-obsidian.git
mklink /J "%USERPROFILE%\.claude\skills\courseware-to-obsidian" "%CD%\courseware-to-obsidian\skills\courseware-to-obsidian"
```

## 依赖

```bash
python -m pip install python-pptx pymupdf python-docx
```

`.ppt` / `.doc` 和截图渲染需要 LibreOffice，或在 Windows 上安装 PowerPoint / Word。

## 自检

```bash
cd courseware-to-obsidian
python skills/courseware-to-obsidian/scripts/selftest.py
```

## 说明

- 笔记由 Claude 直接读课件后写，不调用外部 LLM API。
- `references/writing-standard.md` 是统一的专业写作规范，不按专业维护词典。
