---
name: thesis-defense-ppt
version: "1.0.0"
description: >
  使用 python-pptx 生成中文学术答辩PPT（开题/中期/毕业答辩）。
  提供 ThesisDefensePPT 类，支持 10 种页面类型、富文本红色关键词、校徽 banner。
  触发：用户提到"生成答辩PPT"、"开题答辩PPT"、"毕业答辩PPT"、"thesis defense PPT"、"学术PPT"、"答辩幻灯片"等。
license: MIT
platforms: [windows]
compatibility: "python-pptx, Pillow; PowerPoint COM for PDF export (Windows only), LibreOffice as cross-platform alternative"
metadata:
  category: document-creation
---

# ThesisDefensePPT — 学术答辩PPT生成

基于 python-pptx 的学术答辩PPT生成，自包含类，无需模板。

## 文件结构

```
thesis-defense-ppt/
├── SKILL.md                      # 本文件 — 使用指南
├── scripts/
│   ├── thesis_defense_ppt.py     # 完整类代码，直接复制使用
│   └── qa_check.py               # 程序化QA脚本（无视觉模式）
└── references/                   # (预留) 设计参考
```

## 安装依赖

```bash
pip install python-pptx Pillow
```

## 快速使用

```python
import sys, os
sys.path.insert(0, "path/to/scripts")
from thesis_defense_ppt import ThesisDefensePPT

ppt = ThesisDefensePPT("答辩PPT.pptx", logo_path="校徽.png", name_path="校名.png")
ppt.cover_slide(title="基于XXX的YYY关键技术研究", eng_title="Research on Key...",
                subtitle="硕士学位论文开题答辩", reporter="张三", date_str="2026年7月")
ppt.toc_slide(items=["绪论","系统设计","硬件实现","实验验证","总结展望"])
ppt.content_slide("研究背景", "行业现状与问题",
                  ["当前领域面临【核心问题A】和【核心问题B】。"])
ppt.end_slide("谢谢各位老师！", "敬请批评指正")
ppt.save()
```

## API 总览

| 方法 | 用途 | 关键参数 |
|------|------|----------|
| `cover_slide()` | 封面 | title, eng_title, subtitle, reporter, date_str |
| `toc_slide()` | 目录 | items (list) |
| `section_divider_slide()` | 章节分隔 | chapter_num, chapter_title |
| `content_slide()` | 内容页（富文本） | title, subtitle, rich_lines |
| `bullet_slide()` | 编号要点页 | title, subtitle, bullets |
| `image_text_slide()` | 图文混排 | title, subtitle, rich_lines, img_path, img_caption |
| `two_image_slide()` | 双图对比 | title, subtitle, img1_path, img2_path, caption1, caption2 |
| `full_image_slide()` | 全幅图片 | title, subtitle, img_path, caption |
| `center_image_slide()` | 居中图片 | title, subtitle, img_path, caption |
| `end_slide()` | 结束页 | main_text, sub_text |

`rich_lines` 中的 `【关键词】` 自动渲染为红色加粗。

## 设计规范

| 元素 | 字体 | 字号 | 颜色 |
|------|------|------|------|
| 封面中文标题 | 微软雅黑 | 36pt | 白色 |
| 封面英文标题 | Arial | 22pt | 白色 |
| 内容页标题 | 微软雅黑 | 32pt 加粗 | #C00000 |
| 副标题栏 | 微软雅黑 | 20pt 加粗 | #333333 |
| 正文 | 微软雅黑 | 20pt | #333333 |

配色：Navy #003399 / 标题红 #C00000 / 正文 #333333 ／ 副标题栏背景 #E9EDF4

## 布局规则

1. **标题宽度 ≤ 9.5"**（x=1.5），避免与右上角校名重叠
2. **正文起始 y=1.8"**，与副标题栏 ≥0.2" 间距
3. **Bullet间距** 动态计算 `min(0.85", 可用高度/条目数)`，底部 ≥0.5" margin
4. **封面英文标题** 白色文字在 Navy 条内（y=3.55），避免 Navy-on-Navy
5. **图片** 用 `_add_image_constrained()` 保持宽高比，不要固定尺寸

## 自定义

```python
# 换校色
NAVY = RGBColor(0x1E, 0x27, 0x61)

# 换尺寸
SLIDE_W = 10.0  # 4:3
```

## QA

### 有多模态：导出图片目视检查

```bash
python your_script.py
python -c "import comtypes.client; ..."  # PPT → PDF
pdftoppm -jpeg -r 150 output.pdf slide
```

### 无多模态：程序化检查

```bash
python qa_check.py 答辩PPT.pptx
```

自动检查：总页数、双括号残留、标题宽度、内容非空。

检查清单：
- 校徽/校名正确显示
- 标题不与校名重叠
- 副标题栏与正文不重叠
- Bullet间距均匀
- 图片有图注
- 封面英文清晰（白色在Navy条内）
- 无 `【【` / `】】` 残留
- 文字无溢出

## 已知限制

- PDF导出需 PowerPoint COM (Windows) 或 LibreOffice（跨平台）
- 中文需系统安装微软雅黑
- Pillow 用于图片宽高比计算
