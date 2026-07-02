---
name: cn-defense-deck
version: "1.1.0"
description: >
  使用 python-pptx 生成中文学术答辩PPT（开题/中期/毕业答辩）。
  提供 ThesisDefensePPT 类，支持 10 种页面类型、富文本红色关键词、校徽 banner。
  触发：用户提到"生成答辩PPT"、"开题答辩PPT"、"毕业答辩PPT"、"thesis defense PPT"、"学术PPT"、"答辩幻灯片"、"答辩deck"等。
license: MIT
platforms: [windows]
compatibility: "python-pptx, Pillow; PowerPoint COM for PDF export (Windows only), LibreOffice as cross-platform alternative"
metadata:
  category: document-creation
---

# cn-defense-deck — 学术答辩PPT生成

基于 python-pptx 的学术答辩PPT生成，自包含类，无需模板。

## 文件结构

```
cn-defense-deck/
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
| `cover_slide()` | 封面 | title, eng_title, subtitle, reporter, supervisor, date_str |
| `toc_slide()` | 目录 | items (list) |
| `section_divider_slide()` | 章节分隔 | chapter_num, chapter_title |
| `content_slide()` | 内容页（富文本） | title, subtitle, rich_lines, body_font_size=20, footnotes=None, body_height=4.8 |
| `bullet_slide()` | 编号要点页 | title, subtitle, bullets |
| `image_text_slide()` | 图文混排 | title, subtitle, rich_lines, img_path, img_caption |
| `two_image_slide()` | 双图对比 | title, subtitle, img1_path, img2_path, caption1, caption2 |
| `full_image_slide()` | 全幅图片 | title, subtitle, img_path, caption |
| `center_image_slide()` | 居中图片 | title, subtitle, img_path, caption |
| `end_slide()` | 结束页 | main_text, sub_text |

## `content_slide()` 详解

```python
def content_slide(self, title, subtitle, rich_lines,
                  body_font_size=20, footnotes=None, body_height=4.8)
```

| 参数 | 说明 |
|------|------|
| `title` | 红色层级标题，自动拆分主标题（32pt）和次级标题（24pt） |
| `subtitle` | 浅灰蓝副标题栏的内容。**传空字符串 `""` 则不渲染副标题栏** |
| `rich_lines` | 正文行列表，支持 `【关键词】` 语法（红色加粗）和 `(text, color, bold)` 元组 |
| `body_font_size` | 正文字号，默认 20pt。参考文献页推荐 12-14pt |
| `footnotes` | 底部引用标注文本（14pt 黑色）。每页底部小字标注关键参考文献 |
| `body_height` | 正文区域高度（英寸），默认 4.8。无脚注页可加大到 5.2 |

### 使用示例

```python
# 带脚注的内容页
ppt.content_slide(
    title="2. 研究内容与技术路线 — 特色与创新",
    subtitle="三点应用创新：旋翼干扰自适应抑制、射频直采+ICZT联合应用",
    rich_lines=[
        "（1）【技术链路】形成\"机理建模→参数优选→自适应抑制\"完整链路。",
        "（2）【射频直采+ICZT联合应用】以射频直采简化前端，以ICZT替代IFFT。",
    ],
    footnotes="参考文献：综述[2-3] · ICZT[8-10][33] · 微多普勒[11]"
)

# 无副标题栏的参考文献页
ppt.content_slide(
    title="参考文献",
    subtitle="",
    body_font_size=12,
    body_height=5.2,
    rich_lines=[
        "[1] Author. Title[J]. Journal, Year.",
        ...
    ]
)
```

### 富文本语法
```python
rich_lines=[
    "普通文字，【关键词】自动红色加粗。",
    ("自定义颜色文字", RGBColor(0xDD, 0xDD, 0xDD), False),
]
```

## 设计规范

### 色彩系统
| 用途 | 色值 | 常量名 |
|------|------|--------|
| 主色 Navy | #003399 | `NAVY` |
| 标题红 | #C00000 | `RED` |
| 副标题蓝 | #005DA2 | `SECONDARY_BLUE` |
| 浅灰蓝背景 | #E9EDF4 | `LIGHT_BLUE_GRAY` |
| 正文色 | #333333 | `DARK_GRAY` |
| 白色 | #FFFFFF | `WHITE` |

### 字体规范
| 元素 | 字体 | 字号 | 颜色 |
|------|------|------|------|
| 封面中文标题 | 微软雅黑 | 34pt | 白色 |
| 封面英文标题 | Arial | 22pt | #003399 |
| 内容页标题（主） | 微软雅黑 | 32pt 加粗 | #C00000 |
| 内容页标题（次级） | 微软雅黑 | 24pt 加粗 | #C00000 |
| 副标题栏 | 微软雅黑 | 20pt | #333333 |
| 正文 | 微软雅黑 | 20pt | #333333 |
| 脚注 | 微软雅黑 | 14pt | #333333 |

### 页面尺寸
16:9 宽屏，13.33" × 7.5"

## 布局规则

1. **标题宽度 ≤ 9.5"**（x=1.5），避免与右上角校名重叠
2. **副标题栏** `SHAPE_TO_FIT_TEXT` 自适应高度，subtitle="" 则不渲染
3. **正文起始 y=1.8"**，与副标题栏 ≥0.2" 间距
4. **Bullet间距** 动态计算 `min(0.85", 可用高度/条目数)`，底部 ≥0.5" margin
5. **封面英文标题** 在 Navy 色块下方，NAVY 色文字（不在色块内）
6. **脚注** 在正文区域下方，14pt 黑色
7. **图片** 用 `add_picture_constrained()` 保持宽高比

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
- 封面英文清晰（NAVY色在海军条下方）
- 无 `【【` / `】】` 残留
- 文字无溢出
- 脚注引用编号与末页参考文献编号一致

## 已知限制

- PDF导出需 PowerPoint COM (Windows) 或 LibreOffice（跨平台）
- 中文需系统安装微软雅黑
- Pillow 用于图片宽高比计算
