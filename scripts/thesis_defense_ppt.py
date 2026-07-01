"""
ThesisDefensePPT — 中文学术答辩PPT生成器
基于 python-pptx，支持10种页面类型、富文本红色关键词、校徽banner。
"""
from pptx import Presentation
from pptx.util import Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
import re, os

# ─── 色彩系统（可按需修改） ───────────────────────────────────
NAVY          = RGBColor(0x00, 0x33, 0x99)
RED           = RGBColor(0xC0, 0x00, 0x00)
SECONDARY_BLUE= RGBColor(0x00, 0x5D, 0xA2)
LIGHT_BLUE_GRAY=RGBColor(0xE9, 0xED, 0xF4)
DARK_GRAY     = RGBColor(0x33, 0x33, 0x33)
WHITE         = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_GRAY    = RGBColor(0xAA, 0xBB, 0xCC)

SLIDE_W = 13.33  # 16:9 宽屏（英寸）
SLIDE_H = 7.5

def emu(inches):
    return int(inches * 914400)

def _remove_txBody(shape):
    txBody = shape._element.find(qn('p:txBody'))
    if txBody is not None:
        shape._element.remove(txBody)

class ThesisDefensePPT:
    """
    中文学术答辩PPT生成器。

    参数:
        output_path: 输出 .pptx 文件路径
        logo_path:   校徽 PNG 路径（左上角，可选）
        name_path:   校名 PNG 路径（右上角，可选）
    """

    def __init__(self, output_path, logo_path=None, name_path=None):
        self.prs = Presentation()
        self.prs.slide_width = emu(SLIDE_W)
        self.prs.slide_height = emu(SLIDE_H)
        self.output_path = output_path
        self.logo_path = logo_path
        self.name_path = name_path
        self.blank_layout = self.prs.slide_layouts[6]

    def save(self):
        self.prs.save(self.output_path)

    # ─── 工具方法 ─────────────────────────────────────────────

    def _add_textbox(self, slide, txt, l, t, w, h, sz=16, color=DARK_GRAY,
                     bold=False, align=PP_ALIGN.LEFT, font_name="微软雅黑",
                     line_space=1.2):
        box = slide.shapes.add_textbox(emu(l), emu(t), emu(w), emu(h))
        tf = box.text_frame
        tf.word_wrap = True
        for idx, line in enumerate(txt.split('\n')):
            p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
            p.text = line
            p.font.size = Pt(sz)
            p.font.color.rgb = color
            p.font.bold = bold
            p.font.name = font_name
            p.alignment = align
            p.space_after = Pt(sz * (line_space - 1))
        return box

    def _add_rect(self, slide, l, t, w, h, fill=None):
        sh = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, emu(l), emu(t), emu(w), emu(h))
        if fill:
            sh.fill.solid()
            sh.fill.fore_color.rgb = fill
        else:
            sh.fill.background()
        sh.line.fill.background()
        _remove_txBody(sh)
        return sh

    def _add_triangle(self, slide, l, t, w, h, fill):
        sh = slide.shapes.add_shape(
            MSO_SHAPE.ISOSCELES_TRIANGLE, emu(l), emu(t), emu(w), emu(h))
        sh.fill.solid()
        sh.fill.fore_color.rgb = fill
        sh.line.fill.background()
        _remove_txBody(sh)
        return sh

    def _add_rich_text(self, slide, lines, l, t, w, h, sz=20, color=DARK_GRAY,
                       font_name="微软雅黑", line_space=1.35):
        """【关键词】 → 红色加粗"""
        box = slide.shapes.add_textbox(emu(l), emu(t), emu(w), emu(h))
        tf = box.text_frame
        tf.word_wrap = True
        pattern = re.compile(r'【(.+?)】')
        for idx, line in enumerate(lines):
            p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
            p.space_after = Pt(sz * (line_space - 1))
            last_end = 0
            for m in pattern.finditer(line):
                if m.start() > last_end:
                    run = p.add_run()
                    run.text = line[last_end:m.start()]
                    run.font.size = Pt(sz)
                    run.font.color.rgb = color
                    run.font.name = font_name
                run = p.add_run()
                run.text = m.group(1)
                run.font.size = Pt(sz)
                run.font.color.rgb = RED
                run.font.bold = True
                run.font.name = font_name
                last_end = m.end()
            if last_end < len(line):
                run = p.add_run()
                run.text = line[last_end:]
                run.font.size = Pt(sz)
                run.font.color.rgb = color
                run.font.name = font_name
        return box

    def _add_image_constrained(self, slide, img_path, l, t, max_w, max_h,
                               caption=None, caption_sz=14):
        """按比例缩放图片，可选图注"""
        from PIL import Image
        with Image.open(img_path) as im:
            iw, ih = im.size
        scale = min(max_w / iw, max_h / ih)
        w = int(iw * scale)
        h = int(ih * scale)
        x = l + int((max_w - w) / 2)
        y = t + int((max_h - h) / 2)
        pic = slide.shapes.add_picture(img_path, emu(x), emu(y), emu(w), emu(h))
        bottom = t + max_h
        if caption:
            cap_h = int(0.4 * 914400)
            self._add_textbox(slide, caption, l, bottom + 4, max_w, cap_h,
                              sz=caption_sz, color=DARK_GRAY, align=PP_ALIGN.CENTER)
            bottom += cap_h
        return bottom, h

    def _add_banner(self, slide):
        if self.logo_path and os.path.isfile(self.logo_path):
            slide.shapes.add_picture(self.logo_path, emu(0.3), emu(0.15), height=emu(0.65))
        if self.name_path and os.path.isfile(self.name_path):
            slide.shapes.add_picture(self.name_path, emu(11.1), emu(0.15), height=emu(0.55))

    def _add_title_bar(self, slide, title, y=0.85):
        self._add_textbox(slide, title, 1.5, y, 9.5, 0.65,
                          sz=32, color=RED, bold=True, font_name="微软雅黑")
        bar = self._add_rect(slide, 1.5, y + 0.75, 10.3, 0.45, fill=LIGHT_BLUE_GRAY)
        bar.text_frame.word_wrap = True
        return y + 0.75

    # ─── 页面类型 ─────────────────────────────────────────────

    def cover_slide(self, title, eng_title="", subtitle="",
                    reporter="", date_str="", school=""):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_rect(slide, 0, 1.85, SLIDE_W, 3.36, fill=NAVY)
        self._add_textbox(slide, title, 0.5, 2.1, 12.3, 1.2,
                          sz=36, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
        if eng_title:
            self._add_textbox(slide, eng_title, 0.5, 3.55, 12.3, 0.6,
                              sz=22, color=WHITE, bold=True,
                              font_name="Arial", align=PP_ALIGN.CENTER)
        if subtitle:
            self._add_textbox(slide, subtitle, 0.5, 5.4, 12.3, 0.5,
                              sz=20, color=DARK_GRAY, align=PP_ALIGN.CENTER)
        info_lines = []
        if reporter:
            info_lines.append(f"汇报人：{reporter}")
        if school:
            info_lines.append(f"学  校：{school}")
        if date_str:
            info_lines.append(f"日  期：{date_str}")
        if info_lines:
            self._add_textbox(slide, '\n'.join(info_lines), 0.5, 6.0, 12.3, 1.0,
                              sz=16, color=DARK_GRAY, align=PP_ALIGN.CENTER, line_space=1.5)
        return slide

    def toc_slide(self, items, title="目录 Contents"):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_banner(slide)
        self._add_triangle(slide, 0.4, 1.5, 1.8, 2.5, NAVY)
        self._add_triangle(slide, 1.6, 2.8, 1.2, 1.8, SECONDARY_BLUE)
        self._add_textbox(slide, title, 3.5, 1.5, 8.0, 0.8, sz=36, color=NAVY, bold=True)
        y = 2.6
        for i, item in enumerate(items):
            self._add_textbox(slide, f"{i+1:02d}", 3.8, y, 1.0, 0.6,
                              sz=28, color=NAVY, bold=True, font_name="Arial")
            self._add_textbox(slide, item, 4.9, y + 0.03, 7.0, 0.5, sz=22, color=DARK_GRAY)
            y += 0.85
        return slide

    def section_divider_slide(self, chapter_num, chapter_title):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_banner(slide)
        self._add_triangle(slide, 0.4, 1.5, 1.8, 2.5, NAVY)
        self._add_triangle(slide, 1.6, 2.8, 1.2, 1.8, SECONDARY_BLUE)
        num_str = f"{chapter_num:02d}" if isinstance(chapter_num, int) else str(chapter_num)
        self._add_textbox(slide, num_str, 3.5, 2.0, 2.0, 1.2,
                          sz=72, color=NAVY, bold=True, font_name="Arial")
        self._add_textbox(slide, chapter_title, 3.5, 3.5, 8.0, 0.8,
                          sz=30, color=DARK_GRAY, bold=True)
        return slide

    def content_slide(self, title, subtitle, rich_lines):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_banner(slide)
        bar_y = self._add_title_bar(slide, title)
        self._add_textbox(slide, subtitle, 1.7, bar_y + 0.05, 9.9, 0.35,
                          sz=20, color=DARK_GRAY, bold=True)
        self._add_rich_text(slide, rich_lines, 1.5, 1.8, 10.3, 5.2)
        return slide

    def bullet_slide(self, title, subtitle, bullets):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_banner(slide)
        self._add_title_bar(slide, title)
        self._add_textbox(slide, subtitle, 1.7, 1.55, 9.9, 0.3,
                          sz=18, color=DARK_GRAY, bold=True)
        n = len(bullets)
        start_y = 1.85
        available = SLIDE_H - start_y - 0.5
        spacing = min(0.85, available / n) if n > 0 else 0.85
        y = start_y
        for i, bullet in enumerate(bullets):
            self._add_textbox(slide, f"{i+1:02d}", 1.8, y, 0.7, 0.5,
                              sz=24, color=NAVY, bold=True, font_name="Arial")
            self._add_textbox(slide, bullet, 2.6, y, 9.2, 0.5, sz=20, color=DARK_GRAY)
            y += spacing
        return slide

    def image_text_slide(self, title, subtitle, rich_lines,
                         img_path, img_caption="", image_right=True):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_banner(slide)
        bar_y = self._add_title_bar(slide, title)
        self._add_textbox(slide, subtitle, 1.7, bar_y + 0.05, 9.9, 0.35,
                          sz=18, color=DARK_GRAY, bold=True)
        content_y = 1.85
        if image_right:
            self._add_rich_text(slide, rich_lines, 1.5, content_y, 6.0, 5.0)
            self._add_image_constrained(slide, img_path, 7.8, content_y, 5.0, 4.5, caption=img_caption)
        else:
            self._add_image_constrained(slide, img_path, 1.5, content_y, 5.0, 4.5, caption=img_caption)
            self._add_rich_text(slide, rich_lines, 6.8, content_y, 6.0, 5.0)
        return slide

    def two_image_slide(self, title, subtitle, img1_path, img2_path,
                        caption1="", caption2=""):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_banner(slide)
        self._add_title_bar(slide, title)
        self._add_textbox(slide, subtitle, 1.7, 1.55, 9.9, 0.3, sz=18, color=DARK_GRAY, bold=True)
        self._add_image_constrained(slide, img1_path, 0.5, 1.85, 5.5, 4.2, caption=caption1)
        self._add_image_constrained(slide, img2_path, 6.5, 1.85, 5.5, 4.2, caption=caption2)
        return slide

    def full_image_slide(self, title, subtitle, img_path, caption=""):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_banner(slide)
        self._add_title_bar(slide, title)
        self._add_textbox(slide, subtitle, 1.7, 1.55, 9.9, 0.3, sz=18, color=DARK_GRAY, bold=True)
        self._add_image_constrained(slide, img_path, 0.5, 1.85, 12.3, 4.8, caption=caption)
        return slide

    def center_image_slide(self, title, subtitle, img_path, caption=""):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_banner(slide)
        self._add_title_bar(slide, title)
        self._add_textbox(slide, subtitle, 1.7, 1.55, 9.9, 0.3, sz=18, color=DARK_GRAY, bold=True)
        self._add_image_constrained(slide, img_path, 2.0, 1.85, 9.3, 4.8, caption=caption)
        return slide

    def end_slide(self, main_text="谢谢各位老师！", sub_text="敬请批评指正"):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_rect(slide, 0, 2.6, SLIDE_W, 2.8, fill=NAVY)
        self._add_textbox(slide, main_text, 0.5, 3.0, 12.3, 1.0,
                          sz=40, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
        if sub_text:
            self._add_textbox(slide, sub_text, 0.5, 4.0, 12.3, 0.6,
                              sz=18, color=LIGHT_GRAY, align=PP_ALIGN.CENTER)
        return slide
