from pathlib import Path
from shutil import copy2
import os
import re
import zipfile

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt
from PIL import Image

ROOT = Path(__file__).resolve().parents[1]
ASSETS = ROOT / "daily-news" / "2026-09-28" / "assets"
OUT = ROOT / "outputs" / "今日新闻-2026-09-28.pptx"
PUBLISHED = ROOT / "daily-news" / "2026-09-28" / "今日新闻-2026-09-28.pptx"

C = {
    "ink": RGBColor(0x21, 0x35, 0x47),
    "cream": RGBColor(0xFF, 0xFA, 0xF0),
    "sun": RGBColor(0xFF, 0xC8, 0x57),
    "coral": RGBColor(0xFF, 0x6B, 0x6B),
    "mint": RGBColor(0x63, 0xC7, 0xB2),
    "mint_bg": RGBColor(0xE8, 0xF7, 0xEC),
    "sky": RGBColor(0x72, 0xB7, 0xE6),
    "sky_bg": RGBColor(0xE9, 0xF1, 0xFF),
    "peach": RGBColor(0xFF, 0xF0, 0xDC),
    "pink": RGBColor(0xFF, 0xF1, 0xF2),
    "white": RGBColor(0xFF, 0xFF, 0xFF),
    "muted": RGBColor(0x5C, 0x6E, 0x7C),
}
FONT = "PingFang SC"
BUDGET = {"title": 28, "body": 70, "prompt": 45, "footer": 100, "label": 18}


def guard(text, key):
    limit = BUDGET[key]
    return text if len(text) <= limit else text[: limit - 1] + "…"


def set_run_font(run, size, color, bold=False):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    rpr = run._r.get_or_add_rPr()
    ea = rpr.find(qn("a:ea"))
    if ea is None:
        ea = rpr.makeelement(qn("a:ea"), {"typeface": FONT})
        rpr.append(ea)
    else:
        ea.set("typeface", FONT)


def add_text(slide, text, x, y, w, h, size=24, color=None, bold=False,
             align=PP_ALIGN.LEFT, budget="body", margin=0.04):
    text = guard(text, budget)
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = Inches(margin)
    tf.margin_right = Inches(margin)
    tf.margin_top = Inches(margin)
    tf.margin_bottom = Inches(margin)
    p = tf.paragraphs[0]
    p.alignment = align
    p.space_after = Pt(0)
    run = p.add_run()
    run.text = text
    set_run_font(run, size, color or C["ink"], bold)
    return box


def rect(slide, x, y, w, h, fill, radius=True, line=None):
    kind = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    shape = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    if line:
        shape.line.color.rgb = line
    else:
        shape.line.fill.background()
    shape.shadow.inherit = False
    return shape


def bg(slide, color):
    rect(slide, 0, 0, 13.333, 7.5, color, radius=False)
    rect(slide, 12.72, -0.2, 0.9, 7.9, C["coral"], radius=True)


def add_picture_cover(slide, path, x, y, w, h):
    with Image.open(path) as im:
        iw, ih = im.size
    frame_ratio = w / h
    image_ratio = iw / ih
    pic = slide.shapes.add_picture(str(path), Inches(x), Inches(y), width=Inches(w), height=Inches(h))
    if image_ratio > frame_ratio:
        visible = frame_ratio / image_ratio
        crop = (1 - visible) / 2
        pic.crop_left = crop
        pic.crop_right = crop
    elif image_ratio < frame_ratio:
        visible = image_ratio / frame_ratio
        crop = (1 - visible) / 2
        pic.crop_top = crop
        pic.crop_bottom = crop
    return pic


def add_header(slide, eyebrow, title):
    add_text(slide, eyebrow, 0.65, 0.42, 7.4, 0.38, 12, C["coral"], True, budget="label")
    add_text(slide, title, 0.65, 0.82, 11.7, 0.86, 31, C["ink"], True, budget="title")


def add_prompt(slide, text, x=0.7, y=6.18, w=11.85):
    rect(slide, x, y, w, 0.68, C["white"], radius=True)
    rect(slide, x, y, 0.12, 0.68, C["coral"], radius=False)
    add_text(slide, text, x + 0.22, y + 0.10, w - 0.38, 0.42, 18, C["ink"], True, budget="prompt")


def add_footer(slide, text, number):
    add_text(slide, text, 0.72, 7.08, 10.8, 0.2, 8, C["muted"], budget="footer")
    add_text(slide, f"{number}/7", 11.65, 7.03, 0.7, 0.25, 10, C["muted"], True, PP_ALIGN.RIGHT, "label")


def save_pptx(prs, path):
    """Save and strip the built-in blank thumbnail."""
    prs.save(path)
    tmp = str(path) + ".tmp"
    with zipfile.ZipFile(path, "r") as zin, zipfile.ZipFile(tmp, "w") as zout:
        for item in zin.infolist():
            if item.filename.lower().startswith("docprops/thumbnail"):
                continue
            data = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                text = data.decode("utf-8")
                text = re.sub(r'<Override[^>]*PartName="/docProps/thumbnail[^"]*"[^>]*/?>', "", text, flags=re.I)
                data = text.encode("utf-8")
            if item.filename == "_rels/.rels":
                text = data.decode("utf-8")
                text = re.sub(r'<Relationship[^>]*Target="docProps/thumbnail[^"]*"[^>]*/?>', "", text, flags=re.I)
                data = text.encode("utf-8")
            zout.writestr(item, data)
    os.replace(tmp, path)


def build_deck():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # 1 Cover
    s = prs.slides.add_slide(prs.slide_layouts[6])
    rect(s, 0, 0, 13.333, 7.5, C["mint"], radius=False)
    rect(s, 8.7, -1.6, 5.7, 5.7, C["sun"], radius=True)
    rect(s, -0.7, 6.15, 7.8, 1.7, C["white"], radius=True)
    add_text(s, "幼儿园大班 · 今日分享", 0.78, 0.78, 6.5, 0.4, 14, C["ink"], True, budget="label")
    add_text(s, "小小新闻\n播报员", 0.72, 1.45, 9.1, 2.55, 48, C["ink"], True, budget="title")
    add_text(s, "熊猫旅行、科学家的信、太空猜图", 0.78, 4.45, 10.4, 0.62, 24, C["ink"], True, budget="body")
    rect(s, 0.78, 5.28, 5.4, 0.68, C["white"], radius=True)
    add_text(s, "2026年9月28日  星期一  上海", 1.0, 5.43, 5.0, 0.36, 17, C["ink"], True, budget="body")
    add_text(s, "开场：大家早上好，今天我带来了三个有趣的消息。", 0.78, 6.72, 10.7, 0.28, 11, C["muted"], budget="footer")

    # 2 Weather
    s = prs.slides.add_slide(prs.slide_layouts[6]); bg(s, RGBColor(0xFF, 0xF7, 0xDC)); add_header(s, "先看看窗外", "上海今天有小阵雨")
    add_text(s, "23–25℃", 0.72, 2.05, 4.3, 1.08, 52, C["coral"], True, budget="label")
    add_text(s, "东北风 · 记得带雨具 · 小心路滑", 0.75, 3.26, 5.6, 0.84, 23, C["ink"], True, budget="body")
    rect(s, 7.15, 1.75, 4.55, 3.55, C["white"], radius=True)
    sun = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(7.75), Inches(2.05), Inches(1.3), Inches(1.3)); sun.fill.solid(); sun.fill.fore_color.rgb=C["sun"]; sun.line.fill.background()
    cloud = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.35), Inches(3.05), Inches(2.85), Inches(1.05)); cloud.fill.solid(); cloud.fill.fore_color.rgb=C["white"]; cloud.line.color.rgb=RGBColor(0xE0,0xE7,0xEC)
    for dx in (0, .65, 1.3, 1.95):
        drop=s.shapes.add_shape(MSO_SHAPE.TEAR, Inches(8.55+dx), Inches(4.22), Inches(.24), Inches(.45)); drop.fill.solid(); drop.fill.fore_color.rgb=C["sky"]; drop.line.fill.background()
    add_prompt(s, "问问大家：下雨天出门，我们要带什么？")
    add_footer(s, "天气：wttr.in 上海，获取于2026-09-28", 2)

    # 3 Panda
    s = prs.slides.add_slide(prs.slide_layouts[6]); bg(s, C["mint_bg"]); add_header(s, "动物新闻", "两只大熊猫搬新家")
    add_picture_cover(s, ASSETS / "pandas.jpg", 0.72, 1.82, 6.05, 3.78)
    add_text(s, "平平和福双已经到达美国亚特兰大动物园。", 7.18, 1.92, 4.7, 1.12, 26, C["ink"], True, budget="body")
    add_text(s, "它们要先体检、熟悉新家，之后才和游客见面。", 7.18, 3.22, 4.7, 1.22, 23, C["ink"], budget="body")
    add_prompt(s, "一起想：如果你是熊猫饲养员，会准备什么？")
    add_footer(s, "来源：新华社（2026-09-28）｜图片：央视新闻画面", 3)

    # 4 Science letters
    s = prs.slides.add_slide(prs.slide_layouts[6]); bg(s, C["peach"]); add_header(s, "科学新闻", "科学家给孩子写信")
    add_picture_cover(s, ASSETS / "science-letters.jpg", 7.25, 1.76, 4.7, 3.72)
    facts = ["好奇地看看海底有什么", "去发现地球的小秘密", "数学难，也要勇敢想一想"]
    for i, text in enumerate(facts, 1):
        y = 1.88 + (i-1)*1.16
        rect(s, 0.72, y, 5.85, 0.88, C["white"], radius=True)
        rect(s, 0.88, y+0.15, 0.55, 0.55, C["coral"], radius=True)
        add_text(s, str(i), 0.88, y+0.22, 0.55, 0.28, 15, C["white"], True, PP_ALIGN.CENTER, "label")
        add_text(s, text, 1.65, y+0.18, 4.6, 0.48, 21, C["ink"], True, budget="body")
    add_prompt(s, "你最想问科学家什么问题？")
    add_footer(s, "来源：中国科学院转人民日报（2026-09-28）", 4)

    # 5 Satellite
    s = prs.slides.add_slide(prs.slide_layouts[6]); bg(s, C["sky_bg"]); add_header(s, "太空看地球", "卫星拍到什么？")
    add_picture_cover(s, ASSETS / "satellite-puzzler.jpg", 0.72, 1.76, 6.3, 3.9)
    add_text(s, "这是卫星从太空拍下的美国瓦斯克斯岩。", 7.42, 1.95, 4.45, 1.15, 25, C["ink"], True, budget="body")
    add_text(s, "深绿色是比较湿润的植物，灰色弯线是一条公路。", 7.42, 3.35, 4.45, 1.28, 22, C["ink"], budget="body")
    add_prompt(s, "先别说答案：它像树叶、恐龙，还是一幅画？")
    add_footer(s, "来源与图片：NASA Earth Observatory / Landsat 9", 5)

    # 6 Quiz
    s = prs.slides.add_slide(prs.slide_layouts[6]); bg(s, C["pink"]); add_header(s, "记忆挑战", "今天你记住了吗？")
    questions = ["熊猫叫什么名字？", "今天出门带什么？", "卫星从哪里拍地球？"]
    colors = [C["mint"], C["sun"], C["sky"]]
    for i, (question, color) in enumerate(zip(questions, colors)):
        x = 0.72 + i*4.03
        rect(s, x, 2.02, 3.58, 2.65, C["white"], radius=True)
        rect(s, x+0.25, 2.27, 0.65, 0.65, color, radius=True)
        add_text(s, str(i+1), x+0.25, 2.43, 0.65, 0.25, 16, C["ink"], True, PP_ALIGN.CENTER, "label")
        add_text(s, question, x+0.32, 3.24, 2.94, 0.9, 23, C["ink"], True, PP_ALIGN.CENTER, "body")
    add_prompt(s, "请三位同学，每人回答一道题。")
    add_footer(s, "回答后别忘了说：谢谢你的回答。", 6)

    # 7 Close
    s = prs.slides.add_slide(prs.slide_layouts[6])
    rect(s, 0, 0, 13.333, 7.5, C["ink"], radius=False)
    rect(s, 8.8, -1.8, 5.8, 5.8, C["coral"], radius=True)
    add_text(s, "今天的关键词", 0.82, 0.92, 5.0, 0.42, 15, C["sun"], True, budget="label")
    add_text(s, "好奇 · 观察\n勇敢表达", 0.78, 1.58, 10.5, 2.28, 47, C["white"], True, budget="title")
    add_text(s, "谢谢大家听我分享！", 0.82, 4.45, 8.0, 0.75, 28, C["white"], True, budget="body")
    rect(s, 0.82, 5.52, 7.8, 0.72, C["white"], radius=True)
    add_text(s, "结束：鞠躬，说“我的分享结束了，谢谢大家”。", 1.04, 5.68, 7.3, 0.36, 17, C["ink"], True, budget="prompt")
    add_text(s, "新闻与图片仅用于课堂学习。", 0.82, 6.86, 7.5, 0.28, 10, RGBColor(0xD7,0xE2,0xE8), budget="footer")
    add_text(s, "7/7", 11.65, 7.03, 0.7, 0.25, 10, C["white"], True, PP_ALIGN.RIGHT, "label")

    return prs


def main():
    OUT.parent.mkdir(parents=True, exist_ok=True)
    PUBLISHED.parent.mkdir(parents=True, exist_ok=True)
    prs = build_deck()
    save_pptx(prs, OUT)
    copy2(OUT, PUBLISHED)
    print(f"wrote {OUT}")
    print(f"published copy {PUBLISHED}")


if __name__ == "__main__":
    main()
