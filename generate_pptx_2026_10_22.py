"""
generate_pptx.py — 幼儿园大班今日新闻 2026-10-22 PPTX 生成器
3条新闻，每条一页，无封面，无天气页。
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_AUTO_SIZE, PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# ---------- Palette ----------
PALETTE = {
    "bg":       RGBColor(0xFF, 0xFA, 0xF0),
    "ink":      RGBColor(0x25, 0x45, 0x37),
    "accent":   RGBColor(0xBC, 0x4F, 0x32),
    "chip_bg":  RGBColor(0xE8, 0xF0, 0xE5),
    "card_bg":  RGBColor(0xFF, 0xFF, 0xFF),
    "muted":    RGBColor(0x52, 0x66, 0x5B),
    "white":    RGBColor(0xFF, 0xFF, 0xFF),
}

FONT_HEAD = "Microsoft YaHei"
FONT_BODY = "Microsoft YaHei"

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR = os.path.join(SCRIPT_DIR, "daily-news", "2026-10-22", "assets")

NEWS = [
    {
        "chip": "1 / 3",
        "topic": "新铁路上的大桥",
        "title": "新大桥上的火车试跑",
        "body1": "重庆到万州要修一条新铁路，上面有一座很大很大的桥。",
        "body2": "10月20日，火车第一次在桥上试跑，看看一切安不安全。",
        "purpose": "以后坐火车去万州，会更快更方便。",
        "image": os.path.join(ASSETS_DIR, "yuwan-train.png"),
    },
    {
        "chip": "2 / 3",
        "topic": "科学小发现",
        "title": "废塑料变成小海绵",
        "body1": "科学家把不要的塑料，变成了一种像海绵一样的新材料。",
        "body2": "这种小海绵可以从空气里吸出水来，加热以后水又会跑出来。",
        "purpose": "不要的东西，动脑筋也能变成有用的宝贝。",
        "image": os.path.join(ASSETS_DIR, "molecular-sponge.jpg"),
    },
    {
        "chip": "3 / 3",
        "topic": "太阳能发电",
        "title": "220米高的太阳能大塔",
        "body1": "在新疆，有一座220米高的大塔，周围摆满了镜子。",
        "body2": "镜子把阳光照到塔顶，把盐烧得很热很热，就能发电了。",
        "purpose": "太阳光也能变成电，让灯亮起来、让电视转起来。",
        "image": os.path.join(ASSETS_DIR, "solar-tower.jpg"),
    },
]


def add_text(slide, text, x, y, w, h, *,
             size=14, bold=False, color=None, align=PP_ALIGN.LEFT,
             font=None, font_ea=None):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.name = font or FONT_BODY
    r.font.size = Pt(size)
    r.font.bold = bold
    if color is not None:
        r.font.color.rgb = color
    rPr = r._r.get_or_add_rPr()
    ea = rPr.find(qn("a:ea"))
    if ea is None:
        ea = rPr.makeelement(qn("a:ea"), {"typeface": font_ea or FONT_BODY})
        rPr.append(ea)
    else:
        ea.set("typeface", font_ea or FONT_BODY)
    return tb


def add_rect(slide, x, y, w, h, fill, line=None):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                 Inches(x), Inches(y), Inches(w), Inches(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
    shp.shadow.inherit = False
    return shp


def add_rounded_rect(slide, x, y, w, h, fill):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                 Inches(x), Inches(y), Inches(w), Inches(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def build_news_slide(prs, news):
    s = prs.slides.add_slide(prs.slide_layouts[6])

    add_rect(s, 0, 0, 13.333, 7.5, PALETTE["bg"])

    add_rounded_rect(s, 0.6, 0.45, 1.1, 0.45, PALETTE["chip_bg"])
    add_text(s, news["chip"], 0.6, 0.45, 1.1, 0.45,
             size=14, bold=True, color=PALETTE["ink"],
             align=PP_ALIGN.CENTER)

    add_text(s, news["topic"], 1.9, 0.45, 3.0, 0.45,
             size=14, color=PALETTE["muted"])

    img_path = news["image"]
    if os.path.exists(img_path):
        add_rounded_rect(s, 0.5, 1.1, 6.7, 5.7, PALETTE["card_bg"])
        img_shape = s.shapes.add_picture(img_path, Inches(0.6), Inches(1.2),
                                          Inches(6.5), Inches(5.5))
    else:
        add_rounded_rect(s, 0.6, 1.2, 6.5, 5.5, PALETTE["chip_bg"])
        add_text(s, "（图片未找到）", 0.6, 3.0, 6.5, 1.0,
                 size=20, color=PALETTE["muted"], align=PP_ALIGN.CENTER)

    add_text(s, news["title"], 7.5, 1.2, 5.3, 1.2,
             size=36, bold=True, color=PALETTE["ink"])

    add_text(s, news["body1"], 7.5, 2.6, 5.3, 0.8,
             size=22, color=PALETTE["ink"])

    add_text(s, news["body2"], 7.5, 3.5, 5.3, 0.8,
             size=22, color=PALETTE["ink"])

    add_rounded_rect(s, 7.5, 4.6, 5.3, 0.75, RGBColor(0xFF, 0xF6, 0xE5))
    add_text(s, news["purpose"], 7.6, 4.65, 5.1, 0.65,
             size=18, bold=True, color=PALETTE["accent"])

    add_text(s, "小小新闻播报 · 2026年10月22日 星期四", 0.6, 7.0, 6.0, 0.35,
             size=10, color=PALETTE["muted"])
    add_text(s, "每条新闻独立一页，共3页", 7.5, 7.0, 5.3, 0.35,
             size=10, color=PALETTE["muted"], align=PP_ALIGN.RIGHT)


def save_pptx(prs, path):
    prs.save(path)
    import zipfile as _zf
    import re as _re
    tmp = path + ".tmp"
    try:
        with _zf.ZipFile(path, "r") as zin, _zf.ZipFile(tmp, "w") as zout:
            for item in zin.infolist():
                if item.filename.lower().startswith("docprops/thumbnail"):
                    continue
                data = zin.read(item.filename)
                if item.filename == "[Content_Types].xml":
                    content = data.decode("utf-8")
                    content = _re.sub(
                        r'<Override[^>]*PartName="/docProps/thumbnail[^"]*"[^>]*/?>',
                        '', content, flags=_re.IGNORECASE)
                    data = content.encode("utf-8")
                if item.filename == "_rels/.rels":
                    content = data.decode("utf-8")
                    content = _re.sub(
                        r'<Relationship[^>]*Target="docProps/thumbnail[^"]*"[^>]*/?>',
                        '', content, flags=_re.IGNORECASE)
                    data = content.encode("utf-8")
                zout.writestr(item, data)
        os.replace(tmp, path)
    except Exception as e:
        print("warning: failed to strip template thumbnail:", e)
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except Exception:
            pass


def main():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    for news in NEWS:
        build_news_slide(prs, news)

    out_dir = os.path.join(SCRIPT_DIR, "daily-news", "2026-10-22")
    os.makedirs(out_dir, exist_ok=True)
    out = os.path.join(out_dir, "今日新闻-2026-10-22.pptx")
    save_pptx(prs, out)
    print("wrote", out)


if __name__ == "__main__":
    main()
