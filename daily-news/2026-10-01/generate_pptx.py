"""
generate_pptx.py — 幼儿园今日新闻 PPT 生成器
2026-10-01 星期四
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_AUTO_SIZE, PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
import os

# ---------- Palette ----------
PALETTE = {
    "bg":       RGBColor(0xFF, 0xF9, 0xE8),  # warm cream
    "green":    RGBColor(0xEA, 0xF8, 0xEF),  # card 1
    "blue":     RGBColor(0xEA, 0xF5, 0xFF),  # card 2
    "peach":    RGBColor(0xFF, 0xF0, 0xE7),  # card 3
    "ink":      RGBColor(0x26, 0x38, 0x4A),
    "red":      RGBColor(0xFF, 0x64, 0x75),
    "yellow":   RGBColor(0xFF, 0xD7, 0x5E),
    "white":    RGBColor(0xFF, 0xFF, 0xFF),
    "muted":    RGBColor(0x65, 0x74, 0x82),
}

FONT_HEAD = "PingFang SC"
FONT_BODY = "PingFang SC"
FONT_HEAD_EA = "PingFang SC"
FONT_BODY_EA = "PingFang SC"

CHAR_BUDGET = {
    "title": 40,
    "body": 80,
    "meta": 50,
    "source": 60,
}

def guard(text, key):
    limit = CHAR_BUDGET[key]
    if len(text) <= limit:
        return text
    return text[:limit - 1] + "…"

def add_text(slide, text, x, y, w, h, *,
             size=14, bold=False, color=None, align=PP_ALIGN.LEFT,
             font=None, font_ea=None, budget_key=None):
    if budget_key:
        text = guard(text, budget_key)
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
    from pptx.oxml.ns import qn
    ea = rPr.find(qn("a:ea"))
    if ea is None:
        ea = rPr.makeelement(qn("a:ea"), {"typeface": font_ea or FONT_BODY_EA})
        rPr.append(ea)
    else:
        ea.set("typeface", font_ea or FONT_BODY_EA)
    return tb

def add_rect(slide, x, y, w, h, fill, line=None, radius=None):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    shp = slide.shapes.add_shape(shape_type,
                                 Inches(x), Inches(y), Inches(w), Inches(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
    shp.shadow.inherit = False
    return shp

def add_image(slide, path, x, y, w, h):
    if os.path.exists(path):
        slide.shapes.add_picture(path, Inches(x), Inches(y), Inches(w), Inches(h))

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
                        '', content, flags=_re.IGNORECASE
                    )
                    data = content.encode("utf-8")
                if item.filename == "_rels/.rels":
                    content = data.decode("utf-8")
                    content = _re.sub(
                        r'<Relationship[^>]*Target="docProps/thumbnail[^"]*"[^>]*/?>',
                        '', content, flags=_re.IGNORECASE
                    )
                    data = content.encode("utf-8")
                zout.writestr(item, data)
        os.replace(tmp, path)
    except Exception as e:
        print("warning: failed to strip thumbnail:", e)
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except Exception:
            pass

def main():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    assets = os.path.join(os.path.dirname(__file__), "assets")

    # ===== Slide 1: Three news cards =====
    s = prs.slides.add_slide(prs.slide_layouts[6])
    # Background
    add_rect(s, 0, 0, 13.333, 7.5, PALETTE["bg"])

    # Title
    add_text(s, "小小新闻", 0.5, 0.3, 4, 0.8,
             size=40, bold=True, color=PALETTE["ink"],
             font=FONT_HEAD, font_ea=FONT_HEAD_EA)

    # Meta (date + weather) - top right
    add_rect(s, 8.5, 0.35, 4.5, 0.55, PALETTE["white"], radius=True)
    add_text(s, "10月1日 星期四 上海 阴有小雨 21—23℃", 8.6, 0.38, 4.3, 0.5,
             size=14, bold=True, color=PALETTE["ink"], align=PP_ALIGN.CENTER,
             font=FONT_BODY, font_ea=FONT_BODY_EA, budget_key="meta")

    # --- Card 1: Pandas ---
    card_y = 1.3
    card_h = 5.8
    card_w = 3.8
    gap = 0.35
    x1 = 0.5

    add_rect(s, x1, card_y, card_w, card_h, PALETTE["green"], radius=True)
    # Number badge
    add_rect(s, x1 + 0.15, card_y + 0.15, 0.4, 0.4, PALETTE["yellow"], radius=True)
    add_text(s, "1", x1 + 0.15, card_y + 0.15, 0.4, 0.4,
             size=16, bold=True, color=PALETTE["ink"], align=PP_ALIGN.CENTER)
    # Image
    add_image(s, os.path.join(assets, "pandas-birthday.jpg"),
              x1 + 0.15, card_y + 0.65, card_w - 0.3, 2.2)
    # Photo note
    add_text(s, "图里：三只大熊猫围着生日蛋糕过国庆节",
             x1 + 0.15, card_y + 2.65, card_w - 0.3, 0.35,
             size=9, bold=True, color=PALETTE["white"], align=PP_ALIGN.CENTER)
    # Title
    add_text(s, "上海野生动物园三只大熊猫一起过生日",
             x1 + 0.2, card_y + 3.1, card_w - 0.4, 0.7,
             size=18, bold=True, color=PALETTE["ink"],
             font=FONT_HEAD, font_ea=FONT_HEAD_EA, budget_key="title")
    # Body
    add_text(s, "雪宝、康康和芊金，三只熊猫都在十月过生日。饲养员给它们做了特别的生日蛋糕和礼物。",
             x1 + 0.2, card_y + 3.8, card_w - 0.4, 1.2,
             size=13, color=PALETTE["ink"], budget_key="body")
    # Ask
    add_text(s, "你想送它们什么礼物？",
             x1 + 0.2, card_y + 5.0, card_w - 0.4, 0.35,
             size=13, bold=True, color=PALETTE["red"])
    # Source
    add_text(s, "中国新闻网 · 9月30日",
             x1 + 0.2, card_y + 5.45, card_w - 0.4, 0.25,
             size=8, color=PALETTE["muted"], budget_key="source")

    # --- Card 2: Starry sky ---
    x2 = x1 + card_w + gap
    add_rect(s, x2, card_y, card_w, card_h, PALETTE["blue"], radius=True)
    add_rect(s, x2 + 0.15, card_y + 0.15, 0.4, 0.4, PALETTE["yellow"], radius=True)
    add_text(s, "2", x2 + 0.15, card_y + 0.15, 0.4, 0.4,
             size=16, bold=True, color=PALETTE["ink"], align=PP_ALIGN.CENTER)
    add_image(s, os.path.join(assets, "starry-sky.jpg"),
              x2 + 0.15, card_y + 0.65, card_w - 0.3, 2.2)
    add_text(s, "图里：七颗亮亮的星星和月亮挤在一起",
             x2 + 0.15, card_y + 2.65, card_w - 0.3, 0.35,
             size=9, bold=True, color=PALETTE["white"], align=PP_ALIGN.CENTER)
    add_text(s, "国庆夜空七姐妹星团和月亮做伴",
             x2 + 0.2, card_y + 3.1, card_w - 0.4, 0.7,
             size=18, bold=True, color=PALETTE["ink"],
             font=FONT_HEAD, font_ea=FONT_HEAD_EA, budget_key="title")
    add_text(s, '天还没亮时，七颗亮亮的星星靠着月亮。它们叫"七姐妹星团"，像七个好朋友手拉手。',
             x2 + 0.2, card_y + 3.8, card_w - 0.4, 1.2,
             size=13, color=PALETTE["ink"], budget_key="body")
    add_text(s, "你找到过天上的星星吗？",
             x2 + 0.2, card_y + 5.0, card_w - 0.4, 0.35,
             size=13, bold=True, color=PALETTE["red"])
    add_text(s, "央视网 · 10月1日",
             x2 + 0.2, card_y + 5.45, card_w - 0.4, 0.25,
             size=8, color=PALETTE["muted"], budget_key="source")

    # --- Card 3: Diving ---
    x3 = x2 + card_w + gap
    add_rect(s, x3, card_y, card_w, card_h, PALETTE["peach"], radius=True)
    add_rect(s, x3 + 0.15, card_y + 0.15, 0.4, 0.4, PALETTE["yellow"], radius=True)
    add_text(s, "3", x3 + 0.15, card_y + 0.15, 0.4, 0.4,
             size=16, bold=True, color=PALETTE["ink"], align=PP_ALIGN.CENTER)
    add_image(s, os.path.join(assets, "diving-gold.jpg"),
              x3 + 0.15, card_y + 0.65, card_w - 0.3, 2.2)
    add_text(s, "图里：中国选手从高处跳下翻进水里",
             x3 + 0.15, card_y + 2.65, card_w - 0.3, 0.35,
             size=9, bold=True, color=PALETTE["white"], align=PP_ALIGN.CENTER)
    add_text(s, "中国跳水队在亚运会拿到第100块金牌",
             x3 + 0.2, card_y + 3.1, card_w - 0.4, 0.7,
             size=18, bold=True, color=PALETTE["ink"],
             font=FONT_HEAD, font_ea=FONT_HEAD_EA, budget_key="title")
    add_text(s, "亚运会跳水比赛，中国队一直拿第一。现在一共拿了100块金牌，太厉害了！",
             x3 + 0.2, card_y + 3.8, card_w - 0.4, 1.2,
             size=13, color=PALETTE["ink"], budget_key="body")
    add_text(s, "你数数看100有多大？",
             x3 + 0.2, card_y + 5.0, card_w - 0.4, 0.35,
             size=13, bold=True, color=PALETTE["red"])
    add_text(s, "新华社 · 9月30日",
             x3 + 0.2, card_y + 5.45, card_w - 0.4, 0.25,
             size=8, color=PALETTE["muted"], budget_key="source")

    # Footer
    add_text(s, "小主播：选一条你最喜欢的新闻讲给大家听。",
             0.5, 7.15, 8, 0.3,
             size=11, bold=True, color=PALETTE["muted"])

    out = os.path.join(os.path.dirname(__file__), "今日新闻-2026-10-01.pptx")
    save_pptx(prs, out)
    print("wrote", out)

if __name__ == "__main__":
    main()
