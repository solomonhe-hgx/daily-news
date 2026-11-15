"""
generate_pptx_2026_11_17.py — 幼儿园大班 今日新闻 2026-11-17 PPTX 生成器
3条新闻，每条一页，无封面，无 天 气页。
包含简 单 的 动 画效果。
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_AUTO_SIZE, PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pptx.oxml import parse_xml

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
ASSETS_DIR = os.path.join(SCRIPT_DIR, "daily-news", "2026-11-17", "assets")

# ---------- 统 一 文 案（与 网 页 逐 句 一 致） ----------
NEWS = [
    {
        "chip": "1 / 3",
        "topic": '体育纪录',
        "title": '中国举重选手李闫世锦赛破世界纪录夺三金',
        "body1": '11月8日，在宁波举行的世界举重锦标赛上，中国选手李闫在女子86公斤以上级比赛里，抓举举起147公斤，打破了世界纪录。',
        "body2": '她最后拿到了抓举、挺举和总成绩三块金牌，中国队在本次世锦赛上一共赢了17枚金牌。',
        "purpose": '李闫姐姐能举起比自己重好几倍的杠铃，她的力量来自每一天的刻苦训练。你有没有坚持做过一件事情？',
        "image": os.path.join(ASSETS_DIR, 'weightlifting-liyan.png'),
    },
    {
        "chip": "2 / 3",
        "topic": '环境保护·绿色能源',
        "title": '沙漠里建起了"绿色发电站"',
        "body1": '在宁夏的毛乌素沙漠里，人们建起了一座很大的太阳能发电站，一万多块太阳能板铺在沙地上，像一片蓝色的海洋。',
        "body2": '太阳能板不仅能发电，还能挡住太阳和风，让板下面的沙子不再到处飞，小草和灌木慢慢长了出来，沙漠变成了绿地。',
        "purpose": '这个发电站既能生产干净的电，又能把沙漠变成绿地，一举两得。你知道太阳能是怎么变成电的吗？',
        "image": os.path.join(ASSETS_DIR, 'solar-desert.png'),
    },
    {
        "chip": "3 / 3",
        "topic": '体育·电子竞技',
        "title": '中国战队BLG夺得英雄联盟全球总决赛冠军',
        "body1": '11月14日，在美国纽约举行的英雄联盟全球总决赛上，中国战队BLG以3比1战胜了韩国战队GEN，夺得了冠军。',
        "body2": '这是中国战队时隔五年再次拿到这个比赛的冠军，队员卓定被评为最有价值选手。',
        "purpose": 'BLG战队在比赛中一路打败了四支韩国强队，最后拿到了冠军奖杯。你和小伙伴一起玩游戏或做游戏的时候，是怎么配合的？',
        "image": os.path.join(ASSETS_DIR, 'esports-blg.jpg'),
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


def add_animation(slide, shape, delay_ms=0, duration_ms=500):
    """Add a simple fade-in animation to a shape."""
    timing = slide._element.find(qn('p:timing'))
    if timing is None:
        timing = parse_xml(
            '<p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
            '<p:tnLst><p:par><p:cT id="1" dur="indefinite" restart="never" nodeType="tmRoot">'
            '<p:childTnLst><p:seq concurrent="1" nextAc="seek">'
            '<p:cT id="2" dur="indefinite" nodeType="mainSeq">'
            '<p:childTnLst></p:childTnLst></p:cT></p:seq></p:childTnLst></p:cT></p:par></p:tnLst></p:timing>'
        )
        slide._element.append(timing)

    seq = timing.find(qn('p:tnLst')).find(qn('p:par')).find(qn('p:cT')).find(qn('p:childTnLst')).find(qn('p:seq'))
    child_tn_lst = seq.find(qn('p:cT')).find(qn('p:childTnLst'))

    shape_id = shape.shape_id
    anim_xml = f'''<p:par xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
      <p:cT id="{3 + len(child_tn_lst)}" fill="hold">
        <p:stCondLst><p:cond delay="{delay_ms}"/></p:stCondLst>
        <p:childTnLst>
          <p:set>
            <p:cBhvr><p:cT id="{4 + len(child_tn_lst)}" dur="1" fill="hold">
              <p:stCondLst><p:cond delay="0"/></p:stCondLst>
            </p:cT><p:tgtEl><p:spTgt spid="{shape_id}"/></p:tgtEl>
            <p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>
          </p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>
          <p:animEffect transition="in" filter="fade">
            <p:cBhvr><p:cT id="{5 + len(child_tn_lst)}" dur="{duration_ms}"/>
            <p:tgtEl><p:spTgt spid="{shape_id}"/></p:tgtEl></p:cBhvr>
          </p:animEffect>
        </p:childTnLst>
      </p:cT>
    </p:par>'''

    child_tn_lst.append(parse_xml(anim_xml))


def build_news_slide(prs, news, slide_num):
    s = prs.slides.add_slide(prs.slide_layouts[6])

    # Background
    bg = add_rect(s, 0, 0, 13.333, 7.5, PALETTE["bg"])

    # Chip
    chip = add_rounded_rect(s, 0.6, 0.45, 1.1, 0.45, PALETTE["chip_bg"])
    chip_text = add_text(s, news["chip"], 0.6, 0.45, 1.1, 0.45,
             size=14, bold=True, color=PALETTE["ink"],
             align=PP_ALIGN.CENTER)

    # Topic
    topic = add_text(s, news["topic"], 1.9, 0.45, 3.0, 0.45,
             size=14, color=PALETTE["muted"])

    # Image
    img_path = news["image"]
    if os.path.exists(img_path):
        img_card = add_rounded_rect(s, 0.5, 1.1, 6.7, 5.7, PALETTE["card_bg"])
        img_shape = s.shapes.add_picture(img_path, Inches(0.6), Inches(1.2),
                                          Inches(6.5), Inches(5.5))
    else:
        img_card = add_rounded_rect(s, 0.6, 1.2, 6.5, 5.5, PALETTE["chip_bg"])
        img_shape = add_text(s, "（图片未找到）", 0.6, 3.0, 6.5, 1.0,
                 size=20, color=PALETTE["muted"], align=PP_ALIGN.CENTER)

    # Title
    title = add_text(s, news["title"], 7.5, 1.2, 5.3, 1.45,
             size=36, bold=True, color=PALETTE["ink"])

    # Body 1
    body1 = add_text(s, news["body1"], 7.5, 2.7, 5.3, 1.25,
             size=22, color=PALETTE["ink"])

    # Body 2
    body2 = add_text(s, news["body2"], 7.5, 3.95, 5.3, 1.3,
             size=22, color=PALETTE["ink"])

    # Purpose
    purpose_card = add_rounded_rect(s, 7.5, 5.35, 5.3, 1.35, RGBColor(0xFF, 0xF6, 0xE5))
    purpose = add_text(s, news["purpose"], 7.6, 5.4, 5.1, 1.25,
             size=18, bold=True, color=PALETTE["accent"])

    # Footer
    footer_left = add_text(s, "小小新闻播报 · 2026年11月17日 星期二", 0.6, 7.0, 6.0, 0.35,
             size=10, color=PALETTE["muted"])
    footer_right = add_text(s, f"第{slide_num}页，共3页", 7.5, 7.0, 5.3, 0.35,
             size=10, color=PALETTE["muted"], align=PP_ALIGN.RIGHT)

    # Add animations with staggered delays
    add_animation(s, bg, delay_ms=0, duration_ms=300)
    add_animation(s, chip, delay_ms=200, duration_ms=400)
    add_animation(s, topic, delay_ms=300, duration_ms=400)
    add_animation(s, img_card, delay_ms=400, duration_ms=500)
    add_animation(s, title, delay_ms=600, duration_ms=500)
    add_animation(s, body1, delay_ms=800, duration_ms=500)
    add_animation(s, body2, delay_ms=1000, duration_ms=500)
    add_animation(s, purpose_card, delay_ms=1200, duration_ms=500)
    add_animation(s, footer_left, delay_ms=1400, duration_ms=400)
    add_animation(s, footer_right, delay_ms=1400, duration_ms=400)


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

    for i, news in enumerate(NEWS, 1):
        build_news_slide(prs, news, i)

    out_dir = os.path.join(SCRIPT_DIR, "daily-news", "2026-11-17")
    os.makedirs(out_dir, exist_ok=True)
    out = os.path.join(out_dir, "今 日 新 闻-2026-11-17.pptx")
    save_pptx(prs, out)
    print("wrote", out)


if __name__ == "__main__":
    main()
