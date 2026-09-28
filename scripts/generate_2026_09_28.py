from pathlib import Path
from shutil import copy2
import os, re, zipfile
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

ROOT = Path(__file__).resolve().parents[1]
PHOTO = ROOT / "daily-news/2026-09-28/assets/pandas.jpg"
OUT = ROOT / "outputs/今日新闻-2026-09-28.pptx"
WEB = ROOT / "daily-news/2026-09-28/今日新闻-2026-09-28.pptx"
FONT = "PingFang SC"
C = {"ink":RGBColor(0x26,0x38,0x4A),"red":RGBColor(0xFF,0x64,0x75),"yellow":RGBColor(0xFF,0xD7,0x5E),"blue":RGBColor(0x78,0xCC,0xEF),"greenbg":RGBColor(0xEA,0xF8,0xEF),"weather":RGBColor(0xFF,0xF5,0xCC),"white":RGBColor(0xFF,0xFF,0xFF),"muted":RGBColor(0x60,0x70,0x80)}
BUDGET={"title":18,"body":30,"tip":22,"small":70}

def guard(text,key):
    if len(text)>BUDGET[key]: raise ValueError(f"text too long for {key}: {text}")
    return text

def font(run,size,color,bold=False):
    run.font.name=FONT; run.font.size=Pt(size); run.font.bold=bold; run.font.color.rgb=color
    rpr=run._r.get_or_add_rPr(); ea=rpr.find(qn("a:ea"))
    if ea is None: ea=rpr.makeelement(qn("a:ea"),{"typeface":FONT}); rpr.append(ea)
    else: ea.set("typeface",FONT)

def text(slide,value,x,y,w,h,size,color=None,bold=False,align=PP_ALIGN.LEFT,budget="body"):
    value=guard(value,budget); box=slide.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h)); tf=box.text_frame
    tf.word_wrap=True; tf.auto_size=MSO_AUTO_SIZE.NONE; tf.margin_left=Inches(.04); tf.margin_right=Inches(.04); tf.margin_top=Inches(.03); tf.margin_bottom=Inches(.03)
    p=tf.paragraphs[0]; p.alignment=align; p.space_after=Pt(0); r=p.add_run(); r.text=value; font(r,size,color or C["ink"],bold); return box

def shape(slide,kind,x,y,w,h,fill,line=None):
    s=slide.shapes.add_shape(kind,Inches(x),Inches(y),Inches(w),Inches(h)); s.fill.solid(); s.fill.fore_color.rgb=fill
    if line: s.line.color.rgb=line
    else: s.line.fill.background()
    s.shadow.inherit=False; return s

def rect(slide,x,y,w,h,fill,round=True): return shape(slide,MSO_SHAPE.ROUNDED_RECTANGLE if round else MSO_SHAPE.RECTANGLE,x,y,w,h,fill)

def photo_cover(slide,path,x,y,w,h):
    with Image.open(path) as im: iw,ih=im.size
    pic=slide.shapes.add_picture(str(path),Inches(x),Inches(y),width=Inches(w),height=Inches(h)); fr=w/h; ir=iw/ih
    if ir>fr: crop=(1-fr/ir)/2; pic.crop_left=crop; pic.crop_right=crop
    elif ir<fr: crop=(1-ir/fr)/2; pic.crop_top=crop; pic.crop_bottom=crop
    return pic

def save_pptx(prs,path):
    prs.save(path); tmp=str(path)+".tmp"
    with zipfile.ZipFile(path,"r") as zin, zipfile.ZipFile(tmp,"w") as zout:
        for item in zin.infolist():
            if item.filename.lower().startswith("docprops/thumbnail"): continue
            data=zin.read(item.filename)
            if item.filename=="[Content_Types].xml": data=re.sub(r'<Override[^>]*PartName="/docProps/thumbnail[^"]*"[^>]*/?>','',data.decode(),flags=re.I).encode()
            if item.filename=="_rels/.rels": data=re.sub(r'<Relationship[^>]*Target="docProps/thumbnail[^"]*"[^>]*/?>','',data.decode(),flags=re.I).encode()
            zout.writestr(item,data)
    os.replace(tmp,path)

def build():
    prs=Presentation(); prs.slide_width=Inches(13.333); prs.slide_height=Inches(7.5)
    # Page 1: weather
    s=prs.slides.add_slide(prs.slide_layouts[6]); rect(s,0,0,13.333,7.5,C["weather"],False)
    text(s,"小小新闻 · 第1页",.72,.55,5,.4,14,C["red"],True,budget="small")
    rect(s,.72,1.1,4.85,.65,C["white"]); text(s,"9月28日  星期一  上海",.95,1.25,4.4,.3,17,C["ink"],True,budget="small")
    text(s,"今天有小雨",.7,2.05,6.5,1.2,48,C["ink"],True,budget="title")
    text(s,"23—25℃",.72,3.45,4.2,.8,34,C["red"],True,budget="body")
    rect(s,.72,4.55,5.45,.82,C["white"]); text(s,"带小伞，小心地上滑。",.97,4.76,5,.36,21,C["ink"],True,budget="tip")
    shape(s,MSO_SHAPE.OVAL,9.9,.75,1.45,1.45,C["yellow"])
    cloud=rect(s,7.2,2.75,4.5,1.35,C["white"])
    for x,y,d in [(7.75,2.18,1.45),(9.0,1.88,1.8)]: shape(s,MSO_SHAPE.OVAL,x,y,d,d,C["white"])
    for x in (7.75,8.7,9.65,10.6): shape(s,MSO_SHAPE.TEAR,x,4.35,.28,.58,C["blue"])
    text(s,"大家早上好！今天上海有小雨。",.76,6.82,8.7,.3,12,C["muted"],True,budget="small")
    text(s,"1/2",11.7,6.82,.7,.3,11,C["muted"],True,PP_ALIGN.RIGHT,"small")
    # Page 2: real-time panda news
    s=prs.slides.add_slide(prs.slide_layouts[6]); rect(s,0,0,13.333,7.5,C["greenbg"],False)
    photo_cover(s,PHOTO,.62,.6,6.35,5.45); rect(s,5.34,.38,1.9,.68,C["yellow"]); text(s,"到新家啦！",5.5,.56,1.58,.3,17,C["ink"],True,PP_ALIGN.CENTER,"tip")
    text(s,"今天的新鲜事",7.45,.65,4.7,.4,14,C["red"],True,budget="small")
    text(s,"熊猫坐飞机",7.38,1.16,5.05,.88,34,C["ink"],True,budget="title")
    text(s,"平平和福双\n到了新家。",7.42,2.28,4.8,1.35,27,C["ink"],True,budget="body")
    text(s,"它们先休息，\n过些天再见大家。",7.42,3.92,4.8,1.25,23,C["ink"],False,budget="body")
    rect(s,7.42,5.43,4.45,.78,C["white"]); text(s,"你想送熊猫什么礼物？",7.64,5.63,4.05,.32,18,C["ink"],True,budget="tip")
    text(s,"新闻：新华社 2026年9月28日　图片：央视新闻",.68,6.35,8.5,.25,9,C["muted"],False,budget="small")
    text(s,"我的分享说完了，谢谢大家！",.68,6.82,8.7,.3,12,C["muted"],True,budget="small")
    text(s,"2/2",11.7,6.82,.7,.3,11,C["muted"],True,PP_ALIGN.RIGHT,"small")
    return prs

def main():
    OUT.parent.mkdir(parents=True,exist_ok=True); WEB.parent.mkdir(parents=True,exist_ok=True)
    save_pptx(build(),OUT); copy2(OUT,WEB); print(OUT); print(WEB)

if __name__=="__main__": main()
