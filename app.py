import streamlit as st
import openai
import json, io, base64, re
from PIL import Image
import numpy as np
from sklearn.cluster import KMeans
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from reportlab.pdfgen import canvas

# ── App Config ───────────────────────────────────────────────────────────
st.set_page_config(page_title="Deckly – Pro‑grade decks, zero hassle.", layout="wide")
openai.api_key = st.secrets.get("OPENAI_API_KEY", "")

# ── 1) Bullets → Outline + Speaker Notes ────────────────────────────────
@st.cache_data
def bullets_to_outline(bullets: str, tone: str = "Neutral") -> list[dict]:
    """Return [{title, points[], speaker_notes}] via GPT."""
    prompt = f"""TONE: {tone}
Convert these raw bullet points into a polished investor story.
Return ONLY valid JSON: {{"sections":[{{"title":"…","points":["…"],"speaker_notes":"…"}}]}}
BULLETS:\n{bullets}"""
    try:
        resp = openai.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role":"system","content":"You are a world‑class venture analyst & presentation writer."},
                      {"role":"user","content":prompt}],
            temperature=0.25
        )
        content = resp.choices[0].message.content
        sections = json.loads(re.search(r"\{.*\}", content, re.S).group())['sections']
    except Exception as e:
        st.error(f"Outline error: {e}")
        sections = [{"title":"Overview","points":bullets.split('\n'),"speaker_notes":"TODO"}]
    return sections

# ── 2) Palette Extraction ───────────────────────────────────────────────
@st.cache_data
def extract_palette(img_b: bytes, k: int = 4):
    arr = np.array(Image.open(io.BytesIO(img_b)).convert("RGB")).reshape(-1,3)
    km = KMeans(n_clusters=k, n_init="auto").fit(arr)
    return [f"#{r:02x}{g:02x}{b:02x}" for r,g,b in km.cluster_centers_.astype(int)]

# ── 3) Vision Logo Style ────────────────────────────────────────────────
@st.cache_data
def describe_logo(img_b: bytes):
    try:
        b64 = base64.b64encode(img_b).decode()
        resp = openai.chat.completions.create(
            model="gpt-4o-mini",
            messages=[[{"type":"text","text":"Describe this logo style in one sentence."},
                       {"type":"image_url","image_url":{"url":f"data:image/png;base64,{b64}"}}]],
            temperature=0.5
        )
        return resp.choices[0].message.content.strip()
    except: return ""

# ── 4) PPTX Builder with Accent Bar + Fonts + Notes ─────────────────────
@st.cache_data
def build_deck(sections, palette, logo_b=None):
    prs = Presentation()
    sw, sh = prs.slide_width, prs.slide_height
    hex2rgb = lambda h: tuple(int(h.lstrip('#')[i:i+2],16) for i in (0,2,4))
    bg_rgb = hex2rgb(palette[0]) if palette else (255,255,255)
    accent_rgb = hex2rgb(palette[1]) if len(palette)>1 else (0,0,0)

    for idx,sec in enumerate(sections):
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.background.fill.solid(); slide.background.fill.fore_color.rgb = RGBColor(*bg_rgb)
        # title
        title_shape = slide.shapes.title or slide.shapes.add_textbox(Inches(0.6),Inches(0.3),Inches(9),Inches(1))
        title_shape.text = sec['title']
        title_shape.text_frame.paragraphs[0].font.size = Pt(42)
        title_shape.text_frame.paragraphs[0].font.color.rgb = RGBColor(*accent_rgb)
        # body bullets
        body_tf = None
        try:
            body_tf = slide.placeholders[1].text_frame; body_tf.clear()
        except:
            for sh in slide.shapes:
                if getattr(sh,'has_text_frame',False) and sh!=title_shape:
                    body_tf = sh.text_frame; body_tf.clear(); break
        if body_tf:
            for ptxt in sec['points']:
                p = body_tf.add_paragraph(); p.text = ptxt; p.level=0; p.font.size = Pt(20)
        # speaker notes
        slide.notes_slide.notes_text_frame.text = sec.get('speaker_notes','')
        # accent bar bottom
        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, sh-Inches(0.25), sw, Inches(0.25))
        bar.fill.solid(); bar.fill.fore_color.rgb = RGBColor(*accent_rgb); bar.line.color.rgb = RGBColor(*accent_rgb)
        # logo on title slide only
        if logo_b and idx==0:
            slide.shapes.add_picture(io.BytesIO(logo_b), sw-Inches(1.5), Inches(0.2), width=Inches(1.2))
    buf = io.BytesIO(); prs.save(buf); return buf.getvalue()

# ── 5) Executive Summary via LLM then PDF ───────────────────────────────
@st.cache_data
def render_brief(sections):
    try:
        resp = openai.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role":"system","content":"Write a 1‑page crisp investment brief summarizing the deck."},
                      {"role":"user","content":json.dumps(sections)}],
            temperature=0.4
        )
        summary = resp.choices[0].message.content.strip()
    except Exception as e:
        summary = "Executive Summary\n\n" + json.dumps(sections)[:1000]
    # simple PDF rendering
    buf = io.BytesIO(); c = canvas.Canvas(buf); w,h = c._pagesize; y=h-50
    c.setFont("Helvetica",12)
    for line in summary.split('\n'):
        for chunk in [line[i:i+100] for i in range(0,len(line),100)]:
            c.drawString(40,y,chunk); y-=15
            if y<50: c.showPage(); y=h-50
    c.save(); return buf.getvalue()

# ── Streamlit UI ─────────────────────────────────────────────────────────
st.title("🛠️ Deckly – Pro‑grade decks, zero hassle.")
with st.sidebar:
    raw = st.text_area("Paste bullets (one per line)",height=250)
    tone = st.select_slider("Tone",["Conservative","Neutral","Bold"],value="Neutral")
    logo_file = st.file_uploader("Upload logo",type=["png","jpg","jpeg"])
    k = st.slider("Palette size",2,8,4)
if st.button("Generate Deck"):
    if not raw.strip(): st.warning("Enter some bullets first.")
    else:
        sections = bullets_to_outline(raw,tone)
        if logo_file:
            lb = logo_file.read(); palette = extract_palette(lb,k); st.caption(describe_logo(lb))
        else:
            palette=["#2b2b2b","#007acc","#585858"]
        deck_bytes = build_deck(sections,palette,lb if logo_file else None)
        pdf_bytes  = render_brief(sections)
        st.success("Ready! Download below →")
        st.download_button("⬇️ Deck (PPTX)",deck_bytes,"deckly_deck.pptx")
        st.download_button("⬇️ 1‑Pager (PDF)",pdf_bytes,"deckly_brief.pdf")
