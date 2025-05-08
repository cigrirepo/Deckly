import streamlit as st
import openai
import json
import io
import base64
import re
from PIL import Image
import numpy as np
from sklearn.cluster import KMeans
from pptx import Presentation
from pptx.util import Inches
from pptx.dml.color import RGBColor
from reportlab.pdfgen import canvas

# -- Configuration
st.set_page_config(page_title="Deckly – Pro-grade decks, zero hassle.", layout="wide")
openai.api_key = st.secrets.get("OPENAI_API_KEY", "")

# -- 1) Bullets → Structured Outline via LLM (v1 client)
@st.cache_data
def bullets_to_outline(bullets: str, tone: str = "Neutral") -> list[dict]:
    prompt = f"""TONE: {tone}
Convert the following bullet points into a structured JSON outline.
Provide ONLY the JSON array of sections with keys 'title' and 'points'.
BULLETS:
{bullets}"""
    try:
        response = openai.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "You are a world-class pitch-deck strategist."},
                {"role": "user", "content": prompt},
            ],
            temperature=0.3
        )
        content = response.choices[0].message.content
        match = re.search(r"(\[.*\])", content, re.S)
        if match:
            outline = json.loads(match.group(1))
        else:
            raise ValueError("No JSON array found in model output.")
    except Exception as e:
        st.error(f"Outline generation error: {e}")
        outline = [{"title": "Overview", "points": bullets.splitlines()}]
    return outline

# -- 2) Palette Extraction via KMeans
@st.cache_data
def extract_palette(image_bytes: bytes, k: int = 4) -> list[str]:
    img = Image.open(io.BytesIO(image_bytes)).convert("RGB")
    arr = np.array(img).reshape(-1, 3)
    km = KMeans(n_clusters=k, n_init="auto").fit(arr)
    colors = km.cluster_centers_.astype(int)
    return [f"#{r:02x}{g:02x}{b:02x}" for r, g, b in colors]

# -- 3) Logo Style Description (Vision) via v1 client
@st.cache_data
def describe_logo(image_bytes: bytes) -> str:
    try:
        b64 = base64.b64encode(image_bytes).decode()
        # Multimodal payload for Vision call
        payload = [
            {"type": "text", "text": "Describe this logo's style and mood in one sentence."},
            {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{b64}"}}
        ]
        resp = openai.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": payload}],
            temperature=0.5
        )
        return resp.choices[0].message.content.strip()
    except Exception:
        return ""

# -- 4) Slide Generation with python-pptx
@st.cache_data
def build_deck(outline: list[dict], palette: list[str], logo_bytes: bytes | None) -> bytes:
    prs = Presentation()
    def hex_to_rgb(h: str) -> tuple[int,int,int]:
        h = h.lstrip('#')
        return tuple(int(h[i:i+2], 16) for i in (0,2,4))
    bg = hex_to_rgb(palette[0]) if palette else (255,255,255)
    for sec in outline:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor(*bg)
        title_txt = sec.get("title", "")
        if slide.shapes.title:
            slide.shapes.title.text = title_txt
        else:
            tx = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(1))
            tx.text_frame.text = title_txt
        body_tf = None
        try:
            body_tf = slide.placeholders[1].text_frame
            body_tf.clear()
        except Exception:
            for shape in slide.shapes:
                if getattr(shape, 'has_text_frame', False) and shape is not slide.shapes.title:
                    body_tf = shape.text_frame
                    body_tf.clear()
                    break
        if body_tf:
            for pt in sec.get("points", []):
                p = body_tf.add_paragraph()
                p.text = pt
                p.level = 0
        if logo_bytes:
            try:
                stream = io.BytesIO(logo_bytes)
                prs_width, prs_height = prs.slide_width, prs.slide_height
                pic = slide.shapes.add_picture(
                    stream,
                    prs_width - Inches(1.2), prs_height - Inches(1.2),
                    width=Inches(1), height=Inches(1)
                )
            except Exception:
                pass
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()

# -- 5) PDF Brief via ReportLab
@st.cache_data
def render_brief(outline: list[dict], palette: list[str]) -> bytes:
    buf = io.BytesIO()
    c = canvas.Canvas(buf)
    width, height = c._pagesize
    y = height - 40
    c.setFont("Helvetica-Bold", 18)
    c.drawString(40, y, "Executive Summary")
    c.setFont("Helvetica", 12)
    y -= 30
    for sec in outline:
        c.drawString(40, y, sec.get("title", ""))
        y -= 20
        for pt in sec.get("points", []):
            c.drawString(60, y, f"- {pt}")
            y -= 15
            if y < 50:
                c.showPage()
                y = height - 40
    c.save()
    return buf.getvalue()

# -- Streamlit UI
st.title("🛠️ Deckly – Pro-grade decks, zero hassle.")
with st.sidebar:
    st.header("Inputs")
    raw_bullets = st.text_area("Paste bullets (one per line)", height=250)
    tone = st.select_slider("Tone", ["Conservative","Neutral","Bold"], value="Neutral")
    logo_file = st.file_uploader("Upload logo (PNG/JPG)", type=["png","jpg","jpeg"])
    k = st.slider("Palette colors", 2, 8, 4)
if st.button("Generate Deck"):
    if not raw_bullets.strip():
        st.warning("Enter some bullet points first.")
    else:
        outline = bullets_to_outline(raw_bullets, tone)
        if logo_file:
            img_b = logo_file.read()
            palette = extract_palette(img_b, k)
            style = describe_logo(img_b)
            st.write("**Logo style:**", style)
        else:
            palette = ["#333333","#666666","#999999"]
        deck = build_deck(outline, palette, logo_file.read() if logo_file else None)
        brief = render_brief(outline, palette)
        st.success("Done! Download below:")
        st.download_button("Download PPTX", deck, "deckly_deck.pptx")
        st.download_button("Download PDF", brief, "deckly_summary.pdf")
