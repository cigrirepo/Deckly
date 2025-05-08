import streamlit as st
import openai
import json
import io
import base64
from PIL import Image
import numpy as np
from sklearn.cluster import KMeans
from pptx import Presentation
from pptx.util import Inches
from reportlab.pdfgen import canvas

# -- Configuration
st.set_page_config(page_title="Deckly – Pro-grade decks, zero hassle.", layout="wide")
openai.api_key = st.secrets["OPENAI_API_KEY"]

# -- LLM Prompt Chaining: Bullets to Structured Outline
@st.cache_data
def bullets_to_outline(bullets: str, tone: str = "Neutral") -> list:
    prompt = (
        f"TONE: {tone}\n"
        "Convert the following bullet points into a structured JSON outline.\n"
        "Output a JSON list of sections with 'title' and 'points'.\n"
        f"BULLETS:\n{bullets}"
    )
    response = openai.ChatCompletion.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": "You are a world-class pitch-deck strategist."},
            {"role": "user", "content": prompt},
        ],
        temperature=0.3,
    )
    content = response.choices[0].message.content
    try:
        outline = json.loads(content)
    except json.JSONDecodeError:
        st.error("Failed to parse outline: ensure the model response is valid JSON.")
        outline = []
    return outline

# -- OpenAI Vision + KMeans Palette Extraction
@st.cache_data
def extract_palette(image_bytes: bytes, k: int = 4) -> list[str]:
    img = Image.open(io.BytesIO(image_bytes)).convert("RGB")
    arr = np.array(img).reshape(-1, 3)
    km = KMeans(n_clusters=k, n_init="auto").fit(arr)
    colors = km.cluster_centers_.astype(int)
    return [f"#{r:02x}{g:02x}{b:02x}" for r, g, b in colors]

@st.cache_data
def describe_logo(image_bytes: bytes) -> str:
    b64 = base64.b64encode(image_bytes).decode()
    multimodal = [
        {"type": "text", "text": "Describe the style and mood of this logo in one sentence."},
        {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{b64}"}},
    ]
    response = openai.ChatCompletion.create(
        model="gpt-4o-mini",
        messages=[{"role": "user", "content": multimodal}],
        temperature=0.5,
    )
    return response.choices[0].message.content.strip()

# -- Slide Generation with python-pptx
@st.cache_data
def build_deck(outline: list[dict], palette: list[str], logo_bytes: bytes | None) -> bytes:
    prs = Presentation()
    for section in outline:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = section.get("title", "")
        body = slide.shapes.placeholders[1].text_frame
        for point in section.get("points", []):
            body.add_paragraph().text = point
        # TODO: apply palette & add logo watermark per slide
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()

# -- Brief PDF Generation with ReportLab
@st.cache_data
def render_brief(outline: list[dict], palette: list[str]) -> bytes:
    buf = io.BytesIO()
    c = canvas.Canvas(buf)
    width, height = c._pagesize
    y = height - 50
    c.setFont("Helvetica-Bold", 16)
    c.drawString(50, y, "Executive Summary")
    c.setFont("Helvetica", 12)
    y -= 30
    for section in outline:
        c.drawString(50, y, section.get("title", ""))
        y -= 20
        for point in section.get("points", []):
            c.drawString(70, y, f"- {point}")
            y -= 15
            if y < 100:
                c.showPage()
                y = height - 50
    c.save()
    return buf.getvalue()

# -- Streamlit UI
st.title("🛠️ Deckly – Pro-grade decks, zero hassle.")

with st.sidebar:
    st.header("Input")
    raw_bullets = st.text_area("Paste your bullet points here", height=250)
    tone = st.select_slider("Tone", ["Conservative", "Neutral", "Bold"], value="Neutral")
    logo_file = st.file_uploader("Upload logo (PNG, JPG)", type=["png", "jpg", "jpeg"])
    k = st.slider("Palette colors", 2, 8, 4)

if st.button("Generate Deck"):
    if not raw_bullets:
        st.error("Please enter bullet points to proceed.")
    else:
        with st.spinner("Generating outline…"):
            outline = bullets_to_outline(raw_bullets, tone)
        if logo_file:
            image_bytes = logo_file.read()
            with st.spinner("Extracting palette…"):
                palette = extract_palette(image_bytes, k)
            with st.spinner("Analyzing logo style…"):
                style_desc = describe_logo(image_bytes)
                st.write("**Logo style:**", style_desc)
        else:
            palette = ["#000000", "#444444", "#888888"]
        with st.spinner("Building PPTX…"):
            deck = build_deck(outline, palette, logo_file.read() if logo_file else None)
        with st.spinner("Building PDF…"):
            brief = render_brief(outline, palette)
        st.success("Your Deck and Brief are ready:")
        st.download_button("📥 Download Deck (PPTX)", deck, "deckly_deck.pptx")
        st.download_button("📥 Download Brief (PDF)", brief, "deckly_summary.pdf")

