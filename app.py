import streamlit as st
import openai
import json
import io
import base64
from PIL import Image
import numpy as np
from sklearn.cluster import KMeans
from pptx import Presentation
from reportlab.pdfgen import canvas

# -- Configuration
st.set_page_config(page_title="Deckly – Pro-grade decks, zero hassle.", layout="wide")
openai.api_key = st.secrets["OPENAI_API_KEY"]

# -- LLM Prompt Chaining: Bullets to Structured Outline
@st.cache_data
def bullets_to_outline(bullets: str, tone: str = "Neutral") -> list[dict]:
    """
    Converts raw bullet text into a structured outline via LLM JSON output.
    Returns a list of sections, each a dict with 'title' and 'points'.
    """
    prompt = f"""TONE: {tone}
Convert the following bullet points into a structured JSON outline.
Output only a JSON array of sections with keys 'title' and 'points'.
BULLETS:
{bullets}"""
    try:
        response = openai.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "You are a world-class pitch-deck strategist."},
                {"role": "user", "content": prompt},
            ],
            temperature=0.3,
            response_format={"type": "json_object"}
        )
        outline = response.choices[0].message.json()
    except Exception as e:
        st.error(f"Failed to generate outline: {e}")
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
    response = openai.chat.completions.create(
        model="gpt-4o-mini",
        messages=[{"role": "user", "content": multimodal}],
        temperature=0.5,
    )
    return response.choices[0].message.content.strip()

# -- Slide Generation with python-pptx
from pptx.dml.color import RGBColor

@st.cache_data
def build_deck(outline: list[dict], palette: list[str], logo_bytes: bytes | None) -> bytes:
    prs = Presentation()
    # Convert hex colors to RGB tuples
    def hex_to_rgb(h):
        h = h.lstrip('#')
        return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))

    # Use first palette color as slide background
    bg_color = hex_to_rgb(palette[0]) if palette else (255, 255, 255)

    for section in outline:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        # Set slide background color
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor(*bg_color)
        # Title handling
        title_shape = slide.shapes.title
        title_text = section.get("title", "") or ""
        if title_shape:
            title_shape.text = title_text
        else:
            # Fallback: add textbox
            txBox = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(1))
            tf = txBox.text_frame
            tf.text = title_text
        # Content handling
        body_tf = None
        # Preferred: placeholder index 1
        try:
            body_tf = slide.placeholders[1].text_frame
        except Exception:
            # Fallback: first available text frame not title
            for shape in slide.shapes:
                if getattr(shape, 'has_text_frame', False) and shape != title_shape:
                    body_tf = shape.text_frame
                    break
        # Add bullet points
        if body_tf:
            # Clear any default text
            body_tf.clear()
            for point in section.get("points", []):
                p = body_tf.add_paragraph()
                p.text = point
                p.level = 0
        # Add logo if provided
        if logo_bytes:
            try:
                img_stream = io.BytesIO(logo_bytes)
                # Position: bottom-right corner
                pic = slide.shapes.add_picture(img_stream, prs.slide_width - Inches(1.2), prs.slide_height - Inches(1.2), width=Inches(1), height=Inches(1))
                pic.shadow.inherit = False
            except Exception:
                pass
    # Save presentation to bytes
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
