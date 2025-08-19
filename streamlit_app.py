pip install streamlit python-pptx pillow openai anthropic
export OPENAI_API_KEY=sk-...     # or set ANTHROPIC_API_KEY instead
streamlit run app.py
# app.py
# Streamlit → one-input (Customer Name) → LLM-generated content → PPTX in your left-rail layout
# pip install streamlit python-pptx pillow openai anthropic

import io
import os
from datetime import datetime

import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

# ----------- Optional LLM clients -----------
OPENAI_AVAILABLE, ANTHROPIC_AVAILABLE = False, False
try:
    from openai import OpenAI
    OPENAI_AVAILABLE = True
except Exception:
    pass

try:
    import anthropic
    ANTHROPIC_AVAILABLE = True
except Exception:
    pass

# ----------- Streamlit UI -----------
st.set_page_config(page_title="Customer Update Slide Generator", page_icon="📊", layout="centered")
st.title("📊 One-Input Customer Update Slide (PPTX)")

with st.expander("How this works", expanded=False):
    st.markdown(
        """
        1) Enter a **customer name** (e.g., *Rugs USA*).  
        2) The app asks an LLM to draft 5 sections from *publicly available context*.  
        3) It renders your **left-rail** slide and lets you **download** the PPTX.  
        *Tip*: Set `OPENAI_API_KEY` (preferred) **or** `ANTHROPIC_API_KEY` in your environment.
        """
    )

# Inputs
customer = st.text_input("Customer name", "Rugs USA")
accent = st.color_picker("Accent color (vertical bar & rules)", value="#08574A")
logo = st.file_uploader("Optional: Upload customer logo (PNG)", type=["png"])
provider = st.selectbox("LLM provider", ["OpenAI (default)", "Anthropic", "None (manual edit)"])

# ----------- Prompt template -----------
SYSTEM_INSTRUCTIONS = """You are a precise enterprise research and writing assistant for Customer Success Managers.
Generate concise, presentation-ready bullets based solely on widely known public context for the named company.
Avoid speculation. Prefer short, useful bullets (6–14 words each)."""

USER_PROMPT_TMPL = """Company: {customer}

Task: Draft sections for an executive slide called "Customer Updates – Strategic Priorities & Supply Chain Context".
Keep it concise, factual, and presentation-ready. Use bullets where applicable. No fluff. No marketing tone.

Sections to return as strict JSON with these keys:
- corporate_vision: 2-3 sentences, clear and concise.
- business_strategies: 4-6 bullets.
- supply_chain_contribution: 4-6 bullets.
- risks_of_supply_chain_failure: 4-6 bullets.
- critical_capabilities: 4-6 bullets.

Rules:
- Base on publicly available, broadly reported info (brand positioning, ecom/retail ops patterns, logistics practices).
- If specifics are not well-documented, give conservative, generic-but-relevant bullets for a digital-first retailer
  (e.g., fulfillment speed, DC consolidation, marketplaces, designer collabs, WMS/OMS reliance), clearly worded.
- Do not include sources or URLs. Keep each bullet crisp.
Return only JSON.
"""

# ----------- LLM helpers -----------
def call_openai(customer_name: str):
    if not OPENAI_AVAILABLE:
        raise RuntimeError("OpenAI SDK not installed.")
    if not os.getenv("OPENAI_API_KEY"):
        raise RuntimeError("OPENAI_API_KEY not set.")
    client = OpenAI()
    prompt = USER_PROMPT_TMPL.format(customer=customer_name)
    resp = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": SYSTEM_INSTRUCTIONS},
            {"role": "user", "content": prompt},
        ],
        temperature=0.2,
        max_tokens=900,
    )
    return resp.choices[0].message.content

def call_anthropic(customer_name: str):
    if not ANTHROPIC_AVAILABLE:
        raise RuntimeError("Anthropic SDK not installed.")
    if not os.getenv("ANTHROPIC_API_KEY"):
        raise RuntimeError("ANTHROPIC_API_KEY not set.")
    client = anthropic.Anthropic()
    prompt = USER_PROMPT_TMPL.format(customer=customer_name)
    msg = client.messages.create(
        model="claude-3-5-sonnet-20240620",
        system=SYSTEM_INSTRUCTIONS,
        max_tokens=900,
        temperature=0.2,
        messages=[{"role": "user", "content": prompt}],
    )
    return msg.content[0].text

def safe_json_parse(s: str):
    import json, re
    # try plain JSON, then attempt to extract a fenced block
    try:
        return json.loads(s)
    except Exception:
        pass
    code_blocks = re.findall(r"```(?:json)?\s*(\{.*?\})\s*```", s, flags=re.S)
    for block in code_blocks:
        try:
            return json.loads(block)
        except Exception:
            continue
    # last resort: rough fixes
    try:
        s2 = s.strip()
        return json.loads(s2)
    except Exception as e:
        raise RuntimeError(f"Could not parse model output as JSON: {e}")

# ----------- PPTX drawing helpers -----------
def rgb_hex_to_tuple(hex_color: str):
    hex_color = hex_color.lstrip("#")
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))

def _add_logo(slide, file, width_in=1.6):
    try:
        slide.shapes.add_picture(file, Inches(11.4), Inches(0.3), width=Inches(width_in))
    except Exception:
        pass

def _add_rule(slide, y_in, color, thickness=2.0):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(y_in), Inches(12.3), Inches(0.02*thickness))
    shp.fill.solid(); shp.fill.fore_color.rgb = color
    shp.line.fill.background()

def _left_label(slide, text, y_in):
    box = slide.shapes.add_textbox(Inches(0.3), Inches(y_in), Inches(2.5), Inches(1.0))
    tf = box.text_frame; tf.clear()
    p = tf.paragraphs[0]; p.text = text
    p.font.size = Pt(24); p.font.bold = True
    p.font.color.rgb = RGBColor(0, 0, 0)

def _right_title_and_body(slide, title, body_lines, y_in):
    # Title
    tbox = slide.shapes.add_textbox(Inches(3.0), Inches(y_in), Inches(9.6), Inches(0.7))
    tf = tbox.text_frame; tf.clear()
    p = tf.paragraphs[0]; p.text = title
    p.font.size = Pt(22); p.font.bold = True

    # Body
    bbox = slide.shapes.add_textbox(Inches(3.0), Inches(y_in + 0.55), Inches(9.6), Inches(1.4))
    btf = bbox.text_frame; btf.clear()
    if isinstance(body_lines, str):
        # treat as paragraph
        p = btf.paragraphs[0]
        p.text = body_lines
        p.font.size = Pt(14)
        return
    for i, ln in enumerate(body_lines):
        p = btf.add_paragraph() if i else btf.paragraphs[0]
        p.text = ln
        p.font.size = Pt(14)
        p.level = 0
        p.space_after = Pt(2)

def build_slide_pptx(customer_name, accent_rgb, sections, logo_bytes=None):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.33), Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Header
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.35), Inches(9), Inches(0.9))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = "Customer Updates"
    p.font.size = Pt(44); p.font.bold = True

    # Optional logo
    if logo_bytes:
        _add_logo(slide, logo_bytes)

    # Accent vertical bar
    vbar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(2.8), Inches(1.2), Inches(0.06), Inches(5.8))
    vbar.fill.solid(); vbar.fill.fore_color.rgb = RGBColor(*accent_rgb)
    vbar.line.fill.background()

    # Section rows
    y0, row_h = 1.25, 1.4
    rule_color = RGBColor(165, 175, 190)

    rows = [
        ("Corporate\nVision", "Corporate Vision", sections["corporate_vision"]),
        ("Business\nStrategies", "Business Strategies", ["• " + s for s in sections["business_strategies"]]),
        ("Supply Chain\nContribution", "Supply Chain Contribution", ["• " + s for s in sections["supply_chain_contribution"]]),
        ("Risks of\nSupply Chain\nFailure", "Risks of Supply Chain Failure", ["• " + s for s in sections["risks_of_supply_chain_failure"]]),
        ("Critical\nCapabilities", "Critical Capabilities", ["• " + s for s in sections["critical_capabilities"]]),
    ]

    y = y0
    for rail, title, body in rows:
        _left_label(slide, rail, y - 0.05)
        _right_title_and_body(slide, title, body, y - 0.02)
        _add_rule(slide, y + row_h, rule_color, thickness=2.2)
        y += row_h

    # Footer (date/customer)
    foot = slide.shapes.add_textbox(Inches(0.5), Inches(6.9), Inches(12.3), Inches(0.4))
    ft = foot.text_frame; ft.clear()
    fp = ft.paragraphs[0]
    fp.text = f"{customer_name}  •  {datetime.now().strftime('%b %d, %Y')}"
    fp.font.size = Pt(12); fp.font.color.rgb = RGBColor(90, 90, 90)
    fp.alignment = PP_ALIGN.RIGHT

    return prs

# ----------- Generate button -----------
if st.button("Generate PPTX"):
    # 1) Get content from provider (or manual fallback)
    try:
        raw = None
        if provider.startswith("OpenAI") and os.getenv("OPENAI_API_KEY"):
            raw = call_openai(customer)
        elif provider.startswith("Anthropic") and os.getenv("ANTHROPIC_API_KEY"):
            raw = call_anthropic(customer)
        elif provider.startswith("None"):
            st.warning("No LLM selected — generating a blank editable skeleton.")
            raw = """{
              "corporate_vision": "Edit me: 2–3 sentence corporate vision.",
              "business_strategies": ["Edit me","Add bullets here"],
              "supply_chain_contribution": ["Edit me","Add bullets here"],
              "risks_of_supply_chain_failure": ["Edit me","Add bullets here"],
              "critical_capabilities": ["Edit me","Add bullets here"]
            }"""
        else:
            st.error("No API key found for the chosen provider. Set OPENAI_API_KEY or ANTHROPIC_API_KEY.")
            st.stop()

        data = safe_json_parse(raw)

        # 2) Build PPTX
        accent_rgb = rgb_hex_to_tuple(accent)
        logo_bytes = None
        if logo is not None:
            logo_bytes = logo

        prs = build_slide_pptx(customer, accent_rgb, data, logo_bytes)

        # 3) Download
        buf = io.BytesIO()
        prs.save(buf); buf.seek(0)
        st.success("PPTX generated!")
        st.download_button(
            "⬇️ Download PPTX",
            data=buf,
            file_name=f"{customer.replace(' ','_')}_Customer_Updates_{datetime.now().strftime('%Y%m%d')}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
    except Exception as e:
        st.error(f"Generation failed: {e}")

# ----------- Preview / Debug (optional) -----------
st.caption("Tip: Use OpenAI as the default provider for best results. You can upload a PNG logo for branding.")
