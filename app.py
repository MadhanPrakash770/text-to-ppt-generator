import os
import io
from datetime import datetime
from flask import Flask, request, send_file, render_template, jsonify
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from openai import OpenAI
from dotenv import load_dotenv

load_dotenv()
app = Flask(__name__)
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

# =====================
# 🎨 DESIGN HELPERS
# =====================
def set_gradient_background(slide, color1, color2):
    bg = slide.background
    fill = bg.fill
    fill.gradient()
    stops = fill.gradient_stops
    stops[0].color.rgb = RGBColor(*color1)
    stops[1].color.rgb = RGBColor(*color2)

def add_footer(slide, text):
    shape = slide.shapes.add_shape(1, Inches(0), Inches(6.7), Inches(10), Inches(0.5))
    fill = shape.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(25, 45, 95)
    frame = shape.text_frame
    p = frame.add_paragraph()
    p.text = text
    p.font.size = Pt(12)
    p.font.color.rgb = RGBColor(255, 215, 0)
    p.alignment = PP_ALIGN.CENTER

def add_decorative_bar(slide, color):
    shape = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(0.4), Inches(7.5))
    fill = shape.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(*color)

def add_highlight_box(slide, left, top, width, height, color, transparency=0.15):
    box = slide.shapes.add_shape(1, Inches(left), Inches(top), Inches(width), Inches(height))
    fill = box.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(*color)
    fill.transparency = transparency
    box.line.fill.background()
    return box

# =====================
# 🧠 AI CONTENT
# =====================
def generate_slide_content(paragraph):
    prompt = (
        "Convert the following paragraph into a PowerPoint slide.\n"
        "Output format:\n"
        "Slide Title:\n<Title>\n"
        "- Bullet point 1\n- Bullet point 2\n- Bullet point 3\n\n"
        "Keep the bullets short (max 15 words each).\n"
        f"Paragraph:\n{paragraph}"
    )
    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=250,
            temperature=0.5,
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        print("❌ GPT Error:", e)
        return None

# =====================
# ✨ AUTO-FIT HELPERS
# =====================
def fit_title_text(title_box, text, max_font=38, min_font=24):
    """Automatically shrink title text to fit box width."""
    tf = title_box.text_frame
    tf.clear()
    p = tf.add_paragraph()
    p.text = text
    font_size = max_font

    # simulate shrink until under limit (approx.)
    while len(text) > 30 and font_size > min_font:
        font_size -= 2
        text = text[:60] + ("..." if len(text) > 60 else "")
    p.font.size = Pt(font_size)
    p.font.bold = True
    p.font.color.rgb = RGBColor(30, 50, 120)
    return tf

# =====================
# 🌐 FLASK ROUTES
# =====================
@app.route('/')
def home():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate():
    data = request.get_json()
    text = data.get("text", "").strip()
    if not text:
        return jsonify({"error": "No text provided"}), 400

    prs = Presentation()
    paragraphs = [p.strip() for p in text.split("\n\n") if p.strip()]

    # 🏆 Title Slide
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_gradient_background(slide, (10, 25, 70), (50, 80, 150))
    add_decorative_bar(slide, (255, 215, 0))

    title_box = slide.shapes.add_textbox(Inches(1.5), Inches(2.5), Inches(8), Inches(2))
    title_tf = title_box.text_frame
    title_p = title_tf.add_paragraph()
    title_p.text = "AI in Education – TEAM1 Project"
    title_p.font.size = Pt(54)
    title_p.font.bold = True
    title_p.font.color.rgb = RGBColor(255, 215, 0)
    title_p.alignment = PP_ALIGN.CENTER

    subtitle_box = slide.shapes.add_textbox(Inches(1.5), Inches(4), Inches(8), Inches(1))
    subtitle_tf = subtitle_box.text_frame
    subtitle_p = subtitle_tf.add_paragraph()
    subtitle_p.text = f"Presented by TEAM1 | {datetime.now().strftime('%B %Y')}"
    subtitle_p.font.size = Pt(24)
    subtitle_p.font.color.rgb = RGBColor(255, 255, 255)
    subtitle_p.alignment = PP_ALIGN.CENTER

    add_footer(slide, "AI Presentation Builder | TEAM1")

    # 📊 Content Slides
    for idx, para in enumerate(paragraphs):
        ai_text = generate_slide_content(para)
        if not ai_text:
            continue

        lines = ai_text.split("\n")
        title = next((l.replace("Slide Title:", "").strip() for l in lines if "Slide Title:" in l), f"Slide {idx+1}")
        bullets = [l.lstrip("- ").strip() for l in lines if l.strip().startswith("-")]

        slide = prs.slides.add_slide(prs.slide_layouts[6])
        set_gradient_background(slide, (25 + idx * 10, 45, 110 + idx * 10), (75, 120, 190))
        add_decorative_bar(slide, (255, 215, 0))
        add_highlight_box(slide, 0.9, 1.5, 8.2, 4.8, (255, 255, 255), transparency=0.85)

        # Title (auto fit)
        title_box = slide.shapes.add_textbox(Inches(1.3), Inches(0.8), Inches(8), Inches(1))
        title_tf = fit_title_text(title_box, title)

        # Bullets
        content_box = slide.shapes.add_textbox(Inches(1.6), Inches(2), Inches(8), Inches(4.5))
        tf = content_box.text_frame
        tf.word_wrap = True
        tf.auto_size = False

        for b in bullets:
            if len(b) > 90:
                b = b[:87] + "..."
            p = tf.add_paragraph()
            p.text = b
            p.font.size = Pt(22)
            p.font.color.rgb = RGBColor(25, 35, 75)
            p.space_after = Pt(12)
            p.alignment = PP_ALIGN.LEFT

        add_footer(slide, f"TEAM1 | {datetime.now().strftime('%B %Y')}")

    # 🙏 Thank You Slide
    thank_slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_gradient_background(thank_slide, (15, 40, 90), (60, 100, 180))
    add_decorative_bar(thank_slide, (255, 215, 0))

    title_box = thank_slide.shapes.add_textbox(Inches(1.5), Inches(2.5), Inches(8), Inches(2))
    title_tf = title_box.text_frame
    title_p = title_tf.add_paragraph()
    title_p.text = "THANK YOU"
    title_p.font.size = Pt(60)
    title_p.font.bold = True
    title_p.font.color.rgb = RGBColor(255, 215, 0)
    title_p.alignment = PP_ALIGN.CENTER

    subtitle_box = thank_slide.shapes.add_textbox(Inches(1.5), Inches(4.2), Inches(8), Inches(1))
    subtitle_tf = subtitle_box.text_frame
    subtitle_p = subtitle_tf.add_paragraph()
    subtitle_p.text = "Presented by TEAM1"
    subtitle_p.font.size = Pt(28)
    subtitle_p.font.color.rgb = RGBColor(255, 255, 255)
    subtitle_p.alignment = PP_ALIGN.CENTER

    add_footer(thank_slide, "AI Presentation Builder | TEAM1")

    # Save to memory
    pptx_stream = io.BytesIO()
    prs.save(pptx_stream)
    pptx_stream.seek(0)

    print("✅ TEAM1 Presentation Generated – Auto-fit titles + perfect layout!")

    return send_file(
        pptx_stream,
        as_attachment=True,
        download_name="TEAM1_Final_Presentation.pptx",
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )

if __name__ == "__main__":
    app.run(debug=True, port=5001)
