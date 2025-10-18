import os
import io
import base64
from flask import Flask, request, send_file, render_template, jsonify
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from openai import OpenAI
from dotenv import load_dotenv

# Load environment variables from .env
load_dotenv()

app = Flask(__name__)

# Initialize OpenAI client
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

# ======================================================
# 🔹 Generate Slide Text (Title + Bullets)
# ======================================================
def generate_slide_content(paragraph):
    prompt = (
        "You are a professional assistant that converts a paragraph into a PowerPoint slide.\n"
        "Generate:\n"
        "- A concise slide title (max 8 words)\n"
        "- Exactly 3 clear, concise bullet points summarizing the key ideas.\n\n"
        "Format strictly as:\n"
        "Slide Title:\n"
        "<Title>\n"
        "- Bullet 1\n"
        "- Bullet 2\n"
        "- Bullet 3\n\n"
        f"Paragraph:\n{paragraph}"
    )

    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=250,
            temperature=0.5,
        )
        text = response.choices[0].message.content.strip()
        print("✅ AI Response:", text)
        return text
    except Exception as e:
        print("❌ OpenAI text generation error:", e)
        return None


# ======================================================
# 🔹 Generate AI Image for Slide
# ======================================================
def generate_image(prompt):
    try:
        response = client.images.generate(
            model="gpt-image-1",
            prompt=f"Professional, creative PowerPoint illustration about {prompt}. "
                   f"Use abstract, futuristic visuals with blue tones, minimalistic design.",
            size="512x512"
        )
        image_b64 = response.data[0].b64_json
        return io.BytesIO(base64.b64decode(image_b64))
    except Exception as e:
        print("❌ Image generation error:", e)
        return None


# ======================================================
# 🔹 Flask Routes
# ======================================================
@app.route('/')
def home():
    return render_template('index.html')


@app.route('/generate', methods=['POST'])
def generate():
    data = request.get_json()
    text = data.get('text', '').strip()

    if not text:
        return jsonify({'error': 'No text provided'}), 400

    # Try to load the design template
    template_path = "template.pptx"
    if os.path.exists(template_path):
        prs = Presentation(template_path)
        print("🎨 Using custom template.pptx for design.")
        # Remove any existing template slides (like 'Your Slide Title')
        if len(prs.slides) > 0:
            xml_slides = prs.slides._sldIdLst
            for _ in range(len(prs.slides)):
                xml_slides.remove(xml_slides[0])
    else:
        prs = Presentation()
        print("⚠️ template.pptx not found — using blank layout.")

    paragraphs = [p.strip() for p in text.split('\n\n') if p.strip()]
    print(f"📄 Found {len(paragraphs)} paragraphs to convert into slides.")

    for idx, para in enumerate(paragraphs):
        ai_output = generate_slide_content(para)
        if not ai_output:
            continue

        # --- Parse AI output ---
        lines = [l.strip() for l in ai_output.splitlines() if l.strip()]
        title_text, bullets = "", []

        for line in lines:
            if line.lower().startswith("slide title:"):
                continue
            elif line.startswith("-"):
                bullets.append(line.lstrip("- ").strip())
            elif not title_text:
                title_text = line

        # Fix bullet merging if GPT puts all in one line
        if len(bullets) == 1 and " - " in bullets[0]:
            bullets = [b.strip() for b in bullets[0].split("-") if b.strip()]

        if not title_text:
            title_text = f"Slide {idx + 1}"

        # --- Create slide ---
        slide_layout = prs.slide_layouts[6]  # blank layout
        slide = prs.slides.add_slide(slide_layout)

        # --- Add Title ---
        title_box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(1.2))
        title_tf = title_box.text_frame
        p = title_tf.add_paragraph()
        p.text = title_text
        p.font.size = Pt(40)
        p.font.bold = True
        p.font.color.rgb = RGBColor(10, 50, 120)

        # --- Add Bullet Points ---
        content_box = slide.shapes.add_textbox(Inches(1.2), Inches(2.2), Inches(7.5), Inches(4))
        content_tf = content_box.text_frame
        for bullet in bullets:
            para_obj = content_tf.add_paragraph()
            para_obj.text = bullet
            para_obj.font.size = Pt(22)
            para_obj.font.color.rgb = RGBColor(40, 40, 40)
            para_obj.level = 0

        # --- Add AI Image ---
        img_stream = generate_image(title_text)
        if img_stream:
            try:
                slide.shapes.add_picture(img_stream, Inches(5.3), Inches(1.3), width=Inches(4), height=Inches(3))
            except Exception as e:
                print(f"⚠️ Could not add image on slide {idx}: {e}")

    # --- Save to BytesIO ---
    ppt_stream = io.BytesIO()
    prs.save(ppt_stream)
    ppt_stream.seek(0)

    print("✅ Presentation generation complete.")
    return send_file(
        ppt_stream,
        as_attachment=True,
        download_name='ai_generated_presentation.pptx',
        mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
    )


# ======================================================
# 🔹 Run Flask App
# ======================================================
if __name__ == '__main__':
    app.run(debug=True, port=5001)
