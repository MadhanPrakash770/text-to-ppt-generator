from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# Create a blank presentation
prs = Presentation()

# Use a blank layout
blank_slide_layout = prs.slide_layouts[6]

# Create one styled slide
slide = prs.slides.add_slide(blank_slide_layout)

# === Background ===
background = slide.background
fill = background.fill
fill.solid()
fill.fore_color.rgb = RGBColor(235, 244, 255)  # light blue

# === Decorative top banner ===
banner = slide.shapes.add_shape(
    1,  # rectangle
    Inches(0), Inches(0),
    Inches(10), Inches(0.4)
)
banner.fill.solid()
banner.fill.fore_color.rgb = RGBColor(0, 102, 204)
banner.line.fill.background()

# === Title text ===
title_box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(1.2))
title_tf = title_box.text_frame
p = title_tf.add_paragraph()
p.text = "Your Slide Title"
p.font.size = Pt(44)
p.font.bold = True
p.font.color.rgb = RGBColor(10, 60, 150)
p.alignment = PP_ALIGN.LEFT

# === Bullet text ===
content_box = slide.shapes.add_textbox(Inches(1.2), Inches(2.5), Inches(7.5), Inches(4))
content_tf = content_box.text_frame

for text in ["- Bullet point example 1", "- Bullet point example 2", "- Bullet point example 3"]:
    p = content_tf.add_paragraph()
    p.text = text
    p.font.size = Pt(22)
    p.font.color.rgb = RGBColor(40, 40, 40)
    p.alignment = PP_ALIGN.LEFT

# Save file
prs.save("template.pptx")
print("✅ Created template.pptx successfully.")
