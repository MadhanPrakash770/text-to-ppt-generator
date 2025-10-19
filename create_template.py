from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

# Create a new PowerPoint presentation
prs = Presentation()

# --- Title Slide ---
title_slide_layout = prs.slide_layouts[6]  # blank layout
slide = prs.slides.add_slide(title_slide_layout)

# Background color (blue gradient tone)
background = slide.background
fill = background.fill
fill.solid()
fill.fore_color.rgb = RGBColor(25, 50, 112)  # deep navy blue

# Add title text
title_box = slide.shapes.add_textbox(Inches(1.5), Inches(2.5), Inches(8), Inches(2))
title_frame = title_box.text_frame
title = title_frame.add_paragraph()
title.text = "AI Presentation Template"
title.font.bold = True
title.font.size = Pt(48)
title.font.color.rgb = RGBColor(255, 215, 0)  # gold accent

# Add subtitle text
subtitle_box = slide.shapes.add_textbox(Inches(1.5), Inches(4), Inches(8), Inches(1.5))
subtitle_frame = subtitle_box.text_frame
subtitle = subtitle_frame.add_paragraph()
subtitle.text = "Generated using OpenAI + Unsplash"
subtitle.font.size = Pt(28)
subtitle.font.color.rgb = RGBColor(255, 255, 255)

# --- Content Slide Layout ---
slide_layout = prs.slide_layouts[6]
slide = prs.slides.add_slide(slide_layout)

# Background (lighter gradient style)
background = slide.background
fill = background.fill
fill.solid()
fill.fore_color.rgb = RGBColor(230, 240, 255)  # soft light blue

# Placeholder for content (used in app.py)
content_box = slide.shapes.add_textbox(Inches(1), Inches(1.2), Inches(8), Inches(5))
frame = content_box.text_frame
frame.word_wrap = True
frame.text = "Content Slide Example"
frame.paragraphs[0].font.size = Pt(24)
frame.paragraphs[0].font.color.rgb = RGBColor(0, 51, 102)

# Save the template
prs.save("template.pptx")
print("✅ Created professional blue-gold template.pptx successfully.")
