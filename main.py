import os
import collections
import collections.abc
from pptx import Presentation
from pptx.util import Inches
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

import openai

openai.api_key = "my-api-key"


def generate_bullet_points(prompt, max_tokens=100):
    response = openai.Completion.create(
        engine="text-davinci-003",
        prompt=prompt,
        temperature=0.7,
        max_tokens=max_tokens,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0,
    )

    bullet_points = response.choices[0].text.strip().split("\n")
    return bullet_points


def add_slide(prs, title, bullet_points):
    slide_layout = prs.slide_layouts[5]  # Use the blank layout
    slide = prs.slides.add_slide(slide_layout)

    # Add a title to the slide
    title_shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(0.5), Inches(9), Inches(1))
    title_shape.text = title
    title_shape.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

    # Add bullet points
    bullet_shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(1.5), Inches(9), Inches(5))
    text_frame = bullet_shape.text_frame
    for idx, point in enumerate(bullet_points):
        if idx > 0:
            text_frame.add_paragraph()
        paragraph = text_frame.paragraphs[idx]
        paragraph.text = point
        paragraph.level = 0
        paragraph.space_after = Inches(0.1)

    return slide


# Create the presentation
prs = Presentation()

# Loop to add slides based on user input
while True:
    topic = input("Enter a topic for the next slide (or type 'exit' to finish): ")

    if topic.lower() == "exit":
        break

    # Generate bullet points using GPT-3
    prompt = f"Explain the topic '{topic}' and provide a few bullet points, not exceeding 90 words total"
    bullet_points = generate_bullet_points(prompt)

    slide = add_slide(prs, topic, bullet_points)
    print(f"Slide '{topic}' added with bullet points.")

# Save the presentation
presentation_name = "runtime_presentation.pptx"
prs.save(presentation_name)
os.startfile(presentation_name)  # Open the presentation in PowerPoint
