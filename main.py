import os
import random

import win32com.client
import openai

with open("api_key.txt", "r") as f:
    openai.api_key = f.read().strip()


def generate_title(prompt, max_tokens=40):
    response = openai.Completion.create(
        engine="text-davinci-003",
        prompt=prompt,
        temperature=0.7,
        max_tokens=max_tokens,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0,
    )

    return response.choices[0].text


def generate_bullet_points(prompt, max_tokens=300):
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


# Start an instance of PowerPoint
PowerPointApp = win32com.client.Dispatch("PowerPoint.Application")

# Make the PowerPoint application visible
PowerPointApp.Visible = True

# Create a new presentation
presentation = PowerPointApp.Presentations.Add()

num_slides = 1

while True:
    # Get user input for the slide topic
    topic = input("Enter a topic for the next slide (or type 'exit' to finish): ")

    if topic.lower() == "exit":
        break

    prompt = f"Formulate the question: \"'{topic}'\" as a nice title (don't make it too formal) for a slide in a presentation"
    title = generate_title(prompt).replace('"', '')

    num_bullets = random.randint(2, 4)
    prompt = f"Explain the topic \"'{title}'\" in short by {num_bullets}-{num_bullets + 1} short bullet points with " \
             f"interesting information on the subject. "
    bullet_points = generate_bullet_points(prompt)

    # Add a slide to the presentation
    slide = presentation.Slides.Add(num_slides, 2)

    # Set the slide title
    title_shape = slide.Shapes.Title
    title_shape.TextFrame.TextRange.Text = title

    # Delete the text box shape from the slide
    text_box = slide.Shapes.Placeholders.Item(2)
    text_box.Delete()

    title_shape.Top = 0  # Move the title higher, adjust this value as needed
    title_shape.Width = 900  # Set the width to the slide width
    title_shape.Height = 40  # Set the height, adjust as needed
    title_shape.TextFrame.WordWrap = False  # Disable word wrapping
    title_shape.TextFrame.AutoSize = 0  # Disable auto resizing
    title_shape.TextFrame.MarginLeft = 0  # Remove left margin
    title_shape.TextFrame.MarginRight = 0  # Remove right margin

    # Reduce the font of the title if it exceeds the slide width
    title_text = title
    while title_shape.TextFrame.TextRange.BoundWidth > title_shape.Width:
        title_shape.TextFrame.TextRange.Font.Size -= 2

    # Add bullet points to the slide
    text_box = slide.Shapes.AddTextbox(
        1,  # Orientation
        100,  # Left
        100,  # Top
        400,  # Width
        100,  # Height
    )
    text_frame = text_box.TextFrame

    paragraph = text_frame.TextRange
    for idx, point in enumerate(bullet_points):
        if point.startswith("•"):
            if idx < len(bullet_points) - 1:
                paragraph = text_frame.TextRange.InsertAfter(point.lstrip("•").lstrip() + "\n")
            else:
                paragraph = text_frame.TextRange.InsertAfter(point.lstrip("•").lstrip())
            paragraph.ParagraphFormat.Bullet.Type = 1

    num_slides += 1

# Save the presentation
presentation.SaveAs(os.path.join(os.getcwd(), "real_time_presentation.pptx"))

# Close the presentation
presentation.Close()

# Quit the PowerPoint application
PowerPointApp.Quit()
