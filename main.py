import os
import win32com.client
import openai

openai.api_key = "my-api-key"


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

    # Generate bullet points using GPT-3
    prompt = f"Explain the topic '{topic}' and provide a 3-5 short bullet points, each bullet points starting with a " \
             f"new line (no need to start with a bullet) "
    bullet_points = generate_bullet_points(prompt)

    # Add a slide to the presentation
    slide = presentation.Slides.Add(num_slides, 2)

    # Set the slide title
    title_shape = slide.Shapes.Title
    title_shape.TextFrame.TextRange.Text = topic

    # Delete the text box shape from the slide
    text_box = slide.Shapes.Placeholders.Item(2)
    text_box.Delete()

    # Add bullet points to the slide
    text_box = slide.Shapes.AddTextbox(
        1, # Orientation
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
                paragraph = text_frame.TextRange.InsertAfter(point.lstrip("•") + "\n")
            else:
                paragraph = text_frame.TextRange.InsertAfter(point.lstrip("•"))
            paragraph.ParagraphFormat.Bullet.Type = 1

    num_slides += 1

# Save the presentation
presentation.SaveAs(os.path.join(os.getcwd(), "real_time_presentation.pptx"))

# Close the presentation
presentation.Close()

# Quit the PowerPoint application
PowerPointApp.Quit()
