"""
AI description prompts for Moondream
"""

PROMPTS = {
    "general": """Describe this image for a blind viewer watching a recorded lecture.

Rules:
- Only describe what is actually visible in this specific image
- Use present tense
- Be objective - describe actions, not emotions or interpretations
- Identify people by their clothing and position, not by name
- Mention any visible text on screen
- If the screen is blank, black, or unclear, simply state that
- Keep to 2-3 clear sentences

Describe what you see:""",

    "slide": """Describe this presentation slide for a blind viewer.

Rules:
- Only describe what is actually visible on this slide
- State the title if one exists
- Read text content in logical order
- Describe any charts, graphs, or images
- If the slide is blank or the content is unclear, simply state that
- Keep to 3-4 clear sentences

What does this slide show?""",

    "slide_ocr": """This is a presentation slide. The text extracted from the slide reads:

{ocr_text}

Using this text and what you see, describe the slide for a blind viewer:
- Only describe what is actually present
- State the title if visible
- Read through content in logical order
- Describe any charts, graphs, or images
- If the slide appears blank or unclear, state that
- Keep to 3-4 clear sentences

Describe the slide:"""
}