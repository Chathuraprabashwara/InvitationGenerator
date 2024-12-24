python3 -m venv env
source env/bin/activate


    input_pdf_path="input.pdf",      # Path to the input PDF with images
    output_pdf_path="output_with_text.pdf",  # Path to save the output PDF
    text="Overlay Text Example",            # Text to add
    position=(75, 284),                    # Coordinates (x, y)
    font_size=14,                           # Font size
    font_color=(1, 0, 0)                    # Red color (RGB)