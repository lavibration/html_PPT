import os
import sys
import json
import asyncio
from playwright.async_api import async_playwright
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

# Configuration for 16:9 aspect ratio
WIDTH_IN = 13.333
HEIGHT_IN = 7.5
PX_TO_IN = WIDTH_IN / 1280  # Based on 1280px reference width

def parse_rgb(rgb_str):
    """Converts 'rgb(r, g, b)' or 'rgba(r, g, b, a)' to RGBColor."""
    if not rgb_str or 'rgba(0, 0, 0, 0)' in rgb_str:
        return None
    import re
    # Match both rgb and rgba
    match = re.search(r'rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*[\d\.]+)?\)', rgb_str)
    if match:
        return RGBColor(int(match.group(1)), int(match.group(2)), int(match.group(3)))
    return None

async def html_to_pptx(html_content, output_file="presentation.pptx"):
    prs = Presentation()
    # Set slide size to 16:9
    prs.slide_width = Inches(WIDTH_IN)
    prs.slide_height = Inches(HEIGHT_IN)

    async with async_playwright() as p:
        browser = await p.chromium.launch()
        page = await browser.new_page(viewport={'width': 1280, 'height': 720})

        # Inject Tailwind for consistent rendering if needed, or rely on provided styles
        full_html = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <script src="https://cdn.tailwindcss.com"></script>
            <style>
                body {{ margin: 0; padding: 0; overflow: hidden; }}
                .slide {{ width: 1280px; height: 720px; position: relative; overflow: hidden; background: white; }}
            </style>
        </head>
        <body>
            {html_content}
        </body>
        </html>
        """
        await page.set_content(full_html)
        await page.wait_for_timeout(500)  # Wait for styles/fonts

        # Find all slides (sections or elements with .slide class)
        slides = await page.query_selector_all('section, .slide')
        if not slides:
            # Fallback: treat body as a single slide if no explicit markers
            slides = [await page.query_selector('body')]

        for slide_el in slides:
            blank_slide_layout = prs.slide_layouts[6] # Blank
            slide = prs.slides.add_slide(blank_slide_layout)

            # Get bounding rect of the slide for absolute positioning
            slide_rect = await slide_el.bounding_box()

            # Query all elements inside the slide
            elements = await slide_el.query_selector_all('*')

            # Track processed text nodes to avoid duplicates
            processed_elements = set()

            for el in elements:
                # Get computed styles and bounding box
                style = await page.evaluate('''(el) => {
                    const s = window.getComputedStyle(el);
                    return {
                        fontSize: s.fontSize,
                        fontWeight: s.fontWeight,
                        color: s.color,
                        backgroundColor: s.backgroundColor,
                        textAlign: s.textAlign,
                        fontFamily: s.fontFamily,
                        display: s.display,
                        padding: s.padding,
                        borderWidth: s.borderWidth,
                        borderColor: s.borderColor,
                        borderRadius: s.borderRadius
                    };
                }''', el)

                rect = await el.bounding_box()
                if not rect or rect['width'] == 0 or rect['height'] == 0:
                    continue

                # Relative coordinates to the slide
                x = (rect['x'] - slide_rect['x']) * PX_TO_IN
                y = (rect['y'] - slide_rect['y']) * PX_TO_IN
                w = rect['width'] * PX_TO_IN
                h = rect['height'] * PX_TO_IN

                # Pass 1: Handle Background/Shapes
                bg_color = parse_rgb(style['backgroundColor'])

                # Handle multi-value border width (e.g. '0px 0px 0px 6px')
                border_raw = style['borderWidth'] or '0'
                # Remove 'px' and split by space
                border_values = [float(v.replace('px', '')) for v in border_raw.split() if v.strip()]
                border_width = max(border_values) if border_values else 0

                # Check if this element should be a shape or just a container
                has_visible_bg = bg_color is not None

                # Pass 2: Handle Images
                tag_name = await page.evaluate('(el) => el.tagName', el)
                if tag_name == 'IMG':
                    src = await page.evaluate('(el) => el.src', el)
                    if src.startswith('http'):
                        import requests
                        from io import BytesIO
                        try:
                            response = requests.get(src, timeout=5)
                            if response.status_code == 200:
                                image_stream = BytesIO(response.content)
                                slide.shapes.add_picture(image_stream, Inches(x), Inches(y), width=Inches(w), height=Inches(h))
                        except:
                            pass
                    continue

                # Pass 3: Handle Text
                # We check for direct text content
                text_content = await page.evaluate('''(el) => {
                    const childNodes = Array.from(el.childNodes);
                    const textNode = childNodes.find(n => n.nodeType === 3 && n.textContent.trim().length > 0);
                    return textNode ? el.innerText : null;
                }''', el)

                if text_content and el not in processed_elements:
                    # Add Text Box
                    txBox = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
                    tf = txBox.text_frame
                    tf.word_wrap = True

                    p = tf.paragraphs[0]
                    p.text = text_content.strip()

                    # Apply styles
                    # Handle possible complex font-size strings (though usually simple px)
                    fs_raw = style['fontSize'].split()[0].replace('px', '')
                    font_size = float(fs_raw) * 0.75 # px to pt
                    p.font.size = Pt(font_size)
                    p.font.name = 'Calibri'
                    p.font.bold = int(style['fontWeight']) >= 600 if style['fontWeight'].isdigit() else style['fontWeight'] == 'bold'

                    text_color = parse_rgb(style['color'])
                    if text_color:
                        p.font.color.rgb = text_color

                    # Alignment
                    if style['textAlign'] == 'center':
                        p.alignment = PP_ALIGN.CENTER
                    elif style['textAlign'] == 'right':
                        p.alignment = PP_ALIGN.RIGHT
                    else:
                        p.alignment = PP_ALIGN.LEFT

                    # Fill background if needed
                    if has_visible_bg:
                        txBox.fill.solid()
                        txBox.fill.fore_color.rgb = bg_color

                    # Mark children as processed to avoid duplicates in deep trees
                    children = await el.query_selector_all('*')
                    for child in children:
                        processed_elements.add(child)

                elif has_visible_bg:
                    # If it's just a colored box without direct text
                    from pptx.enum.shapes import MSO_SHAPE
                    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
                    shape.fill.solid()
                    shape.fill.fore_color.rgb = bg_color
                    shape.line.width = Pt(0) # Default no border

        await browser.close()

    prs.save(output_file)
    print(f"Presentation saved as {output_file}")

if __name__ == "__main__":
    # Get HTML from command line or environment variable
    html_input = os.environ.get('CLIENT_PAYLOAD_HTML', '<h1>Empty Slide</h1>')
    if len(sys.argv) > 1:
        html_input = sys.argv[1]

    asyncio.run(html_to_pptx(html_input))
