import os
import sys
import json
import asyncio
import re
from io import BytesIO
import requests
from playwright.async_api import async_playwright
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor

# Configuration for 16:9 aspect ratio
WIDTH_IN = 13.333
HEIGHT_IN = 7.5
PX_TO_IN = WIDTH_IN / 1280  # Based on 1280px reference width

def parse_rgb(rgb_str):
    """Converts 'rgb(r, g, b)' or 'rgba(r, g, b, a)' to RGBColor."""
    if not rgb_str or 'rgba(0, 0, 0, 0)' in rgb_str:
        return None
    # Match both rgb and rgba
    match = re.search(r'rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*[\d\.]+)?\)', rgb_str)
    if match:
        return RGBColor(int(match.group(1)), int(match.group(2)), int(match.group(3)))
    return None

def parse_px(px_str):
    """Safely converts '12px' to float 12.0."""
    if not px_str: return 0.0
    try:
        # Handle complex shorthand like '0px 0px 0px 6px'
        parts = [float(p.replace('px', '')) for p in px_str.split() if 'px' in p or p == '0']
        return max(parts) if parts else 0.0
    except:
        return 0.0

async def html_to_pptx(html_content, output_file="presentation.pptx"):
    # Ensure Playwright dependencies are installed if not already
    # For Streamlit Cloud, it's better to run 'playwright install' as a separate step

    prs = Presentation()
    # Set slide size to 16:9
    prs.slide_width = Inches(WIDTH_IN)
    prs.slide_height = Inches(HEIGHT_IN)

    async with async_playwright() as p:
        browser = await p.chromium.launch()
        page = await browser.new_page(viewport={'width': 1280, 'height': 720})

        full_html = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <script src="https://cdn.tailwindcss.com"></script>
            <style>
                body {{ margin: 0; padding: 0; overflow: hidden; font-family: Calibri, sans-serif; }}
                .slide {{ width: 1280px; height: 720px; position: relative; overflow: hidden; background: white; }}
            </style>
        </head>
        <body>
            {html_content}
        </body>
        </html>
        """
        await page.set_content(full_html)
        await page.wait_for_timeout(1000)  # Wait for Tailwind and fonts

        # Find all slides
        slides_elements = await page.query_selector_all('section, .slide')
        if not slides_elements:
            slides_elements = [await page.query_selector('body')]

        for slide_el in slides_elements:
            blank_slide_layout = prs.slide_layouts[6] # Blank
            slide = prs.slides.add_slide(blank_slide_layout)
            slide_rect = await slide_el.bounding_box()

            # Query all elements
            elements = await slide_el.query_selector_all('*')

            for el in elements:
                # Check if element was already processed (marked in DOM)
                is_processed = await page.evaluate('(el) => el.hasAttribute("data-pptx-processed")', el)
                if is_processed: continue

                # Get styles
                style = await page.evaluate('''(el) => {
                    const s = window.getComputedStyle(el);
                    return {
                        tagName: el.tagName,
                        fontSize: s.fontSize,
                        fontWeight: s.fontWeight,
                        color: s.color,
                        backgroundColor: s.backgroundColor,
                        textAlign: s.textAlign,
                        fontFamily: s.fontFamily,
                        display: s.display,
                        paddingTop: s.paddingTop,
                        paddingRight: s.paddingRight,
                        paddingBottom: s.paddingBottom,
                        paddingLeft: s.paddingLeft,
                        borderTopWidth: s.borderTopWidth,
                        borderRightWidth: s.borderRightWidth,
                        borderBottomWidth: s.borderBottomWidth,
                        borderLeftWidth: s.borderLeftWidth,
                        borderColor: s.borderColor,
                        borderTopColor: s.borderTopColor,
                        borderRightColor: s.borderRightColor,
                        borderBottomColor: s.borderBottomColor,
                        borderLeftColor: s.borderLeftColor,
                        borderRadius: s.borderRadius,
                        visibility: s.visibility,
                        opacity: s.opacity,
                        boxShadow: s.boxShadow,
                        textShadow: s.textShadow
                    };
                }''', el)

                rect = await el.bounding_box()
                if not rect or rect['width'] < 1 or rect['height'] < 1 or style['visibility'] == 'hidden' or float(style['opacity']) == 0:
                    continue

                x, y = (rect['x'] - slide_rect['x']) * PX_TO_IN, (rect['y'] - slide_rect['y']) * PX_TO_IN
                w, h = rect['width'] * PX_TO_IN, rect['height'] * PX_TO_IN

                # 1. Handle Images
                if style['tagName'] == 'IMG':
                    src = await page.evaluate('(el) => el.src', el)
                    if src.startswith('http'):
                        try:
                            response = requests.get(src, timeout=5)
                            if response.status_code == 200:
                                slide.shapes.add_picture(BytesIO(response.content), Inches(x), Inches(y), width=Inches(w), height=Inches(h))
                        except: pass
                    continue

                # 2. Extract Text Content
                text_content = await page.evaluate('''(el) => {
                    const childNodes = Array.from(el.childNodes);
                    const hasDirectText = childNodes.some(n => n.nodeType === 3 && n.textContent.trim().length > 0);
                    if (hasDirectText) return el.innerText.trim();
                    return null;
                }''', el)

                # 3. Background and Border logic
                bg_color = parse_rgb(style['backgroundColor'])
                border_radius = parse_px(style['borderRadius'])

                # Check for individual borders
                btw = parse_px(style['borderTopWidth'])
                brw = parse_px(style['borderRightWidth'])
                bbw = parse_px(style['borderBottomWidth'])
                blw = parse_px(style['borderLeftWidth'])

                is_uniform_border = (btw == brw == bbw == blw) and btw > 0
                has_visible_bg = bg_color is not None

                # Create a shape only if it has a visible effect or text
                if text_content or has_visible_bg or is_uniform_border:
                    shape_type = MSO_SHAPE.RECTANGLE
                    if border_radius > 0:
                        shape_type = MSO_SHAPE.ROUNDED_RECTANGLE

                    shape = slide.shapes.add_shape(shape_type, Inches(x), Inches(y), Inches(w), Inches(h))

                    # DISABLE ALL DEFAULT SHADOWS
                    shape.shadow.inherit = False

                    # Fill
                    if has_visible_bg:
                        shape.fill.solid()
                        shape.fill.fore_color.rgb = bg_color
                    else:
                        shape.fill.background()

                    # Outline (Uniform Border)
                    if is_uniform_border:
                        b_color = parse_rgb(style['borderTopColor']) or parse_rgb(style['borderColor'])
                        if b_color:
                            shape.line.color.rgb = b_color
                            shape.line.width = Pt(btw * 0.75)
                        else:
                            shape.line.fill.background()
                    else:
                        # CRITICAL: Prevents default blue/white border
                        shape.line.fill.background()
                        shape.line.width = Pt(0)

                    # Border Radius adjustment
                    if border_radius > 0 and shape_type == MSO_SHAPE.ROUNDED_RECTANGLE:
                        min_dim = min(rect['width'], rect['height'])
                        adj_val = min(0.5, border_radius / min_dim) if min_dim > 0 else 0
                        shape.adjustments[0] = adj_val

                    # Text Frame Logic
                    if text_content:
                        tf = shape.text_frame
                        tf.word_wrap = True
                        tf.margin_top = Inches(parse_px(style['paddingTop']) * PX_TO_IN)
                        tf.margin_right = Inches(parse_px(style['paddingRight']) * PX_TO_IN)
                        tf.margin_bottom = Inches(parse_px(style['paddingBottom']) * PX_TO_IN)
                        tf.margin_left = Inches(parse_px(style['paddingLeft']) * PX_TO_IN)

                        p = tf.paragraphs[0]
                        p.text = text_content

                        # Font Style
                        fs_raw = style['fontSize'].split()[0].replace('px', '')
                        p.font.size = Pt(float(fs_raw) * 0.75)
                        p.font.name = 'Calibri'
                        p.font.bold = int(style['fontWeight']) >= 600 if style['fontWeight'].isdigit() else style['fontWeight'] == 'bold'

                        # DISABLE TEXT SHADOW
                        p.font.shadow = False

                        t_color = parse_rgb(style['color'])
                        if t_color: p.font.color.rgb = t_color

                        if style['textAlign'] == 'center': p.alignment = PP_ALIGN.CENTER
                        elif style['textAlign'] == 'right': p.alignment = PP_ALIGN.RIGHT
                        else: p.alignment = PP_ALIGN.LEFT

                        # Block child elements from double-rendering
                        # We mark them in the DOM since Python ElementHandles are not reliably hashable/comparable
                        await page.evaluate('(el) => { el.querySelectorAll("*").forEach(c => c.setAttribute("data-pptx-processed", "true")); }', el)

                # 4. Handle non-uniform borders as separate lines
                if not is_uniform_border:
                    border_configs = [
                        (btw, style['borderTopColor'], x, y, w, 0),    # Top
                        (brw, style['borderRightColor'], x + w, y, 0, h), # Right
                        (bbw, style['borderBottomColor'], x, y + h, w, 0), # Bottom
                        (blw, style['borderLeftColor'], x, y, 0, h)     # Left
                    ]
                    for b_width, b_color_str, lx, ly, lw, lh in border_configs:
                        if b_width > 0:
                            b_color = parse_rgb(b_color_str)
                            if b_color:
                                # Use a small rectangle as a line
                                line_w = max(lw, b_width * PX_TO_IN) if lh > 0 else lw
                                line_h = max(lh, b_width * PX_TO_IN) if lw > 0 else lh
                                line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(lx), Inches(ly), Inches(line_w), Inches(line_h))
                                line.shadow.inherit = False
                                line.fill.solid()
                                line.fill.fore_color.rgb = b_color
                                line.line.fill.background()
                                line.line.width = Pt(0)

        await browser.close()

    prs.save(output_file)
    print(f"Presentation saved as {output_file}")

if __name__ == "__main__":
    html_input = os.environ.get('CLIENT_PAYLOAD_HTML', '<h1>Empty Slide</h1>')
    output_file = "presentation.pptx"

    if len(sys.argv) > 1:
        if len(sys.argv) > 2:
            html_input = sys.argv[1]
            output_file = sys.argv[2]
        else:
            if sys.argv[1].endswith('.pptx'):
                output_file = sys.argv[1]
            else:
                html_input = sys.argv[1]

    asyncio.run(html_to_pptx(html_input, output_file))
