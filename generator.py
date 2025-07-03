import json
import base64
from pptx import Presentation
from pptx.util import Emu, Pt, Inches
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.enum.dml import MSO_LINE
import requests
from io import BytesIO
from PIL import Image
import os
import re

def safe_int_conversion(value, default=0):
    if value is None or value == 'auto' or value == 'normal':
        return default
    try:
        if isinstance(value, str):
            numeric_part = re.search(r'[\d.]+', value)
            if numeric_part:
                return int(float(numeric_part.group()))
        return int(float(value))
    except (ValueError, TypeError):
        return default

def parse_color(color_str):
    if not color_str or color_str in ['transparent', 'rgba(0, 0, 0, 0)']:
        return None
    
    if color_str.startswith('rgb'):
        color_values = re.findall(r'\d+', color_str)
        if len(color_values) >= 3:
            return RGBColor(int(color_values[0]), int(color_values[1]), int(color_values[2]))
    
    elif color_str.startswith('#'):
        color_str = color_str.lstrip('#')
        if len(color_str) == 6:
            return RGBColor(int(color_str[0:2], 16), int(color_str[2:4], 16), int(color_str[4:6], 16))
        elif len(color_str) == 3:
            return RGBColor(int(color_str[0]*2, 16), int(color_str[1]*2, 16), int(color_str[2]*2, 16))
    
    named_colors = {
        'black': RGBColor(0, 0, 0),
        'white': RGBColor(255, 255, 255),
        'red': RGBColor(255, 0, 0),
        'green': RGBColor(0, 128, 0),
        'blue': RGBColor(0, 0, 255),
        'yellow': RGBColor(255, 255, 0),
        'gray': RGBColor(128, 128, 128),
        'grey': RGBColor(128, 128, 128),
    }
    
    return named_colors.get(color_str.lower())

def parse_border(border_str):
    if not border_str or border_str == 'none':
        return None, None, None
    
    parts = border_str.split()
    width = 0
    style = 'solid'
    color = None
    
    for part in parts:
        if 'px' in part:
            width = safe_int_conversion(part.replace('px', ''))
        elif part in ['solid', 'dashed', 'dotted']:
            style = part
        elif part.startswith('rgb') or part.startswith('#'):
            color = parse_color(part)
    
    if width == 0:
        match = re.match(r'(\d+)px\s+(\w+)\s+(.+)', border_str)
        if match:
            width = int(match.group(1))
            style = match.group(2)
            color = parse_color(match.group(3))
    
    return width, style, color

def parse_border_radius(radius_str):
    if not radius_str or radius_str == '0px':
        return 0
    
    try:
        match = re.search(r'[\d.]+', radius_str)
        if match:
            radius = float(match.group())
            return min(radius / 100, 1.0)
    except (ValueError, TypeError):
        return 0
    return 0

def pixels_to_pt(pixels):
    return pixels * 0.75

def scale_coordinates(x, y, width, height, source_width, source_height, target_width, target_height):
    if source_width <= 0 or source_height <= 0:
        return x, y, width, height
    
    scale_x = target_width / source_width
    scale_y = target_height / source_height
    
    return (
        int(x * scale_x),
        int(y * scale_y),
        int(width * scale_x),
        int(height * scale_y)
    )

def add_text_element(slide, element, ppt_width, ppt_height):
    x, y, width, height = element['x'], element['y'], element['width'], element['height']
    text = element.get('text', '').strip()
    
    if not text:
        return
    
    if x < 0 or y < 0 or x >= ppt_width or y >= ppt_height:
        print(f"Skipping text element - out of bounds: x={x}, y={y}")
        return
    
    width = min(width, ppt_width - x)
    height = min(height, ppt_height - y)
    
    if width <= 0 or height <= 0:
        print(f"Skipping text element - invalid dimensions")
        return
    
    try:
        textbox = slide.shapes.add_textbox(
            Emu(x * 9525), Emu(y * 9525),
            Emu(width * 9525), Emu(height * 9525)
        )
        
        text_frame = textbox.text_frame
        text_frame.clear()
        text_frame.text = text
        text_frame.word_wrap = True
        text_frame.auto_size = None
        
        paragraph = text_frame.paragraphs[0]
        styles = element.get('styles', {})
        
        font_size = safe_int_conversion(styles.get('fontSize', '12').replace('px', ''))
        if font_size > 0:
            paragraph.font.size = Pt(pixels_to_pt(font_size))
        
        font_family = styles.get('fontFamily', 'Arial')
        if font_family:
            paragraph.font.name = font_family
        
        font_weight = styles.get('fontWeight', 'normal')
        if font_weight == 'bold' or safe_int_conversion(font_weight) >= 700:
            paragraph.font.bold = True
        
        if styles.get('fontStyle') == 'italic':
            paragraph.font.italic = True
        
        text_color = parse_color(styles.get('color', 'black'))
        if text_color:
            paragraph.font.color.rgb = text_color
        
        text_align = styles.get('textAlign', 'left')
        alignment_map = {
            'left': PP_ALIGN.LEFT,
            'center': PP_ALIGN.CENTER,
            'right': PP_ALIGN.RIGHT,
            'justify': PP_ALIGN.JUSTIFY
        }
        paragraph.alignment = alignment_map.get(text_align, PP_ALIGN.LEFT)
        
        bg_color = parse_color(styles.get('backgroundColor'))
        if bg_color:
            fill = textbox.fill
            fill.solid()
            fill.fore_color.rgb = bg_color
        
        border_width, border_style, border_color = parse_border(styles.get('border'))
        border_radius = parse_border_radius(styles.get('borderRadius'))
        if border_width and border_width > 0:
            line = textbox.line
            line.width = Pt(border_width)
            if border_color:
                line.color.rgb = border_color
            if border_style == 'dashed':
                line.dash_style = MSO_LINE.DASH
        
        padding = styles.get('padding', '0px')
        if padding and padding != '0px':
            text_frame.margin_left = Emu(safe_int_conversion(padding.replace('px', '')) * 9525)
            text_frame.margin_right = Emu(safe_int_conversion(padding.replace('px', '')) * 9525)
            text_frame.margin_top = Emu(safe_int_conversion(padding.replace('px', '')) * 9525)
            text_frame.margin_bottom = Emu(safe_int_conversion(padding.replace('px', '')) * 9525)
        
        print(f"Added text: '{text[:30]}...' at ({x}, {y})")
        
    except Exception as e:
        print(f"Failed to add text element: {e}")

def add_shape_element(slide, element, ppt_width, ppt_height):
    x, y, width, height = element['x'], element['y'], element['width'], element['height']
    styles = element.get('styles', {})
    
    bg_color = parse_color(styles.get('backgroundColor'))
    border_width, border_style, border_color = parse_border(styles.get('border'))
    border_radius = parse_border_radius(styles.get('borderRadius'))
    
    if not bg_color and not border_width:
        return
    
    if x < 0 or y < 0 or x >= ppt_width or y >= ppt_height:
        print(f"Skipping shape element - out of bounds: x={x}, y={y}")
        return
    
    width = min(width, ppt_width - x)
    height = min(height, ppt_height - y)
    
    if width <= 0 or height <= 0:
        print(f"Skipping shape element - invalid dimensions")
        return
    
    try:
        shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if border_radius > 0 else MSO_SHAPE.RECTANGLE
        shape = slide.shapes.add_shape(
            shape_type,
            Emu(x * 9525), Emu(y * 9525),
            Emu(width * 9525), Emu(height * 9525)
        )

        if border_radius > 0 and shape_type == MSO_SHAPE.ROUNDED_RECTANGLE:
            shape.adjustments[0] = border_radius
        
        if bg_color:
            fill = shape.fill
            fill.solid()
            fill.fore_color.rgb = bg_color
        else:
            shape.fill.background()
        
        if border_width and border_width > 0:
            line = shape.line
            line.width = Pt(border_width)
            if border_color:
                line.color.rgb = border_color
            if border_style == 'dashed':
                line.dash_style = MSO_LINE.DASH
        else:
            shape.line.fill.background()
        
        print(f"Added shape (type: {shape_type}) at ({x}, {y}) with size ({width}, {height}), radius={border_radius}")
        
    except Exception as e:
        print(f"Failed to add shape element: {e}")

def add_image_element(slide, element, ppt_width, ppt_height):
    x, y, width, height = element['x'], element['y'], element['width'], element['height']
    img_src = element.get('src', '')
    
    if not img_src:
        return
    
    if x < 0 or y < 0 or x >= ppt_width or y >= ppt_height:
        print(f"Skipping image element - out of bounds: x={x}, y={y}")
        return
    
    width = min(width, ppt_width - x)
    height = min(height, ppt_height - y)
    
    if width <= 0 or height <= 0:
        print(f"Skipping image element - invalid dimensions")
        return
    
    try:
        temp_path = None
        
        if img_src.startswith('data:'):
            header, data = img_src.split(',', 1)
            img_data = base64.b64decode(data)
            temp_path = 'temp_image.png'
            with open(temp_path, 'wb') as f:
                f.write(img_data)
        
        elif img_src.startswith('http'):
            response = requests.get(img_src, timeout=10)
            if response.status_code == 200:
                temp_path = 'temp_image.png'
                with open(temp_path, 'wb') as f:
                    f.write(response.content)
            else:
                print(f"Failed to download image: {img_src}, status: {response.status_code}")
                return
        
        else:
            if os.path.exists(img_src):
                temp_path = img_src
            else:
                print(f"Image file not found: {img_src}")
                return

        try:
            with Image.open(temp_path) as img:
                img.verify()
        except Exception as e:
            print(f"Invalid image file: {temp_path}, error: {e}")
            if temp_path != img_src and os.path.exists(temp_path):
                os.remove(temp_path)
            return
        
        slide.shapes.add_picture(
            temp_path,
            Emu(x * 9525), Emu(y * 9525),
            Emu(width * 9525), Emu(height * 9525)
        )
        
        if temp_path != img_src and os.path.exists(temp_path):
            os.remove(temp_path)
        
        print(f"Added image: {img_src} at ({x}, {y})")
        
    except Exception as e:
        print(f"Failed to add image element: {e}")
        if temp_path and temp_path != img_src and os.path.exists(temp_path):
            os.remove(temp_path)

def create_pptx_from_json(json_path, output_path='output.pptx'):
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            slides_data = json.load(f)
    except Exception as e:
        print(f"Error reading JSON file: {e}")
        return
    
    prs = Presentation()
    ppt_width = 1920
    ppt_height = 1080
    prs.slide_width = Emu(ppt_width * 9525)
    prs.slide_height = Emu(ppt_height * 9525)
    
    print(f"Creating presentation with {len(slides_data)} slides")
    
    for slide_info in slides_data:
        slide_id = slide_info.get('slideId', 'Unknown')
        elements = slide_info.get('elements', [])
        
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        print(f"\nProcessing slide {slide_id} with {len(elements)} elements")
        
        source_width = elements[0].get('slideWidth', ppt_width) if elements else ppt_width
        source_height = elements[0].get('slideHeight', ppt_height) if elements else ppt_height
        
        def get_z_index(element):
            z_index = element.get('styles', {}).get('zIndex', '0')
            return safe_int_conversion(z_index, 0)
        
        elements_sorted = sorted(elements, key=get_z_index)
        
        for element in elements_sorted:
            if source_width != ppt_width or source_height != ppt_height:
                scaled_x, scaled_y, scaled_width, scaled_height = scale_coordinates(
                    element['x'], element['y'], element['width'], element['height'],
                    source_width, source_height, ppt_width, ppt_height
                )
                element['x'] = scaled_x
                element['y'] = scaled_y
                element['width'] = scaled_width
                element['height'] = scaled_height
            
            element_type = element.get('type', '').lower()
            
            if element_type == 'img':
                add_image_element(slide, element, ppt_width, ppt_height)
            elif element_type in ['div', 'span', 'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6'] and element.get('text'):
                add_text_element(slide, element, ppt_width, ppt_height)
            elif element_type == 'div':
                styles = element.get('styles', {})
                if (styles.get('backgroundColor') and styles['backgroundColor'] != 'rgba(0, 0, 0, 0)') or \
                   (styles.get('border') and styles['border'] != 'none') or \
                   (styles.get('borderRadius') and styles['borderRadius'] != '0px'):
                    add_shape_element(slide, element, ppt_width, ppt_height)
    
    try:
        prs.save(output_path)
        print(f"\nPowerPoint presentation saved successfully as '{output_path}'")
    except Exception as e:
        print(f"Error saving presentation: {e}")

if __name__ == "__main__":
    create_pptx_from_json('slides_data.json', 'output.pptx')