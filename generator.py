import json
import base64
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.dml import MSO_LINE
import requests
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

def safe_float_conversion(value, default=0.0):
    if value is None or value == 'auto' or value == 'normal':
        return default
    try:
        if isinstance(value, str):
            numeric_part = re.search(r'[\d.]+', value)
            if numeric_part:
                return float(numeric_part.group())
        return float(value)
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
            return min(radius / 20, 1.0)
    except (ValueError, TypeError):
        return 0
    return 0

def pixels_to_emu(pixels):
    return int(pixels * 9525)

def get_font_size_pt(font_size_px):
    if font_size_px <= 0:
        return 12
    return max(8, int(font_size_px * 0.8))

def add_text_element(slide, element, debug=False):
    x, y, width, height = element['x'], element['y'], element['width'], element['height']
    text = element.get('text', '').strip()
    
    if not text:
        return

    x = max(0, min(x, 1280))
    y = max(0, min(y, 720))
    width = max(10, min(width, 1280 - x))
    height = max(10, min(height, 720 - y))
    
    try:
        textbox = slide.shapes.add_textbox(
            pixels_to_emu(x), pixels_to_emu(y),
            pixels_to_emu(width), pixels_to_emu(height)
        )
        
        text_frame = textbox.text_frame
        text_frame.clear()
        text_frame.text = text

        text_frame.word_wrap = True
        text_frame.auto_size = None
        text_frame.vertical_anchor = MSO_ANCHOR.TOP

        text_frame.margin_left = 0
        text_frame.margin_right = 0
        text_frame.margin_top = 0
        text_frame.margin_bottom = 0
        
        paragraph = text_frame.paragraphs[0]
        styles = element.get('styles', {})

        font_size_px = safe_float_conversion(styles.get('fontSize', '12').replace('px', ''))
        if font_size_px > 0:
            font_size_pt = get_font_size_pt(font_size_px)
            paragraph.font.size = Pt(font_size_pt)
            if debug:
                print(f"Font size: {font_size_px}px -> {font_size_pt}pt")
        
        font_family = styles.get('fontFamily', 'Arial')
        if font_family:
            if ',' in font_family:
                font_family = font_family.split(',')[0].strip()
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
        else:
            textbox.fill.background()

        border_width, border_style, border_color = parse_border(styles.get('border'))
        if border_width and border_width > 0:
            line = textbox.line
            line.width = Pt(border_width)
            if border_color:
                line.color.rgb = border_color
            if border_style == 'dashed':
                line.dash_style = MSO_LINE.DASH
        else:
            textbox.line.fill.background()
        
        padding_top = safe_int_conversion(styles.get('paddingTop', '0').replace('px', ''))
        padding_right = safe_int_conversion(styles.get('paddingRight', '0').replace('px', ''))
        padding_bottom = safe_int_conversion(styles.get('paddingBottom', '0').replace('px', ''))
        padding_left = safe_int_conversion(styles.get('paddingLeft', '0').replace('px', ''))
        
        if padding_top > 0:
            text_frame.margin_top = pixels_to_emu(padding_top)
        if padding_right > 0:
            text_frame.margin_right = pixels_to_emu(padding_right)
        if padding_bottom > 0:
            text_frame.margin_bottom = pixels_to_emu(padding_bottom)
        if padding_left > 0:
            text_frame.margin_left = pixels_to_emu(padding_left)
        
        if debug:
            print(f"Added text: '{text[:50]}...' at ({x}, {y}) size ({width}x{height})")
        
    except Exception as e:
        print(f"Failed to add text element: {e}")
        print(f"Element data: x={x}, y={y}, width={width}, height={height}")

def add_shape_element(slide, element, debug=False):
    x, y, width, height = element['x'], element['y'], element['width'], element['height']
    styles = element.get('styles', {})
    
    bg_color = parse_color(styles.get('backgroundColor'))
    border_width, border_style, border_color = parse_border(styles.get('border'))
    border_radius = parse_border_radius(styles.get('borderRadius'))

    if not bg_color and not border_width:
        return

    x = max(0, min(x, 1280))
    y = max(0, min(y, 720))
    width = max(1, min(width, 1280 - x))
    height = max(1, min(height, 720 - y))
    
    try:
        shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if border_radius > 0 else MSO_SHAPE.RECTANGLE
        
        shape = slide.shapes.add_shape(
            shape_type,
            pixels_to_emu(x), pixels_to_emu(y),
            pixels_to_emu(width), pixels_to_emu(height)
        )

        if border_radius > 0 and shape_type == MSO_SHAPE.ROUNDED_RECTANGLE:
            try:
                shape.adjustments[0] = border_radius
            except:
                pass
        
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
        
        if debug:
            print(f"Added shape at ({x}, {y}) size ({width}x{height}), radius={border_radius}")
        
    except Exception as e:
        print(f"Failed to add shape element: {e}")

def add_image_element(slide, element, debug=False):
    x, y, width, height = element['x'], element['y'], element['width'], element['height']
    img_src = element.get('src', '')
    
    if not img_src:
        return

    x = max(0, min(x, 1280))
    y = max(0, min(y, 720))
    width = max(1, min(width, 1280 - x))
    height = max(1, min(height, 720 - y))
    
    try:
        temp_path = None

        if img_src.startswith('data:'):
            header, data = img_src.split(',', 1)
            img_data = base64.b64decode(data)
            temp_path = 'temp_image.png'
            with open(temp_path, 'wb') as f:
                f.write(img_data)

        elif img_src.startswith('http'):
            try:
                response = requests.get(img_src, timeout=10)
                if response.status_code == 200:
                    temp_path = 'temp_image.png'
                    with open(temp_path, 'wb') as f:
                        f.write(response.content)
                else:
                    print(f"Failed to download image: {img_src}, status: {response.status_code}")
                    return
            except Exception as e:
                print(f"Error downloading image {img_src}: {e}")
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
            pixels_to_emu(x), pixels_to_emu(y),
            pixels_to_emu(width), pixels_to_emu(height)
        )

        if temp_path != img_src and os.path.exists(temp_path):
            os.remove(temp_path)
        
        if debug:
            print(f"Added image: {img_src} at ({x}, {y}) size ({width}x{height})")
        
    except Exception as e:
        print(f"Failed to add image element: {e}")
        if temp_path and temp_path != img_src and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except:
                pass

def create_pptx_from_json(json_path, output_path='output.pptx', debug=False):
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            slides_data = json.load(f)
    except Exception as e:
        print(f"Error reading JSON file: {e}")
        return

    prs = Presentation()
    ppt_width = 1280
    ppt_height = 720
    prs.slide_width = pixels_to_emu(ppt_width)
    prs.slide_height = pixels_to_emu(ppt_height)
    
    print(f"Creating 720p presentation ({ppt_width}x{ppt_height}) with {len(slides_data)} slides")
    
    for slide_info in slides_data:
        slide_id = slide_info.get('slideId', 'Unknown')
        elements = slide_info.get('elements', [])

        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        if debug:
            print(f"\nProcessing slide {slide_id} with {len(elements)} elements")

        def get_z_index(element):
            return element.get('zIndex', 0)
        
        elements_sorted = sorted(elements, key=get_z_index)

        for element in elements_sorted:
            element_type = element.get('type', '').lower()
            
            if debug:
                print(f"Processing {element_type} at ({element['x']}, {element['y']}) size ({element['width']}x{element['height']})")

            if element_type == 'img':
                add_image_element(slide, element, debug)

            elif element_type == 'div':
                styles = element.get('styles', {})
                has_background = styles.get('backgroundColor') and styles['backgroundColor'] != 'rgba(0, 0, 0, 0)'
                has_border = styles.get('border') and styles['border'] != 'none'
                has_border_radius = styles.get('borderRadius') and styles['borderRadius'] != '0px'
                
                if has_background or has_border or has_border_radius:
                    add_shape_element(slide, element, debug)

                if element.get('text'):
                    add_text_element(slide, element, debug)
            
            elif element_type in ['span', 'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6'] and element.get('text'):
                add_text_element(slide, element, debug)
    
    try:
        prs.save(output_path)
        print(f"\n720p PowerPoint presentation saved successfully as '{output_path}'")
        print(f"Slide dimensions: {ppt_width}x{ppt_height} pixels")
    except Exception as e:
        print(f"Error saving presentation: {e}")

if __name__ == "__main__":
    create_pptx_from_json('slides_data.json', 'output_720p.pptx', debug=True)