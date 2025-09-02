import re
import os
import base64
import json
import math
import requests
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.dml import MSO_LINE, MSO_COLOR_TYPE
import requests
from PIL import Image, ImageDraw
import os
import re
import math
from pptx.oxml.xmlchemy import OxmlElement
from pptx.oxml.ns import qn
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
import json
import re
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.chart import XL_LEGEND_POSITION
from PIL import Image
import os

def SubElement(parent, tagname, **kwargs):
    element = OxmlElement(tagname)
    element.attrib.update(kwargs)
    parent.append(element)
    return element


def _set_cell_border(cell, border_color="000000", border_width='12700'):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for lines in ['a:lnL','a:lnR','a:lnT','a:lnB']:
        # Every time before a node is inserted, the nodes with the same tag should be removed.
        tag = lines.split(":")[-1]
        for e in tcPr.getchildren():
            if tag in str(e.tag):
                tcPr.remove(e)
        # end
        ln = SubElement(tcPr, lines, w=border_width, cap='flat', cmpd='sng', algn='ctr')
        solidFill = SubElement(ln, 'a:solidFill')
        srgbClr = SubElement(solidFill, 'a:srgbClr', val=border_color)
        prstDash = SubElement(ln, 'a:prstDash', val='solid')
        round_ = SubElement(ln, 'a:round')
        headEnd = SubElement(ln, 'a:headEnd', type='none', w='med', len='med')
        tailEnd = SubElement(ln, 'a:tailEnd', type='none', w='med', len='med')
    return cell

def set_cell_border_enhanced(cell, side, width_px, color_rgb, style='solid'):
    """Enhanced cell border setting using the new SubElement approach"""
    if width_px <= 0 or not color_rgb:
        return

    # Convert color to hex string
    try:
        color_hex = f'{color_rgb.r:02X}{color_rgb.g:02X}{color_rgb.b:02X}'
    except:
        color_hex = "000000"
    
    # Convert width to EMU (1pt = 12700 EMU, 1px ≈ 0.75pt)
    width_emu = str(int(width_px * 0.75 * 12700))
    
    # Map side to line element
    side_tag_map = {'left': 'a:lnL', 'right': 'a:lnR', 'top': 'a:lnT', 'bottom': 'a:lnB'}
    ln_tag = side_tag_map.get(side.lower())
    if not ln_tag:
        return

    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    
    # Remove existing border if present
    existing_border = tcPr.find(qn(ln_tag))
    if existing_border is not None:
        tcPr.remove(existing_border)
    
    # Create new border
    ln = SubElement(tcPr, ln_tag, w=width_emu, cap='flat', cmpd='sng', algn='ctr')
    solidFill = SubElement(ln, 'a:solidFill')
    srgbClr = SubElement(solidFill, 'a:srgbClr', val=color_hex)
    
    # Set dash style based on border style
    dash_val = 'solid'
    if style == 'dashed':
        dash_val = 'dash'
    elif style == 'dotted':
        dash_val = 'sysDot'
    
    prstDash = SubElement(ln, 'a:prstDash', val=dash_val)
    round_ = SubElement(ln, 'a:round')
    headEnd = SubElement(ln, 'a:headEnd', type='none', w='med', len='med')
    tailEnd = SubElement(ln, 'a:tailEnd', type='none', w='med', len='med')

def set_cell_border(cell, side, width_px, color_rgb, style='solid'):
    """Apply a border to a table cell side using enhanced method"""
    set_cell_border_enhanced(cell, side, width_px, color_rgb, style)

def safe_int(value, default=0):
    try:
        return int(float(value))
    except (ValueError, TypeError):
        return default
def safe_float(value, default=0.0):
    """Safely convert value to float with better error handling"""
    try:
        if isinstance(value, str):
            value = value.replace('px', '').replace('%', '').replace('em', '').replace('rem', '')
        return float(value) if value else default
    except (ValueError, TypeError):
        return default
def pixels_to_emu(pixels):
    """Convert pixels to EMU with high precision (96 PPI)"""
    return int(round(pixels * 9525))
def px_to_pt(px):
    """Convert pixels to points"""
    return px * 0.75
def get_font_size_pt(font_size_px):
    if font_size_px <= 0:
        return 12
    return max(6, int(px_to_pt(font_size_px)))
def parse_border_radius(radius_str, shape_width_px, shape_height_px):
    """Parse border radius from CSS string"""
    if not radius_str or radius_str == '0px':
        return 0
    try:
        if isinstance(radius_str, str):
            radius_px = safe_float(radius_str.replace('px', '').replace('%', ''))
        else:
            radius_px = safe_float(radius_str)
        min_dimension = min(shape_width_px, shape_height_px)
        if min_dimension > 0:
            return min(0.5, radius_px / min_dimension)
        return 0
    except (ValueError, TypeError):
        return 0
def parse_color(color_str):
    """Enhanced color parsing with better RGB extraction"""
    if not color_str or color_str in ['transparent', 'rgba(0, 0, 0, 0)', 'none', 'initial', 'inherit']:
        return None
   
    # Handle linear-gradient
    if color_str.startswith('linear-gradient'):
        # For gradients, take the first color as approximation
        match = re.search(r'rgb[a]?\(\d+,\s*\d+,\s*\d+', color_str)
        if match:
            return parse_color(match.group())
   
    # Handle RGB/RGBA
    if color_str.startswith(('rgb', 'rgba')):
        parts = re.findall(r'[\d.]+', color_str)
        if len(parts) >= 3:
            return RGBColor(min(255, int(float(parts[0]))), min(255, int(float(parts[1]))), min(255, int(float(parts[2]))))
   
    # Handle hex colors
    elif color_str.startswith('#'):
        color_str = color_str.lstrip('#')
        if len(color_str) == 6:
            return RGBColor(int(color_str[0:2], 16), int(color_str[2:4], 16), int(color_str[4:6], 16))
        elif len(color_str) == 3:
            return RGBColor(int(color_str[0]*2, 16), int(color_str[1]*2, 16), int(color_str[2]*2, 16))
   
    # Named colors
    named_colors = {
        'black': RGBColor(0, 0, 0), 'white': RGBColor(255, 255, 255),
        'red': RGBColor(255, 0, 0), 'green': RGBColor(0, 128, 0),
        'blue': RGBColor(0, 0, 255), 'yellow': RGBColor(255, 255, 0),
        'gray': RGBColor(128, 128, 128), 'grey': RGBColor(128, 128, 128),
        'silver': RGBColor(192, 192, 192), 'maroon': RGBColor(128, 0, 0),
        'olive': RGBColor(128, 128, 0), 'lime': RGBColor(0, 255, 0),
        'aqua': RGBColor(0, 255, 255), 'teal': RGBColor(0, 128, 128),
        'navy': RGBColor(0, 0, 128), 'fuchsia': RGBColor(255, 0, 255),
        'purple': RGBColor(128, 0, 128)
    }
   
    return named_colors.get(color_str.lower())
def is_uniform_border(styles):
    widths = [safe_float(styles.get(f'border{side}Width', '0px')) for side in ['Top', 'Right', 'Bottom', 'Left']]
    styles_list = [styles.get(f'border{side}Style', 'none') for side in ['Top', 'Right', 'Bottom', 'Left']]
    colors = [styles.get(f'border{side}Color', '') for side in ['Top', 'Right', 'Bottom', 'Left']]
    return len(set(widths)) == 1 and len(set(styles_list)) == 1 and len(set(colors)) == 1 and widths[0] > 0
def has_any_border(styles):
    """Enhanced border detection"""
    for side in ['Top', 'Right', 'Bottom', 'Left']:
        width_key = f'border{side}Width'
        style_key = f'border{side}Style'
        color_key = f'border{side}Color'
       
        width = safe_float(styles.get(width_key, '0px'))
        style = styles.get(style_key, 'none')
        color = parse_color(styles.get(color_key, ''))
       
        if width > 0 and style not in ['none', 'hidden'] and color is not None:
            return True
    return False
def get_border_info(styles):
    """Enhanced border information extraction"""
    border_info = {}
    for side in ['Top', 'Right', 'Bottom', 'Left']:
        width_key = f'border{side}Width'
        style_key = f'border{side}Style'
        color_key = f'border{side}Color'
       
        width = safe_float(styles.get(width_key, '0px'))
        style = styles.get(style_key, 'none')
        color = parse_color(styles.get(color_key, ''))
       
        border_info[side.lower()] = {
            'width': width,
            'style': style,
            'color': color,
            'has_border': width > 0 and style not in ['none', 'hidden'] and color is not None
        }
    return border_info
def make_rounded_image(input_path, output_path, radius):
    im = Image.open(input_path).convert("RGBA")
    mask = Image.new("L", im.size, 0)
    draw = ImageDraw.Draw(mask)
    draw.rounded_rectangle((0, 0) + im.size, radius=radius, fill=255)
    im.putalpha(mask)
    im.save(output_path, "PNG")
def create_precise_border_shapes(slide, x, y, width, height, border_info, border_radius=0):
    """Create precise border shapes with accurate positioning"""
    shapes_created = []
    # Always check and draw all four borders if needed
    for side, info in border_info.items():
        if not info['has_border']:
            continue
        border_width = info['width']
        border_color = info['color']
        border_style = info['style']
        # Calculate precise line positions
        if side == 'top':
            line_x = x
            line_y = y
            line_width = width
            line_height = border_width
        elif side == 'right':
            line_x = x + width - border_width
            line_y = y
            line_width = border_width
            line_height = height
        elif side == 'bottom':
            line_x = x
            line_y = y + height - border_width
            line_width = width
            line_height = border_width
        elif side == 'left':
            line_x = x
            line_y = y
            line_width = border_width
            line_height = height
        try:
            # Always use rectangle for side borders, rounded only for full shape
            shape_type = MSO_SHAPE.RECTANGLE
            border_shape = slide.shapes.add_shape(
                shape_type,
                pixels_to_emu(line_x), pixels_to_emu(line_y),
                pixels_to_emu(line_width), pixels_to_emu(line_height)
            )
            border_shape.fill.solid()
            border_shape.fill.fore_color.rgb = border_color
            border_shape.line.fill.background()
            border_shape.shadow.inherit = False
            if border_style == 'dashed':
                border_shape.line.dash_style = MSO_LINE.DASH
            elif border_style == 'dotted':
                border_shape.line.dash_style = MSO_LINE.ROUND_DOT
            shapes_created.append(border_shape)
        except Exception as e:
            print(f"Error creating {side} border: {e}")
    return shapes_created
def add_bg_shape(slide, styles, x, y, width, height):
    """Enhanced background shape creation with precise positioning"""
    bg_color = parse_color(styles.get('backgroundColor'))
    border_radius_str = styles.get('borderRadius', '0px')
    border_radius = parse_border_radius(border_radius_str, width, height)
    has_radius = border_radius > 0
    box_shadow = styles.get('boxShadow', 'none')
    has_shadow = box_shadow != 'none'
    has_uniform_border = is_uniform_border(styles)
    has_any_border_sides = has_any_border(styles)
    shapes_created = []
    # Create main background shape
    if bg_color or has_uniform_border or has_radius or has_shadow:
        shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if has_radius else MSO_SHAPE.RECTANGLE
        try:
            bg_shape = slide.shapes.add_shape(
                shape_type,
                pixels_to_emu(x), pixels_to_emu(y),
                pixels_to_emu(width), pixels_to_emu(height)
            )
            if has_radius:
                bg_shape.adjustments[0] = border_radius
            if bg_color:
                bg_shape.fill.solid()
                bg_shape.fill.fore_color.rgb = bg_color
            else:
                bg_shape.fill.background()
            bg_shape.shadow.inherit = False
            # Handle uniform borders
            if has_uniform_border:
                border_width = safe_float(styles.get('borderTopWidth', '1px'))
                border_color = parse_color(styles.get('borderTopColor', 'black'))
                border_style = styles.get('borderTopStyle', 'solid')
                if border_color:
                    bg_shape.line.width = Pt(max(0.5, border_width))
                    bg_shape.line.color.rgb = border_color
                    if border_style == 'dashed':
                        bg_shape.line.dash_style = MSO_LINE.DASH
                    elif border_style == 'dotted':
                        bg_shape.line.dash_style = MSO_LINE.ROUND_DOT
                else:
                    bg_shape.line.fill.background()
            else:
                bg_shape.line.fill.background()
            if has_shadow:
                apply_shadow(bg_shape, box_shadow)
            shapes_created.append(bg_shape)
        except Exception as e:
            print(f"Error adding bg shape: {e}")
    # Always handle non-uniform borders for all sides
    if has_any_border_sides and not has_uniform_border:
        border_info = get_border_info(styles)
        border_shapes = create_precise_border_shapes(
            slide, x, y, width, height, border_info,
            border_radius * min(width, height) if has_radius else 0
        )
        shapes_created.extend(border_shapes)
    return shapes_created
def apply_shadow(shape, box_shadow_str):
    if box_shadow_str == 'none':
        return
    parts = box_shadow_str.split()
    if len(parts) < 3:
        return
    offset_x_px = safe_float(parts[0])
    offset_y_px = safe_float(parts[1])
    blur_px = safe_float(parts[2])
    spread_px = safe_float(parts[3]) if len(parts) > 3 else 0
    color_str = parts[4] if len(parts) > 4 else parts[3]
    if offset_x_px == 0 and offset_y_px == 0 and blur_px == 0 and spread_px == 0:
        return
    color = parse_color(color_str)
    alpha = 1.0
    if 'rgba' in color_str:
        color_parts = re.findall(r'\d+', color_str)
        if len(color_parts) == 4:
            alpha = float(color_parts[3]) / 255
    if not color:
        return
    distance_px = math.sqrt(offset_x_px**2 + offset_y_px**2)
    direction = math.degrees(math.atan2(offset_y_px, offset_x_px)) if distance_px > 0 else 0
    shape.shadow.inherit = False
    shape.shadow.blur = Pt(px_to_pt(blur_px))
    shape.shadow.distance = Pt(px_to_pt(distance_px))
    shape.shadow.angle = direction
    shape.shadow.color.type = MSO_COLOR_TYPE.RGB
    shape.shadow.color.rgb = color
    shape.shadow.transparency = 1 - alpha
def add_inline_group_element(slide, element, slide_width, slide_height, parent_has_shadow=False):
    inline_group = element.get('inlineGroup')
    if not inline_group:
        return
    inline_elements = inline_group.get('inlineElements', [])
    if not inline_elements:
        # Check if this is just a red highlight wrapper without text
        has_red_highlight = inline_group.get('hasRedHighlight', False)
        if not has_red_highlight:
            return
    
    has_content = any(elem.get('text', '').strip() for elem in inline_elements)
    if not has_content and not inline_group.get('hasRedHighlight', False):
        return
    
    group_rect = inline_group.get('groupRect', {})
    x = max(0, min(group_rect.get('x', 0), slide_width - 10))
    y = max(0, min(group_rect.get('y', 0), slide_height - 10))
    width = max(10, min(group_rect.get('width', 100), slide_width - x))
    height = max(10, min(group_rect.get('height', 20), slide_height - y))
    styles = inline_group.get('styles', {})
    
    # Enhanced red highlight wrapper handling
    has_red_highlight = inline_group.get('hasRedHighlight', False)
    red_highlight_styles = inline_group.get('redHighlightStyles', {})
    
    try:
        # Add red highlight wrapper border if present with accurate positioning
        if has_red_highlight and red_highlight_styles:
            extended_rect = red_highlight_styles.get('extendedRect', {})
            if extended_rect:
                # Use the actual extended rectangle from proto3
                wrapper_x = max(0, min(extended_rect.get('x', x), slide_width - 10))
                wrapper_y = max(0, min(extended_rect.get('y', y), slide_height - 10))
                wrapper_width = max(10, min(extended_rect.get('width', width), slide_width - wrapper_x))
                wrapper_height = max(10, min(extended_rect.get('height', height), slide_height - wrapper_y))
            else:
                # Fallback: calculate extension from CSS values
                top_offset = safe_float(red_highlight_styles.get('top', '-4px'))
                left_offset = safe_float(red_highlight_styles.get('left', '-20px'))
                right_offset = safe_float(red_highlight_styles.get('right', '-20px'))
                bottom_offset = safe_float(red_highlight_styles.get('bottom', '-4px'))
                
                # Apply offsets (negative values extend outward)
                wrapper_x = max(0, x + left_offset)
                wrapper_y = max(0, y + top_offset)
                wrapper_width = min(width - left_offset - right_offset, slide_width - wrapper_x)
                wrapper_height = min(height - top_offset - bottom_offset, slide_height - wrapper_y)
            
            border_color = parse_color(red_highlight_styles.get('borderColor', '#ff0000'))
            border_width = safe_float(red_highlight_styles.get('borderWidth', '2px'))
            border_style = red_highlight_styles.get('borderStyle', 'solid')
            
            if border_color and border_width > 0 and wrapper_width > 0 and wrapper_height > 0:
                wrapper_shape = slide.shapes.add_shape(
                    MSO_SHAPE.RECTANGLE,
                    pixels_to_emu(wrapper_x), pixels_to_emu(wrapper_y),
                    pixels_to_emu(wrapper_width), pixels_to_emu(wrapper_height)
                )
                
                # Make background transparent
                wrapper_shape.fill.background()
                wrapper_shape.shadow.inherit = False
                
                # Add the border with correct style
                wrapper_shape.line.width = Pt(max(0.5, border_width))
                wrapper_shape.line.color.rgb = border_color
                if border_style == 'dashed':
                    wrapper_shape.line.dash_style = MSO_LINE.DASH
                elif border_style == 'dotted':
                    wrapper_shape.line.dash_style = MSO_LINE.ROUND_DOT
                else:
                    wrapper_shape.line.dash_style = MSO_LINE.SOLID
        
        # Only add text content if there are inline elements
        if has_content:
            # Add regular background shapes
            if parse_color(styles.get('backgroundColor')) or has_any_border(styles):
                add_bg_shape(slide, styles, x, y, width, height)
            
            textbox = slide.shapes.add_textbox(
                pixels_to_emu(x), pixels_to_emu(y),
                pixels_to_emu(width), pixels_to_emu(height)
            )
            text_frame = textbox.text_frame
            text_frame.word_wrap = True
            text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
            text_frame.margin_left = pixels_to_emu(safe_float(styles.get('paddingLeft', '0px').replace('px', '')))
            text_frame.margin_right = pixels_to_emu(safe_float(styles.get('paddingRight', '0px').replace('px', '')))
            text_frame.margin_top = pixels_to_emu(safe_float(styles.get('paddingTop', '0px').replace('px', '')))
            text_frame.margin_bottom = pixels_to_emu(safe_float(styles.get('paddingBottom', '0px').replace('px', '')))
            textbox.fill.background()
            textbox.line.fill.background()
            textbox.shadow.inherit = False
            text_frame.clear()
            p = text_frame.paragraphs[0]
            first = True
            text_align = styles.get('textAlign', 'left')
            alignment = PP_ALIGN.CENTER if text_align == 'center' else PP_ALIGN.RIGHT if text_align == 'right' else PP_ALIGN.LEFT
            p.alignment = alignment
            
            for inline_element in inline_elements:
                if inline_element.get('type') == 'br':
                    if p is not None:
                        p = text_frame.add_paragraph()
                        p.alignment = alignment
                    first = True
                    continue
                element_text = inline_element.get('text', '')
                if first:
                    element_text = element_text.lstrip()
                if not element_text.strip():
                    continue
                first = False
                run = p.add_run()
                run.text = element_text
                inline_styles = inline_element.get('styles', {})
                font = run.font
                font_size_px = safe_float(inline_styles.get('fontSize', '16').replace('px', ''))
                font.name = inline_styles.get('fontFamily', 'Segoe UI').split(',')[0].strip('"\'')
                font.size = Pt(get_font_size_pt(font_size_px))
                font.bold = inline_styles.get('fontWeight', '400') in ['bold', '600', '700', '800', '900']
                font.italic = inline_styles.get('fontStyle', 'normal') == 'italic'
                color = parse_color(inline_styles.get('color'))
                if color:
                    font.color.rgb = color
            
            # Trim trailing spaces from the last run in the last paragraph
            if text_frame.paragraphs and text_frame.paragraphs[-1].runs:
                last_run = text_frame.paragraphs[-1].runs[-1]
                last_run.text = last_run.text.rstrip()
                
    except Exception as e:
        print(f"Failed to add inline group element: {e}")
def add_overlay_element(slide, element, slide_width, slide_height):
    """Render overlay/highlight elements as shapes using their CSS (generic, not hardcoded)."""
    x = element.get('x', 0)
    y = element.get('y', 0)
    width = max(1, element.get('width', 10))
    height = max(1, element.get('height', 10))
    styles = element.get('styles', {})
    x = max(0, min(x, slide_width - width))
    y = max(0, min(y, slide_height - height))
    bg_color = parse_color(styles.get('backgroundColor'))
    border_radius_str = styles.get('borderRadius', '0px')
    border_radius = parse_border_radius(border_radius_str, width, height)
    has_radius = border_radius > 0
    has_border = is_uniform_border(styles)
    has_any_border_sides = has_any_border(styles)
    box_shadow = styles.get('boxShadow', 'none')
    has_shadow = box_shadow != 'none'
    if bg_color or has_border or has_any_border_sides or has_radius or has_shadow:
        add_bg_shape(slide, styles, x, y, width, height)
    text = element.get('text', '').strip()
    if text:
        textbox = slide.shapes.add_textbox(
            pixels_to_emu(x), pixels_to_emu(y),
            pixels_to_emu(width), pixels_to_emu(height)
        )
        text_frame = textbox.text_frame
        text_frame.word_wrap = True
        text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        text_frame.clear()
        p = text_frame.paragraphs[0]
        run = p.add_run()
        run.text = text
        font = run.font
        font_size_px = safe_float(styles.get('fontSize', '12'))
        font.name = styles.get('fontFamily', 'Arial').split(',')[0].strip('"\'')
        font.size = Pt(max(6, get_font_size_pt(font_size_px)))
        font.bold = styles.get('fontWeight', '400') in ['bold', '700', '800', '900']
        font.italic = styles.get('fontStyle') == 'italic'
        color = parse_color(styles.get('color', 'black'))
        if color:
            font.color.rgb = color
        textbox.fill.background()
        textbox.line.fill.background()
        textbox.shadow.inherit = False

def get_bullet_style_for_list(element, list_info):
    """
    Determine bullet char, color, font, and size for a list based on className or CSS styles.
    Returns (bullet_char, bullet_color, bullet_font, bullet_size_pt)
    """
    class_name = element.get('className', '')
    styles = element.get('styles', {})
    bullet_char = '\u2022'
    bullet_color = RGBColor(40, 167, 69)
    bullet_font = 'Arial'
    bullet_size_pt = 12

    # Try to infer from class or parent context
    if 'top-points' in class_name:
        bullet_color = RGBColor(40, 167, 69)
        bullet_font = 'Arial'
        bullet_size_pt = 14
        bullet_char = '●'
    elif 'category-points' in class_name:
        bullet_color = RGBColor(40, 167, 69)
        bullet_font = 'Arial'
        bullet_size_pt = 12
        bullet_char = '●'
    # Try to get from styles if present
    color_str = styles.get('color')
    if color_str:
        parsed = parse_color(color_str)
        if parsed:
            bullet_color = parsed
    font_size = safe_float(styles.get('fontSize', '12'))
    if font_size:
        bullet_size_pt = get_font_size_pt(font_size)
    font_family = styles.get('fontFamily')
    if font_family:
        bullet_font = font_family.split(',')[0].strip('"\'')
    # Try to get bullet char from ::before content if available
    if 'listInfo' in element and element['listInfo'].get('bulletChar'):
        bullet_char = element['listInfo']['bulletChar']
    return bullet_char, bullet_color, bullet_font, bullet_size_pt

def add_list_paragraphs(text_frame, list_info, level=0, counters=None, element=None, bullet_style_override=None):
    if counters is None:
        counters = {}
    list_type = list_info.get('type')
    is_ordered = list_type == 'ol'
    list_styles = list_info.get('listStyles', {})
    list_style_type = list_styles.get('listStyleType', 'disc' if not is_ordered else 'decimal')
    if is_ordered:
        counter_key = f'ol_{level}'
        counters[counter_key] = list_info.get('start', 1) - 1
    items = list_info.get('items', [])
    avg_space_px = 0
    if len(items) > 1:
        spaces = []
        for i in range(len(items) - 1):
            item_bottom = items[i]['rect']['y'] + items[i]['rect']['height']
            next_top = items[i+1]['rect']['y']
            space_px = next_top - item_bottom
            if space_px > 0:
                spaces.append(space_px)
        avg_space_px = sum(spaces) / len(spaces) if spaces else 0
    space_after_pt = px_to_pt(avg_space_px)
    line_height_str = list_info.get('styles', {}).get('lineHeight', 'normal')
    if line_height_str == 'normal':
        line_spacing = 1.15
    else:
        try:
            line_spacing = float(line_height_str)
        except ValueError:
            line_spacing = 1.15

    # Determine bullet style for this list
    if bullet_style_override:
        bullet_char, bullet_color, bullet_font, bullet_size_pt = bullet_style_override
    else:
        bullet_char, bullet_color, bullet_font, bullet_size_pt = get_bullet_style_for_list(element or {}, list_info)

    first_item = True
    for item in items:
        p = None
        item_styles = item.get('styles', {})
        default_font_size_px = safe_float(item_styles.get('fontSize', '16').replace('px', ''))
        default_font_size_pt = get_font_size_pt(default_font_size_px)
        default_font_name = item_styles.get('fontFamily', 'Segoe UI').split(',')[0].strip('"\'')
        default_color = parse_color(item_styles.get('color'))

        # Per-item bullet info (from proto3 extraction ::before)
        item_bullet = item.get('bulletInfo') or {}
        per_bullet_char = None
        per_bullet_color = None
        if item_bullet:
            # bulletInfo.content holds the ::before content (if any)
            per_bullet_char = item_bullet.get('content') or item_bullet.get('text') or None
            per_styles = item_bullet.get('styles') or {}
            per_bullet_color = parse_color(per_styles.get('color') or per_styles.get('backgroundColor') or per_styles.get('borderColor') or '')

        if item.get('inlineGroup') and item['inlineGroup'].get('inlineElements'):
            inline_elements = item['inlineGroup']['inlineElements']
            first = True
            bullet_added = False
            for inline_element in inline_elements:
                if inline_element.get('type') == 'br':
                    if p is not None:
                        p = text_frame.add_paragraph()
                        p.level = level
                        p.space_after = Pt(space_after_pt)
                        p.line_spacing = line_spacing
                        indent_pt = 18 * level
                        p.left_indent = Pt(indent_pt)
                        p.first_line_indent = Pt(-18)
                    first = True
                    bullet_added = False
                    continue
                element_text = inline_element.get('text', '')
                if p is None:
                    if first_item and level == 0:
                        p = text_frame.paragraphs[0]
                    else:
                        p = text_frame.add_paragraph()
                    p.level = level
                    p.space_after = Pt(space_after_pt)
                    p.line_spacing = line_spacing
                    indent_pt = 18 * level
                    p.left_indent = Pt(indent_pt)
                    p.first_line_indent = Pt(-18)
                if not bullet_added:
                    if is_ordered:
                        counters[counter_key] += 1
                        if list_style_type == 'decimal':
                            marker_str = f"{counters[counter_key]}."
                        elif list_style_type == 'lower-alpha':
                            marker_str = f"{chr(96 + counters[counter_key])}."
                        elif list_style_type == 'upper-alpha':
                            marker_str = f"{chr(64 + counters[counter_key])}."
                        else:
                            marker_str = f"{counters[counter_key]}."
                        marker_run = p.add_run()
                        marker_run.text = marker_str + ' '
                        marker_run.font.name = default_font_name
                        marker_run.font.size = Pt(default_font_size_pt)
                        if default_color:
                            marker_run.font.color.rgb = default_color
                    else:
                        bullet_run = p.add_run()
                        # Prefer per-item bullet char if recorded, else list-level bullet
                        bullet_text = per_bullet_char if per_bullet_char else bullet_char
                        bullet_run.text = (bullet_text or bullet_char) + ' '
                        bullet_run.font.name = bullet_font
                        bullet_run.font.size = Pt(bullet_size_pt)
                        # Use per-item color if present, else list-level color
                        if per_bullet_color:
                            bullet_run.font.color.rgb = per_bullet_color
                        elif bullet_color:
                            bullet_run.font.color.rgb = bullet_color
                    bullet_added = True
                if first:
                    element_text = element_text.lstrip()
                if not element_text.strip():
                    continue
                first = False
                run = p.add_run()
                run.text = element_text
                inline_styles = inline_element.get('styles', {})
                font = run.font
                font_size_px = safe_float(inline_styles.get('fontSize', '16').replace('px', ''))
                font.name = inline_styles.get('fontFamily', 'Segoe UI').split(',')[0].strip('"\'')
                font.size = Pt(get_font_size_pt(font_size_px))
                font.bold = inline_styles.get('fontWeight', '400') in ['bold', '600', '700', '800', '900']
                font.italic = inline_styles.get('fontStyle', 'normal') == 'italic'
                color = parse_color(inline_styles.get('color'))
                if color:
                    font.color.rgb = color
            if text_frame.paragraphs and text_frame.paragraphs[-1].runs:
                last_run = text_frame.paragraphs[-1].runs[-1]
                last_run.text = last_run.text.rstrip()
        else:
            if first_item and level == 0:
                p = text_frame.paragraphs[0]
            else:
                p = text_frame.add_paragraph()
            p.level = level
            p.space_after = Pt(space_after_pt)
            p.line_spacing = line_spacing
            indent_pt = 18 * level
            p.left_indent = Pt(indent_pt)
            p.first_line_indent = Pt(-18)
            if is_ordered:
                counters[counter_key] += 1
                if list_style_type == 'decimal':
                    marker_str = f"{counters[counter_key]}."
                elif list_style_type == 'lower-alpha':
                    marker_str = f"{chr(96 + counters[counter_key])}."
                elif list_style_type == 'upper-alpha':
                    marker_str = f"{chr(64 + counters[counter_key])}."
                else:
                    marker_str = f"{counters[counter_key]}."
                marker_run = p.add_run()
                marker_run.text = marker_str + ' '
                marker_run.font.name = default_font_name
                marker_run.font.size = Pt(default_font_size_pt)
                if default_color:
                    marker_run.font.color.rgb = default_color
            else:
                bullet_run = p.add_run()
                bullet_text = per_bullet_char if per_bullet_char else bullet_char
                bullet_run.text = (bullet_text or bullet_char) + ' '
                bullet_run.font.name = bullet_font
                bullet_run.font.size = Pt(bullet_size_pt)
                if per_bullet_color:
                    bullet_run.font.color.rgb = per_bullet_color
                elif bullet_color:
                    bullet_run.font.color.rgb = bullet_color
            run = p.add_run()
            run.text = item.get('text', '').strip()
            font = run.font
            font.name = default_font_name
            font.size = Pt(default_font_size_pt)
            font.bold = item_styles.get('fontWeight', '400') in ['bold', '600', '700', '800', '900']
            font.italic = item_styles.get('fontStyle', 'normal') == 'italic'
            color = parse_color(item_styles.get('color'))
            if color:
                font.color.rgb = color

        first_item = False
        if item.get('nestedList'):
            add_list_paragraphs(text_frame, item['nestedList'], level + 1, counters, element, bullet_style_override)

def add_list_element(slide, element, slide_width, slide_height, parent_has_shadow=False):
    list_info = element.get('listInfo', {})
    if not list_info.get('items'):
        return
    rect = list_info.get('rect', {})
    x = safe_int(rect.get('x', 0))
    y = safe_int(rect.get('y', 0))
    width = safe_int(rect.get('width', element.get('width', 100)))
    height = safe_int(rect.get('height', element.get('height', 100)))
    styles = element.get('styles', {})
    box_shadow = styles.get('boxShadow', 'none')
    has_shadow = box_shadow != 'none' and not parent_has_shadow
    bg_color = parse_color(styles.get('backgroundColor'))
    border_radius_str = styles.get('borderRadius', '0px')
    border_radius = parse_border_radius(border_radius_str, width, height)
    has_radius = border_radius > 0
    has_border = is_uniform_border(styles)
    has_any_border_sides = has_any_border(styles)
    try:
        if bg_color or has_border or has_any_border_sides or has_radius or has_shadow:
            add_bg_shape(slide, styles, x, y, width, height)
        textbox = slide.shapes.add_textbox(
            pixels_to_emu(x), pixels_to_emu(y),
            pixels_to_emu(width), pixels_to_emu(height)
        )
        text_frame = textbox.text_frame
        text_frame.word_wrap = True
        text_frame.vertical_anchor = MSO_ANCHOR.TOP
        text_frame.margin_left = pixels_to_emu(safe_float(styles.get('paddingLeft', '0px').replace('px', '')))
        text_frame.margin_right = pixels_to_emu(safe_float(styles.get('paddingRight', '0px').replace('px', '')))
        text_frame.margin_top = pixels_to_emu(safe_float(styles.get('paddingTop', '0px').replace('px', '')))
        text_frame.margin_bottom = pixels_to_emu(safe_float(styles.get('paddingBottom', '0px').replace('px', '')))
        textbox.fill.background()
        textbox.line.fill.background()
        textbox.shadow.inherit = False
        text_frame.clear()
        bullet_style = get_bullet_style_for_list(element, list_info)
        add_list_paragraphs(text_frame, list_info, element=element, bullet_style_override=bullet_style)
    except Exception as e:
        print(f"Failed to add list: {e}")
def set_cell_border(cell, side, width_px, color_rgb, style='solid'):
    """Apply a border to a table cell side using enhanced method"""
    set_cell_border_enhanced(cell, side, width_px, color_rgb, style)
def add_table_element(slide, element, slide_width, slide_height, parent_has_shadow=False):
    table_info = element.get('tableInfo', {})
    if not table_info.get('rows'):
        return
    rect = table_info.get('rect', {})
    x = safe_float(rect.get('x', element.get('x', 0)))
    y = safe_float(rect.get('y', element.get('y', 0)))
    width = safe_float(rect.get('width', element.get('width', 100)))
    height = safe_float(rect.get('height', element.get('height', 100)))
    element_class = element.get('className', '')
    styles = table_info.get('styles', {})
    box_shadow = styles.get('boxShadow', 'none')
    has_shadow = box_shadow != 'none' and not parent_has_shadow
    bg_color = parse_color(styles.get('backgroundColor'))
    border_radius_str = styles.get('borderRadius', '0px')
    border_radius = parse_border_radius(border_radius_str, width, height)
    has_radius = border_radius > 0
    has_border = is_uniform_border(styles)
    has_any_border_sides = has_any_border(styles)
    try:
        # Clamp to slide bounds (keep existing safeguard)
        if x + width > slide_width:
            width = max(10, slide_width - x)
        if y + height > slide_height:
            height = max(10, slide_height - y)

        if bg_color or has_border or has_radius or has_shadow:
            add_bg_shape(slide, styles, x, y, width, height)

        rows = table_info['rowCount']
        cols = table_info['columnCount']
        table_shape = slide.shapes.add_table(
            rows, cols,
            pixels_to_emu(x), pixels_to_emu(y),
            pixels_to_emu(width), pixels_to_emu(height)
        )
        table = table_shape.table

        # Set table style to ensure borders render and disable default formatting
        try:
            tbl = table._graphic_frame._graphicFrame.graphicData.tbl
            style_id = '{2D5ABB26-0587-4C30-8999-92F81FD0307C}'
            tbl[0][-1].text = style_id
            table.first_row = False  # Disable header row default styling
        except Exception:
            pass

        # --- Column width calculation ---
        if table_info['rows']:
            first_row = next((r for r in table_info['rows'] if r.get('cells')), None)
            if first_row:
                current_col = 0
                row_cells = first_row['cells']
                total_assigned = 0.0
                for cell_data in row_cells:
                    cell_rect = cell_data.get('rect', {})
                    cell_w = cell_rect.get('width', 0)
                    col_span = max(1, cell_data.get('colSpan', 1))
                    per_col_w = cell_w / col_span if col_span > 1 else cell_w
                    for i in range(col_span):
                        if current_col + i < cols:
                            table.columns[current_col + i].width = pixels_to_emu(max(8, per_col_w))
                            total_assigned += per_col_w
                    current_col += col_span
                # If any remaining columns distribute remaining space
                remaining_cols = cols - current_col
                if remaining_cols > 0:
                    remaining_w = max(0, width - total_assigned)
                    per_remaining = remaining_w / remaining_cols if remaining_cols else 0
                    for i in range(remaining_cols):
                        idx = current_col + i
                        if idx < cols:
                            table.columns[idx].width = pixels_to_emu(max(8, per_remaining))

        # --- Enhanced row heights with consistent sizing ---
        for row_data in table_info['rows']:
            r_index = row_data['index']
            if r_index >= rows:
                continue
            
            # Get CSS height from row data
            row_styles = row_data.get('styles', {})
            css_height = row_styles.get('height', '')
            
            # Parse CSS height (e.g., "25px" -> 25)
            if css_height and 'px' in css_height:
                try:
                    specified_height = float(css_height.replace('px', ''))
                except:
                    specified_height = None
            else:
                specified_height = None
            
            # Get rendered height from rect
            rendered_height = row_data.get('rect', {}).get('height', 25)
            
            # Choose the most appropriate height
            if specified_height:
                # Use CSS specified height as primary
                final_height = specified_height
            else:
                # Fall back to rendered height
                final_height = rendered_height
            
            # Ensure minimum reasonable height and clamp maximum
            final_height = max(20, min(final_height, 120))
            
            table.rows[r_index].height = pixels_to_emu(final_height)

        # Apply default borders to ALL table cells
        default_border_color = "808080"  # Gray color
        default_border_width = '9525'    # 0.75pt in EMU
        
        # Apply borders to every single cell
        for r in range(rows):
            for c in range(cols):
                try:
                    pptx_cell = table.cell(r, c)
                    _set_cell_border(pptx_cell, border_color=default_border_color, border_width=default_border_width)
                except Exception as e:
                    print(f"Error applying border to cell ({r}, {c}): {e}")

        # Populate & style cells with enhanced empty cell handling
        for row_data in table_info['rows']:
            row_index = row_data['index']
            if row_index >= rows:
                continue
            row_bg_color = parse_color(row_data.get('styles', {}).get('backgroundColor'))
            for cell_data in row_data.get('cells', []):
                cell_index = cell_data.get('cellIndex')
                if cell_index is None or cell_index >= cols:
                    continue
                pptx_cell = table.cell(row_index, cell_index)

                # Merge spans
                col_span = max(1, cell_data.get('colSpan', 1))
                row_span = max(1, cell_data.get('rowSpan', 1))
                if col_span > 1 or row_span > 1:
                    try:
                        end_row = min(rows - 1, row_index + row_span - 1)
                        end_col = min(cols - 1, cell_index + col_span - 1)
                        pptx_cell = pptx_cell.merge(table.cell(end_row, end_col))
                        # Reapply borders after merge
                        _set_cell_border(pptx_cell, border_color=default_border_color, border_width=default_border_width)
                    except Exception as e:
                        print(f"Error merging cell ({row_index}, {cell_index}): {e}")

                text_frame = pptx_cell.text_frame
                text_frame.word_wrap = True
                cell_styles = cell_data.get('styles', {})

                vertical_align = (cell_styles.get('verticalAlign') or '').strip().lower()
                if vertical_align == 'top':
                    pptx_cell.vertical_anchor = MSO_ANCHOR.TOP
                elif vertical_align == 'bottom':
                    pptx_cell.vertical_anchor = MSO_ANCHOR.BOTTOM
                else:
                    pptx_cell.vertical_anchor = MSO_ANCHOR.MIDDLE  # default

                # Enhanced padding with consistent values
                padding_left = safe_float(cell_styles.get('paddingLeft', '6px'))
                padding_right = safe_float(cell_styles.get('paddingRight', '6px'))
                padding_top = safe_float(cell_styles.get('paddingTop', '4px'))
                padding_bottom = safe_float(cell_styles.get('paddingBottom', '4px'))
                text_frame.margin_left = pixels_to_emu(padding_left)
                text_frame.margin_right = pixels_to_emu(padding_right)
                text_frame.margin_top = pixels_to_emu(padding_top)
                text_frame.margin_bottom = pixels_to_emu(padding_bottom)

                # Background
                bg_color_cell = parse_color(cell_styles.get('backgroundColor'))
                if bg_color_cell:
                    pptx_cell.fill.solid()
                    pptx_cell.fill.fore_color.rgb = bg_color_cell
                elif row_bg_color:
                    pptx_cell.fill.solid()
                    pptx_cell.fill.fore_color.rgb = row_bg_color
                else:
                    pptx_cell.fill.background()

                # Enhanced text content handling for empty cells
                text_frame.clear()
                p = text_frame.paragraphs[0]
                text_align = cell_styles.get('textAlign', 'left')
                p.alignment = PP_ALIGN.CENTER if text_align == 'center' else PP_ALIGN.RIGHT if text_align == 'right' else PP_ALIGN.LEFT

                # Get cell text content
                cell_text = ""
                if cell_data.get('inlineGroup') and cell_data['inlineGroup'].get('inlineElements'):
                    # Handle inline formatted content
                    first = True
                    for inline_element in cell_data['inlineGroup']['inlineElements']:
                        if inline_element.get('type') == 'br':
                            p = text_frame.add_paragraph()
                            p.alignment = PP_ALIGN.CENTER if text_align == 'center' else PP_ALIGN.RIGHT if text_align == 'right' else PP_ALIGN.LEFT
                            first = True
                            continue
                        t = inline_element.get('text', '')
                        if first:
                            t = t.lstrip()
                        if not t.strip():
                            continue
                        first = False
                        run = p.add_run()
                        run.text = t
                        cell_text += t
                        inline_styles = inline_element.get('styles', {})
                        font = run.font
                        fs_px = safe_float(inline_styles.get('fontSize', '16'))
                        font.size = Pt(max(6, get_font_size_pt(fs_px)))
                        font.name = inline_styles.get('fontFamily', 'Arial').split(',')[0].strip('"\'')
                        font.bold = inline_styles.get('fontWeight', '400') in ['bold', '600', '700', '800', '900']
                        font.italic = inline_styles.get('fontStyle', 'normal') == 'italic'
                        clr = parse_color(inline_styles.get('color') or cell_styles.get('color'))
                        if clr:
                            font.color.rgb = clr
                    # Trim trailing spaces
                    if text_frame.paragraphs and text_frame.paragraphs[-1].runs:
                        last_run = text_frame.paragraphs[-1].runs[-1]
                        last_run.text = last_run.text.rstrip()
                else:
                    # Handle plain text content
                    cell_text = (cell_data.get('text') or '').strip()
                
                # Enhanced empty cell handling
                if not cell_text:
                    # For empty cells, add a non-breaking space to maintain consistent height
                    run = p.add_run()
                    run.text = "\u00A0"  # Non-breaking space
                    font = run.font
                    fs_px = safe_float(cell_styles.get('fontSize', '10'))
                    font.size = Pt(max(6, get_font_size_pt(fs_px)))
                    font.name = cell_styles.get('fontFamily', 'Arial').split(',')[0].strip('"\'')
                    # Make the non-breaking space transparent or same color as background
                    clr = parse_color(cell_styles.get('color', '#000000'))
                    if clr:
                        font.color.rgb = clr
                else:
                    # Add the actual content
                    run = p.add_run()
                    run.text = cell_text
                    font = run.font
                    fs_px = safe_float(cell_styles.get('fontSize', '10'))
                    font.size = Pt(max(6, get_font_size_pt(fs_px)))
                    font.name = cell_styles.get('fontFamily', 'Arial').split(',')[0].strip('"\'')
                    font.bold = cell_styles.get('fontWeight', '400') in ['bold', '600', '700', '800', '900']
                    font.italic = cell_styles.get('fontStyle', 'normal') == 'italic'
                    clr = parse_color(cell_styles.get('color'))
                    if clr:
                        font.color.rgb = clr

                # Ensure consistent paragraph spacing
                p.space_before = Pt(0)
                p.space_after = Pt(0)
                p.line_spacing = 1.0
    except Exception as e:
        print(f"Failed to add table: {e}")
def add_image_element(slide, element, slide_width, slide_height, parent_has_shadow=False):
    media_info = element.get('mediaInfo', {})
    img_src = media_info.get('src', '')
    styles = element.get('styles', {})
    if not img_src:
        return
    # Use precise positioning from extraction
    x = element.get('x', 0)
    y = element.get('y', 0)
    width = max(1, element.get('width', 100))
    height = max(1, element.get('height', 100))
   
    # Ensure coordinates are within slide bounds
    x = max(0, min(x, slide_width - width))
    y = max(0, min(y, slide_height - height))
   
    natural_width = media_info.get('naturalWidth', width)
    border_radius_str = styles.get('borderRadius', '0px')
    radius_ratio = parse_border_radius(border_radius_str, width, height)
    radius_display = radius_ratio * min(width, height)
    has_radius = radius_display > 0
    box_shadow = styles.get('boxShadow', 'none')
    has_shadow = box_shadow != 'none' and not parent_has_shadow
    has_border = is_uniform_border(styles)
   
    try:
        temp_path = None
        if img_src.startswith('data:'):
            _, data = img_src.split(',', 1)
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
                print(f"Failed to download image: {img_src}")
                return
        else:
            if os.path.exists(img_src):
                temp_path = img_src
            else:
                print(f"Image file not found: {img_src}")
                return
       
        with Image.open(temp_path) as img:
            img.verify()
       
        image_to_add = temp_path
        if has_radius:
            scale_x = natural_width / width if width > 0 else 1
            radius_natural = int(radius_display * scale_x)
            temp_rounded = 'temp_rounded.png'
            make_rounded_image(temp_path, temp_rounded, radius_natural)
            image_to_add = temp_rounded
       
        # Add image with precise positioning
        picture = slide.shapes.add_picture(
            image_to_add,
            pixels_to_emu(x), pixels_to_emu(y),
            pixels_to_emu(width), pixels_to_emu(height)
        )
        picture.shadow.inherit = False
       
        # Handle borders and shadows
        if has_border or has_radius or has_shadow:
            shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if has_radius else MSO_SHAPE.RECTANGLE
            border_shape = slide.shapes.add_shape(
                shape_type,
                pixels_to_emu(x), pixels_to_emu(y),
                pixels_to_emu(width), pixels_to_emu(height)
            )
            if has_radius:
                border_shape.adjustments[0] = radius_ratio
            border_shape.fill.background()
            border_shape.shadow.inherit = False
            if has_shadow:
                apply_shadow(border_shape, box_shadow)
            if has_border:
                border_width = safe_float(styles.get('borderTopWidth', '0px'))
                border_color = parse_color(styles.get('borderTopColor'))
                if border_color:
                    border_shape.line.width = Pt(border_width)
                    border_shape.line.color.rgb = border_color
                    border_style = styles.get('borderTopStyle', 'solid')
                    if border_style == 'dashed':
                        border_shape.line.dash_style = MSO_LINE.DASH
                    elif border_style == 'dotted':
                        border_shape.line.dash_style = MSO_LINE.ROUND_DOT
            else:
                border_shape.line.fill.background()
            # Move border shape behind picture
            sp = border_shape._sp
            parent = sp.getparent()
            parent.remove(sp)
            pic_sp = picture._sp
            idx = list(parent).index(pic_sp)
            parent.insert(idx, sp)
        elif has_shadow:
            apply_shadow(picture, box_shadow)
       
        # Clean up temporary files
        if image_to_add != temp_path and os.path.exists(image_to_add):
            os.remove(image_to_add)
        if temp_path != img_src and os.path.exists(temp_path):
            os.remove(temp_path)
    except Exception as e:
        print(f"Failed to add image: {e}")
def add_text_element(slide, element, slide_width, slide_height, parent_has_shadow=False):
    """Enhanced text element creation with precise positioning"""
    text = element.get('text', '').strip()
    if not text:
        return
   
    # Use precise positioning
    x = element.get('x', 0)
    y = element.get('y', 0)
    width = max(1, element.get('width', 100))
    height = max(1, element.get('height', 100))
   
    # Ensure coordinates are within slide bounds
    x = max(0, min(x, slide_width - width))
    y = max(0, min(y, slide_height - height))
   
    styles = element.get('styles', {})
    element_class = element.get('className', '')

    if ('table-heading' in element_class or 'heading' in element_class or 
        'pie-heading' in element_class):
        padding_left = safe_float(styles.get('paddingLeft', '0px'))
        
        # For headings with ::before pseudo elements, adjust position
        if padding_left > 0:
            x += max(0, padding_left - 8)  # Smaller offset to prevent too much gap
            width = max(1, width - padding_left + 8)  # Adjust width accordingly
    
    box_shadow = styles.get('boxShadow', 'none')
    has_shadow = box_shadow != 'none' and not parent_has_shadow
    bg_color = parse_color(styles.get('backgroundColor'))
    border_radius_str = styles.get('borderRadius', '0px')
    border_radius = parse_border_radius(border_radius_str, width, height)
    has_radius = border_radius > 0
    has_border = is_uniform_border(styles)
    has_any_border_sides = has_any_border(styles)
   
    try:
        # Add background/border shapes first
        if bg_color or has_border or has_any_border_sides or has_radius or has_shadow:
            add_bg_shape(slide, styles, x, y, width, height)
       
        # Create text box with precise positioning
        textbox = slide.shapes.add_textbox(
            pixels_to_emu(x), pixels_to_emu(y),
            pixels_to_emu(width), pixels_to_emu(height)
        )
       
        text_frame = textbox.text_frame
        text_frame.word_wrap = True
       
        # Set vertical alignment based on flex properties or default to middle for short text
        display = styles.get('display', 'block')
        align_items = styles.get('alignItems', 'stretch')
        justify_content = styles.get('justifyContent', 'flex-start')
       
        if display == 'flex' and align_items == 'center':
            text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        elif height < 50: # For small elements like company names
            text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        else:
            text_frame.vertical_anchor = MSO_ANCHOR.TOP
       
        # Enhanced margin calculation for headings with pseudo elements
        margin_left = safe_float(styles.get('paddingLeft', '0px'))
        margin_right = safe_float(styles.get('paddingRight', '0px'))
        margin_top = safe_float(styles.get('paddingTop', '0px'))
        margin_bottom = safe_float(styles.get('paddingBottom', '0px'))
        
        # Special adjustment for headings with pseudo elements
        if ('table-heading' in element_class or 'pie-heading' in element_class):
            # Reduce left margin to account for positioning adjustment
            margin_left = max(0, margin_left - 12)
        
        text_frame.margin_left = pixels_to_emu(margin_left)
        text_frame.margin_right = pixels_to_emu(margin_right)
        text_frame.margin_top = pixels_to_emu(margin_top)
        text_frame.margin_bottom = pixels_to_emu(margin_bottom)
       
        # Remove textbox styling
        textbox.fill.background()
        textbox.line.fill.background()
        textbox.shadow.inherit = False
       
        # Clear and set text
        text_frame.clear()
        p = text_frame.paragraphs[0]
        run = p.add_run()
        run.text = text
       
        # Apply text formatting
        font = run.font
        font_size_px = safe_float(styles.get('fontSize', '12'))
        font.name = styles.get('fontFamily', 'Arial').split(',')[0].strip('"\'')
        font.size = Pt(max(6, get_font_size_pt(font_size_px)))
        font.bold = styles.get('fontWeight', '400') in ['bold', '700', '800', '900']
        font.italic = styles.get('fontStyle') == 'italic'
       
        # Apply text color
        color = parse_color(styles.get('color', 'black'))
        if color:
            font.color.rgb = color
       
        # Apply text alignment based on CSS text-align and flex properties
        text_align = styles.get('textAlign', 'left')
        if text_align == 'center' or (display == 'flex' and justify_content == 'center'):
            p.alignment = PP_ALIGN.CENTER
        elif text_align == 'right' or (display == 'flex' and justify_content == 'flex-end'):
            p.alignment = PP_ALIGN.RIGHT
        else:
            p.alignment = PP_ALIGN.LEFT
           
    except Exception as e:
        print(f"Failed to add text: {e}")
def parse_linear_gradient(gradient_str, total_width):
    """Parse linear-gradient color stops (hex, rgb, rgba) into solid segments.

    Supports patterns like:
      linear-gradient(to right, rgb(37,29,95) 0%, rgb(37,29,95) 12%, rgb(217,217,217) 12%, rgb(217,217,217) 100%)
      linear-gradient(to right, #251d5f 0%, #251d5f 12%, #D9D9D9 12%, #D9D9D9 100%)
    Returns list of (x, width, RGBColor)
    """
    if not isinstance(gradient_str, str) or 'linear-gradient' not in gradient_str:
        return []

    try:
        # Match both hex and rgb/rgba colors with percentage stops
        pattern = r'((?:#[0-9a-fA-F]{3,6})|(?:rgb[a]?\(\s*\d+\s*,\s*\d+\s*,\s*\d+(?:\s*,\s*[\d.]+\s*)?\)))\s+([\d.]+)%'
        stops = re.findall(pattern, gradient_str)
        if not stops or len(stops) < 2:
            return []

        # Build segments from consecutive stop pairs
        segments = []
        for i in range(len(stops) - 1):
            color_text, start_pct = stops[i]
            _, end_pct = stops[i + 1]
            start_pct = float(start_pct)
            end_pct = float(end_pct)
            if end_pct <= start_pct:
                continue
            seg_x = (start_pct / 100.0) * total_width
            seg_w = ((end_pct - start_pct) / 100.0) * total_width
            color = parse_color(color_text)
            if color and seg_w > 0:
                segments.append((seg_x, seg_w, color))

        # If last stop color not represented (should not happen with pair iteration), ignore.

        return segments
    except Exception as e:
        print(f"Gradient parse failed: {e}")
        return []
def add_pseudo_element(slide, element, slide_width, slide_height):
    x = safe_float(element.get('x', 0))
    y = safe_float(element.get('y', 0))
    width = max(1, safe_float(element.get('width', 0)))
    height = max(1, safe_float(element.get('height', 0)))
    
    styles = element.get('styles', {})
    text = element.get('text', '')
    pseudo_type = element.get('pseudoType', '')
    parent_class = element.get('parentClassName', '')
    
    x = max(0, min(x, slide_width - width))
    y = max(0, min(y, slide_height - height))
    
    background_full = styles.get('background', '')
    has_gradient = isinstance(background_full, str) and 'linear-gradient' in background_full

    if has_gradient:
        segments = parse_linear_gradient(background_full, width)
        if segments:
            for seg_x, seg_w, color in segments:
                if seg_w > 0 and color:
                    try:
                        shape = slide.shapes.add_shape(
                            MSO_SHAPE.RECTANGLE,
                            pixels_to_emu(x + seg_x),
                            pixels_to_emu(y),
                            pixels_to_emu(seg_w),
                            pixels_to_emu(height)
                        )
                        shape.fill.solid()
                        shape.fill.fore_color.rgb = color
                        shape.line.fill.background()
                        shape.shadow.inherit = False
                    except Exception as e:
                        print(f"Error creating gradient segment: {e}")
        else:
            # Fallback to solid color
            add_bg_shape(slide, styles, x, y, width, height)
    else:
        # Standard pseudo element rendering - ensure proper background color handling
        bg_color = parse_color(styles.get('backgroundColor'))
        if bg_color or has_gradient:
            add_bg_shape(slide, styles, x, y, width, height)

    # Add text content for pseudo elements (usually none for decorative elements)
    if text and text not in ['""', "''", 'none']:
        try:
            textbox = slide.shapes.add_textbox(
                pixels_to_emu(x), pixels_to_emu(y),
                pixels_to_emu(width), pixels_to_emu(height)
            )
            text_frame = textbox.text_frame
            text_frame.word_wrap = True
            text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
            
            # Precise margin handling
            text_frame.margin_left = 0
            text_frame.margin_right = 0
            text_frame.margin_top = 0
            text_frame.margin_bottom = 0
            
            text_frame.clear()
            p = text_frame.paragraphs[0]
            run = p.add_run()
            run.text = text
            
            font = run.font
            font_size_px = safe_float(styles.get('fontSize', '12'))
            font.name = styles.get('fontFamily', 'Meiryo').split(',')[0].strip('"\'')
            font.size = Pt(max(6, get_font_size_pt(font_size_px)))
            font.bold = styles.get('fontWeight', '400') in ['bold', '700', '800', '900']
            font.italic = styles.get('fontStyle') == 'italic'
            
            color = parse_color(styles.get('color', 'black'))
            if color:
                font.color.rgb = color
            
            text_align = styles.get('textAlign', 'left')
            p.alignment = PP_ALIGN.CENTER if text_align == 'center' else PP_ALIGN.RIGHT if text_align == 'right' else PP_ALIGN.LEFT
            
            # Make textbox transparent
            textbox.fill.background()
            textbox.line.fill.background()
            textbox.shadow.inherit = False
        except Exception as e:
            print(f"Failed to add pseudo element text: {e}")
def parse_datalabel_formatter(formatter_str, labels, values, index):
    """
    Parses a Chart.js datalabel formatter string and returns the formatted label.
    Handles different formatter structures for labels, values, or both.
    """
    if not formatter_str or "function" not in formatter_str:
        return str(values[index]) if index < len(values) else ""

    label = labels[index] if index < len(labels) else ""
    value = values[index] if index < len(values) else ""

    # Extract the return statement from the function string
    return_statement_match = re.search(r'return\s+(.*?);', formatter_str)
    if not return_statement_match:
        return str(value)

    return_expression = return_statement_match.group(1).strip()
    
    # Build the formatted string based on the expression
    parts = [part.strip() for part in return_expression.split('+')]
    result_parts = []
    for part in parts:
        part = part.strip('"\'') # Remove string quotes
        if 'context.chart.data.labels[context.dataIndex]' in part:
            result_parts.append(str(label))
        elif 'value' in part:
            result_parts.append(str(value))
        elif part == '\\n': # Handle newlines
            result_parts.append('\n')
        else: # Handle static strings like '%'
            result_parts.append(part)
            
    return "".join(result_parts)

def apply_datalabel_positioning(data_labels, anchor, align, offset=0, chart_type_str="pie"):
    """Apply Chart.js style positioning to PowerPoint data labels"""
    try:
        if chart_type_str == "pie":
            if anchor == "end" and align == "start":
                # Outside the slice, at the edge
                data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
            elif anchor == "center" or (not anchor and not align):
                # Center of slice (default)
                data_labels.position = XL_LABEL_POSITION.CENTER
            elif anchor == "start":
                # Inside, near center
                data_labels.position = XL_LABEL_POSITION.INSIDE_END
            else:
                # Default to outside for better readability
                data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
        elif chart_type_str == "bar":
            if align == "top" or anchor == "end":
                data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
            elif align == "center" or anchor == "center":
                data_labels.position = XL_LABEL_POSITION.CENTER
            else:
                data_labels.position = XL_LABEL_POSITION.INSIDE_END
        else:
            if anchor == "end" or align == "top":
                data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
            elif anchor == "center" or align == "center":
                data_labels.position = XL_LABEL_POSITION.CENTER
            else:
                data_labels.position = XL_LABEL_POSITION.INSIDE_END
    except Exception as e:
        print(f"Error applying data label positioning: {e}")

def add_chart_element(slide, element, slide_width, slide_height):
    """Enhanced chart rendering with accurate configuration extraction"""
    chart_config = None
    
    if element.get('chartInfo'):
        chart_info = element['chartInfo']
        chart_data = chart_info.get('chartData')
        
        if chart_data:
            if isinstance(chart_data, dict):
                # Data is already a parsed dictionary
                chart_config = chart_data
            elif isinstance(chart_data, str):
                # Data is a string, needs parsing
                try:
                    chart_config = json.loads(chart_data)
                except json.JSONDecodeError as e:
                    print(f"Error parsing chart data string: {e}")
                    return
            else:
                # Unsupported format
                return
    
    # Fallback for older format
    if not chart_config:
        chart_config = element.get('chartConfig')

    if not chart_config:
        return
    
    # Use element dimensions directly
    x = max(0, min(element.get('x', 0), slide_width - 10))
    y = max(0, min(element.get('y', 0), slide_height - 10))
    width = max(10, min(element.get('width', 400), slide_width - x))
    height = max(10, min(element.get('height', 300), slide_height - y))
    
    # Apply scaling factor if canvas dimensions differ from CSS dimensions
    if element.get('chartInfo'):
        chart_info = element['chartInfo']
        canvas_width = chart_info.get('width', width)
        canvas_height = chart_info.get('height', height)
        
        # Use canvas dimensions if they're more accurate
        if canvas_width > 0 and canvas_height > 0:
            width = min(canvas_width, slide_width - x)
            height = min(canvas_height, slide_height - y)
    
    try:
        chart_type_str = chart_config.get('type')
        options = chart_config.get('options', {})
        
        # Map chart types with better handling
        if chart_type_str == 'pie':
            chart_type = XL_CHART_TYPE.PIE
        elif chart_type_str == 'line':
            datasets = chart_config.get('data', {}).get('datasets', [])
            has_fill = any(dataset.get('fill') is not None for dataset in datasets)
            if has_fill:
                chart_type = XL_CHART_TYPE.AREA_STACKED if any(dataset.get('fill') == '-1' for dataset in datasets) else XL_CHART_TYPE.AREA
            else:
                point_radius = datasets[0].get('pointRadius', 0) if datasets else 0
                chart_type = XL_CHART_TYPE.LINE_MARKERS if point_radius > 0 else XL_CHART_TYPE.LINE
        elif chart_type_str == 'bar':
            index_axis = options.get('indexAxis', 'x')
            chart_type = XL_CHART_TYPE.BAR_CLUSTERED if index_axis == 'y' else XL_CHART_TYPE.COLUMN_CLUSTERED
        else:
            return
        
        chart_data = CategoryChartData()
        data = chart_config.get('data', {})
        labels = data.get('labels', [])
        
        # Enhanced label handling for multi-line labels
        processed_labels = []
        for label in labels:
            if isinstance(label, str):
                # Handle \n in labels by replacing with actual line breaks
                processed_label = label.replace('\\n', '\n')
                processed_labels.append(processed_label)
            else:
                processed_labels.append(str(label))
        
        chart_data.categories = processed_labels
        
        for dataset in data.get('datasets', []):
            chart_data.add_series(dataset.get('label', ''), tuple(dataset.get('data', [])))
        
        graphic_frame = slide.shapes.add_chart(
            chart_type,
            pixels_to_emu(x),
            pixels_to_emu(y),
            pixels_to_emu(width),
            pixels_to_emu(height),
            chart_data
        )
        chart = graphic_frame.chart

        # Legend handling
        plugins = options.get('plugins', {})
        legend_opts = plugins.get('legend', {})
        if legend_opts.get('display') is False:
            chart.has_legend = False
        else:
            chart.has_legend = True
            position = legend_opts.get('position', 'top')
            if position == 'bottom':
                chart.legend.position = XL_LEGEND_POSITION.BOTTOM
            elif position == 'left':
                chart.legend.position = XL_LEGEND_POSITION.LEFT
            elif position == 'right':
                chart.legend.position = XL_LEGEND_POSITION.RIGHT
            else: # top
                chart.legend.position = XL_LEGEND_POSITION.TOP
            chart.legend.include_in_layout = False

        if chart_type_str != 'pie':
            scales = options.get('scales', {})
            is_horizontal_bar = chart_type == XL_CHART_TYPE.BAR_CLUSTERED
            
            # Get axis configurations
            category_scale = scales.get('y' if is_horizontal_bar else 'x', {})
            value_scale = scales.get('x' if is_horizontal_bar else 'y', {})
            category_axis = chart.value_axis if is_horizontal_bar else chart.category_axis
            value_axis = chart.category_axis if is_horizontal_bar else chart.value_axis

            try:
                # Handle beginAtZero
                if value_scale.get('beginAtZero', True):
                    value_axis.minimum_scale = 0.0
                
                # Handle min/max values
                val_min = value_scale.get('min')
                val_max = value_scale.get('max')
                suggested_max = value_scale.get('suggestedMax')
                
                if val_min is not None:
                    value_axis.minimum_scale = float(val_min)
                
                if val_max is not None:
                    value_axis.maximum_scale = float(val_max)
                elif suggested_max is not None:
                    value_axis.maximum_scale = float(suggested_max)
                
                # Handle step size
                ticks = value_scale.get('ticks', {})
                step_size = ticks.get('stepSize')
                if step_size:
                    value_axis.major_unit = float(step_size)

                if value_axis.tick_labels:
                    tick_font = ticks.get('font', {})
                    if tick_font.get('size'):
                        value_axis.tick_labels.font.size = Pt(tick_font['size'])
                    if tick_font.get('family'):
                        value_axis.tick_labels.font.name = tick_font['family']
                    
                    tick_color = parse_color(ticks.get('color', '#888888'))
                    if tick_color:
                        value_axis.tick_labels.font.color.rgb = tick_color
                
                # Handle axis visibility
                if value_scale.get('display', True) == False:
                    value_axis.visible = False
                
                # Hide axis line if border is not displayed
                if value_scale.get('border', {}).get('display') is False:
                    value_axis.format.line.fill.background()

            except Exception as e:
                print(f"Error configuring value axis: {e}")
            
            # Category axis configuration
            try:
                if category_scale.get('display', True) == False:
                    category_axis.visible = False
                
                # Hide axis line if border is not displayed
                if category_scale.get('border', {}).get('display') is False:
                    category_axis.format.line.fill.background()

                cat_ticks = category_scale.get('ticks', {})
                if category_axis.tick_labels:
                    tick_font = cat_ticks.get('font', {})
                    if tick_font.get('size'):
                        category_axis.tick_labels.font.size = Pt(tick_font['size'])
                    if tick_font.get('family'):
                        category_axis.tick_labels.font.name = tick_font['family']
                    
                    tick_color = parse_color(cat_ticks.get('color', '#666666'))
                    if tick_color:
                        category_axis.tick_labels.font.color.rgb = tick_color
                    
                    # Handle label rotation
                    max_rotation = cat_ticks.get('maxRotation', 0)
                    if max_rotation == 0:
                        category_axis.tick_labels.orientation = 0  # Horizontal
                        
            except Exception as e:
                print(f"Error configuring category axis: {e}")

        # Data Labels
        datalabel_opts = plugins.get('datalabels', {})
        if datalabel_opts.get('display'):
            plot = chart.plots[0]
            plot.has_data_labels = True
            data_labels = plot.data_labels
            
            font_opts = datalabel_opts.get('font', {})
            label_color = parse_color(datalabel_opts.get('color'))

            data_labels.font.bold = font_opts.get('weight') == 'bold'
            if font_opts.get('size'):
                data_labels.font.size = Pt(font_opts['size'])
            if label_color:
                data_labels.font.color.rgb = label_color
            
            apply_datalabel_positioning(
                data_labels,
                datalabel_opts.get('anchor', 'center'),
                datalabel_opts.get('align', 'center'),
                datalabel_opts.get('offset', 0),
                chart_type_str
            )

            formatter_str = datalabel_opts.get('formatter')
            if formatter_str:
                data_labels.show_category_name = False
                data_labels.show_value = False
                data_labels.show_percentage = False
                
                series = chart.series[0]
                for i, point in enumerate(series.points):
                    point.has_data_label = True
                    data_label = point.data_label
                    values = series.values
                    formatted_text = parse_datalabel_formatter(formatter_str, labels, values, i)
                    data_label.text_frame.text = formatted_text
                    
                    if data_label.text_frame.paragraphs:
                        p = data_label.text_frame.paragraphs[0]
                        if font_opts.get('size'):
                            p.font.size = Pt(font_opts['size'])
                        p.font.bold = font_opts.get('weight') == 'bold'
                        if label_color:
                            p.font.color.rgb = label_color

        try:
            if chart_type_str != 'pie':
                y_grid = value_scale.get('grid', {})
                x_grid = category_scale.get('grid', {})
                
                if y_grid.get('display', True) == False:
                    value_axis.has_major_gridlines = False
                elif y_grid.get('display', True) == True:
                    value_axis.has_major_gridlines = True
                    grid_color = parse_color(y_grid.get('color'))
                    if grid_color:
                        value_axis.major_gridlines.format.line.color.rgb = grid_color

                if x_grid.get('display', True) == False:
                    category_axis.has_major_gridlines = False
                elif x_grid.get('display', True) == True:
                    category_axis.has_major_gridlines = True
                    grid_color = parse_color(x_grid.get('color'))
                    if grid_color:
                        category_axis.major_gridlines.format.line.color.rgb = grid_color
                    
        except Exception as e:
            print(f"Error configuring gridlines: {e}")

        # Apply colors to series/points
        try:
            datasets = data.get('datasets', [])
            for i, series in enumerate(chart.series):
                if i < len(datasets):
                    dataset = datasets[i]
                    
                    # Line/Area chart colors
                    if chart_type in [XL_CHART_TYPE.LINE, XL_CHART_TYPE.LINE_MARKERS, XL_CHART_TYPE.AREA, XL_CHART_TYPE.AREA_STACKED]:
                        border_color = parse_color(dataset.get('borderColor'))
                        if border_color:
                            series.format.line.color.rgb = border_color
                        
                        if chart_type in [XL_CHART_TYPE.AREA, XL_CHART_TYPE.AREA_STACKED]:
                            bg_color = parse_color(dataset.get('backgroundColor'))
                            if bg_color:
                                series.format.fill.solid()
                                series.format.fill.fore_color.rgb = bg_color

                    # Bar/Column/Pie chart colors
                    elif chart_type in [XL_CHART_TYPE.COLUMN_CLUSTERED, XL_CHART_TYPE.BAR_CLUSTERED, XL_CHART_TYPE.PIE]:
                        background_colors = dataset.get('backgroundColor', [])
                        for j, point in enumerate(series.points):
                            if j < len(background_colors):
                                color = parse_color(background_colors[j])
                                if color:
                                    point.format.fill.solid()
                                    point.format.fill.fore_color.rgb = color
        except Exception as e:
            print(f"Error applying series colors: {e}")

    except Exception as e:
        print(f"Failed to add chart: {e}")
        return

def add_shape_element(slide, element, slide_width, slide_height):
    """Renders complex shapes based on shapeInfo from proto3."""
    shape_info = element.get('shapeInfo')
    if not shape_info:
        return

    rect = shape_info.get('rect', element)
    x = rect.get('x', 0)
    y = rect.get('y', 0)
    width = max(1, rect.get('width', 10))
    height = max(1, rect.get('height', 10))

    styles = shape_info.get('styles', element.get('styles', {}))
    rotation = shape_info.get('rotation', 0)
    clip_path = shape_info.get('clipPath', '')
    bg_color = parse_color(styles.get('backgroundColor'))

    shape_type = MSO_SHAPE.RECTANGLE # Default
    
    # Map common clip-paths to PowerPoint shapes
    if 'polygon' in clip_path:
        # A simple heuristic for chevron arrows
        if '50% 0%' in clip_path and '100% 50%' in clip_path and '50% 100%' in clip_path:
             shape_type = MSO_AUTO_SHAPE_TYPE.CHEVRON
        # Heuristic for block arrows
        elif '100% 50%' in clip_path and '75% 90%' in clip_path:
            shape_type = MSO_AUTO_SHAPE_TYPE.RIGHT_ARROW
        elif '0% 50%' in clip_path and '25% 90%' in clip_path:
            shape_type = MSO_AUTO_SHAPE_TYPE.LEFT_ARROW

    try:
        shape = slide.shapes.add_shape(
            shape_type,
            pixels_to_emu(x), pixels_to_emu(y),
            pixels_to_emu(width), pixels_to_emu(height)
        )

        if bg_color:
            shape.fill.solid()
            shape.fill.fore_color.rgb = bg_color
        else:
            shape.fill.background()
        
        shape.line.fill.background()
        shape.shadow.inherit = False
        
        if rotation != 0:
            shape.rotation = rotation

    except Exception as e:
        print(f"Failed to add shape element: {e}")

def create_pptx_from_json(json_path, output_path=None):
    """Enhanced PowerPoint generation with precise positioning"""
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            slides_data = json.load(f)
    except Exception as e:
        print(f"Error reading JSON file: {e}")
        return
   
    if not slides_data:
        print("No slides found in JSON")
        return
   
    # Get slide dimensions from first slide
    first_slide = slides_data[0]
    slide_width = safe_int(first_slide.get('slideWidth', 1039))
    slide_height = safe_int(first_slide.get('slideHeight', 585))
   
    # Create presentation with precise dimensions
    prs = Presentation()
    prs.slide_width = pixels_to_emu(slide_width)
    prs.slide_height = pixels_to_emu(slide_height)
   
    for slide_data in slides_data:
        slide_layout = prs.slide_layouts[6] # Blank layout
        slide = prs.slides.add_slide(slide_layout)
       
        # Set slide background
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = parse_color('#ffffff') or RGBColor(255, 255, 255)
       
        # Add slide background styling
        slide_styles = slide_data.get('slideStyles', {})
        if slide_styles:
            add_bg_shape(slide, slide_styles, 0, 0, slide_width, slide_height)
       
        elements = slide_data.get('elements', [])
       
        # Sort elements primarily by z-index, then by y, x for consistent layering
        elements_sorted = sorted(elements, key=lambda e: (e['zIndex'], e.get('y', 0), e.get('x', 0)))
       
        # Track processed elements to prevent duplicates
        processed_element_ids = set()
       
        # Process each element with enhanced positioning
        for element in elements_sorted:
            element_id = f"{element.get('type', 'unknown')}-{element.get('x', 0)}-{element.get('y', 0)}-{element.get('className', '')}"
            if element_id in processed_element_ids:
                continue
            processed_element_ids.add(element_id)
            
            element_type = element.get('type')
            parent_has_shadow = False
            
            # Handle complex shapes first (includes arrows)
            if element.get('shapeInfo'):
                add_shape_element(slide, element, slide_width, slide_height)
                continue

            if element.get('footerElements'):
                # Process individual footer elements with space-between positioning
                footer_elements = element.get('footerElements', [])
                for footer_elem in footer_elements:
                    if footer_elem.get('type') == 'img':
                        add_image_element(slide, footer_elem, slide_width, slide_height, parent_has_shadow)
                    elif footer_elem.get('type') in ['span', 'div']:
                        add_text_element(slide, footer_elem, slide_width, slide_height, parent_has_shadow)
                continue

            if element_type == 'overlay':
                add_overlay_element(slide, element, slide_width, slide_height)
            elif element.get('inlineGroup'):
                add_inline_group_element(slide, element, slide_width, slide_height, parent_has_shadow)
            elif element_type in ['ul', 'ol']:
                add_list_element(slide, element, slide_width, slide_height, parent_has_shadow)
            elif element_type == 'table':
                add_table_element(slide, element, slide_width, slide_height, parent_has_shadow)
            elif element_type == 'img':
                add_image_element(slide, element, slide_width, slide_height, parent_has_shadow)
            elif element_type == 'canvas':
                add_chart_element(slide, element, slide_width, slide_height)
            elif element_type == 'span':
                add_text_element(slide, element, slide_width, slide_height, parent_has_shadow)
            elif element_type == 'pseudo':
                add_pseudo_element(slide, element, slide_width, slide_height)
            elif element_type == 'div' and 'chart' in element.get('className', '') and element.get('chartConfig'):
                add_chart_element(slide, element, slide_width, slide_height)
            elif element_type in ['div', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6']:
                # Enhanced duplicate prevention for div elements
                if (element.get('text', '').strip() or
                    has_any_border(element.get('styles', {})) or
                    parse_color(element.get('styles', {}).get('backgroundColor')) or
                    element.get('styles', {}).get('boxShadow', 'none') != 'none'):
                    
                    # Skip if this is a div that only contains processed inline elements
                    if (element.get('text', '').strip() and not element.get('inlineGroup')):
                        add_text_element(slide, element, slide_width, slide_height, parent_has_shadow)
                    elif not element.get('text', '').strip():
                        # Check if it's a structural div that should be rendered
                        if not element.get('footerElements'):
                            x = element.get('x', 0)
                            y = element.get('y', 0)
                            width = max(1, element.get('width', 100))
                            height = max(1, element.get('height', 100))
                            add_bg_shape(slide, element.get('styles', {}), x, y, width, height)

    if output_path is None:
        base_name = os.path.splitext(os.path.basename(json_path))[0]
        output_path = f"{base_name}_output.pptx"
    try:
        prs.save(output_path)
        print(f"Presentation saved as '{output_path}' with {len(slides_data)} slide(s)")
        print(f"Slide dimensions: {slide_width}x{slide_height} pixels")
    except Exception as e:
        print(f"Error saving presentation: {e}")

if __name__ == "__main__":
    create_pptx_from_json('slides_data.json', 'output.pptx')