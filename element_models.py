from datetime import datetime
from pathlib import Path
import random
from pydantic import BaseModel, Field, root_validator
from typing import Any, Union, List
from pptx.util import Inches, Cm, Pt, Emu, Mm, Centipoints
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_UNDERLINE
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_MARKER_STYLE, XL_TICK_MARK, XL_TICK_LABEL_POSITION, XL_LEGEND_POSITION
from enum import Enum



class PPTRun(BaseModel):
    text: str = ""
    font_size: int = None
    font_name: str = None
    bold: bool = None
    italic: bool = None
    underline: MSO_UNDERLINE = None
    color: Union[list, tuple] = None
    link: str = None
    exclude_in_translation: bool = False



class PPTPara(BaseModel):
    runs: Union[str, PPTRun, List[Union[str, PPTRun]]] = []
    default_run_style: PPTRun = Field(default_factory=PPTRun)
    level: int = None
    alignment: str = None
    space_before: int = None 
    space_after: int = None
    line_spacing: int = None
    bullet_char: bool = None 

    def __init__(self, **data):
        run_style_keys = {"font_size", "font_name", "bold", "italic", "underline", "color", "link"}
        style_args = {k: data.pop(k) for k in list(data) if k in run_style_keys}
        
        if "default_run_style" not in data and style_args:
            data["default_run_style"] = PPTRun(**style_args)
        super().__init__(**data)

    def dict(self, *args, **kwargs):
        runs = self._convert_runs()
        exclude = kwargs.get("exclude")
        if exclude is None:
            exclude = set()
        dict_data =  {
            "runs": [r.dict() for r in runs],
            "level": self.level,
            "alignment": self.alignment,
            "space_before": self.space_before,
            "space_after": self.space_after,
            "line_spacing": self.line_spacing,
            "bullet_char": self.bullet_char,
        }
        return { k:v for k, v in dict_data.items() if k not in exclude}

    def _convert_runs(self) -> List[PPTRun]:
        if isinstance(self.runs, PPTRun):
            return [self.runs]
        if isinstance(self.runs, str):
            return [self._apply_style(self.runs)]
        if isinstance(self.runs, list):
            return [r if isinstance(r, PPTRun) else self._apply_style(r) for r in self.runs]
        return []

    def _apply_style(self, text: str) -> PPTRun:
        return PPTRun(text=text, **self.default_run_style.dict(exclude={"text"}))
          
      
class PPTText(BaseModel):
    type: str = 'text'
    # Pydantic tries to coerce the data according to order defined in annotation, 
    # Overriding the validation in parser
    # paras: Union[List[PPTPara], List[PPTRun], List[str], PPTPara, PPTRun, str] = []
    default_para_style: PPTPara = Field(default_factory=PPTPara)
    bg_color: Union[list, tuple] = None
    vertical_anchor: str = None 
    margin_top: Union[Inches, Cm, Pt, Emu] = None
    margin_bottom: Union[Inches, Cm, Pt, Emu] = None
    margin_left: Union[Inches, Cm, Pt, Emu] = None
    margin_right: Union[Inches, Cm, Pt, Emu] = None
    border_color: str = "f2f2f2"
    
    # only for the cases when we pass text into table, no error in textframes
    row_span: int = None
    col_span: int = None 

    # only for floating text boxes
    height :  Union[Inches, Cm, Pt, Emu] = None
    width :  Union[Inches, Cm, Pt, Emu] = None
    left :  Union[Inches, Cm, Pt, Emu] = None
    top :  Union[Inches, Cm, Pt, Emu] = None
    word_wrap: bool = True 

    paras: Any  # Let us handle validation manually


    @root_validator(pre=True)
    def parse_paras(cls, values):
        raw = values.get("paras")

        if isinstance(raw, list) and all(isinstance(p, PPTPara) for p in raw):
            values["paras"] = raw
        elif isinstance(raw, list) and all(isinstance(p, PPTRun) for p in raw):
            values["paras"] = raw
        elif isinstance(raw, list) and all(isinstance(p, str) for p in raw):
            values["paras"] = raw
        elif isinstance(raw, (PPTPara, PPTRun, str)):
            values["paras"] = [raw]
        else:
            raise TypeError(f"Unsupported paras input: {raw}")

        return values


    def __init__(self, **data):
        para_style_keys = {
                            "level", 
                            "alignment", 
                            "space_before", 
                            "space_after", 
                            "line_spacing",
                            "bullet_char", 
                            "font_size", 
                            "font_name", 
                            "bold", 
                            "italic",
                            "underline",
                            "color", 
                            "link"
                        }
        
        style_args = {k: data.pop(k) for k in list(data) if k in para_style_keys}
        if "default_para_style" not in data and style_args:
            data["default_para_style"] = PPTPara(**style_args)
        super().__init__(**data)

    def dict(self, *args, **kwargs):
        paras = self._convert_para()
        dict_data = {
            "type": "text",
            "content": {
                "paragraphs" : [ p.dict() for p in paras ]    
            },
            "bg_color": self.bg_color,
            "vertical_anchor": self.vertical_anchor,
            "margin_top": self.margin_top,
            "margin_bottom": self.margin_bottom,
            "margin_left": self.margin_left,
            "margin_right": self.margin_right,
            "border_color": self.border_color,
            "row_span": self.row_span,
            "col_span": self.col_span,
            "height": self.height,
            "width": self.width,
            "left": self.left,
            "top": self.top,
            "word_wrap": self.word_wrap,
        }
        exclude = kwargs.get("exclude")
        if exclude is None:
            exclude = set()
        return { k:v for k, v in dict_data.items() if k not in exclude}
        
    def _convert_para(self) -> List[PPTPara]:
        if isinstance(self.paras, PPTPara):
            return [self.paras]
        if isinstance(self.paras, PPTRun):
            return [self._apply_style(self.paras)]
        if isinstance(self.paras, str):
            return [self._apply_style(self.paras)]
        if isinstance(self.paras, list):
            return [p if isinstance(p, PPTPara) else self._apply_style(p) for p in self.paras]
        
        return []

    def _apply_style(self, runs: str) -> PPTPara:
        return  PPTPara(runs=runs, 
                       **self.default_para_style.dict(exclude={"runs"}), 
                       default_run_style = self.default_para_style.default_run_style
                    )


class PPTTitle(BaseModel):
    type: str = 'title'
    text: str 
    font_size: int = None
    font_name: str = None
    bold: bool = None
    font_color: Union[list, tuple] = None
    
class PPTTable(BaseModel):
    type: str = 'table' 
    column_widths: List[ Union[Inches, Cm, Pt, Emu] ]
    row_heights : List[ Union[Inches, Cm, Pt, Emu] ]
    table_data : List[List[Any]]
    
    
class TableSkipCell(BaseModel):
    type: str = 'skip'

class PPTImage(BaseModel):
    type: str = 'img'
    path: str 
    crop_circle: bool = False
    backup_path : str = None
    # only useful in table cells
    height: Emu = None
    width: Emu = None 
    vertical_position: Union[Centipoints, Pt, Mm, Cm, Emu, Inches, str] = None
    horizontal_position: Union[Centipoints, Pt, Mm, Cm, Emu, Inches, str] = None
    bg_color: Union[list, tuple] = None
    border_color: str = "f2f2f2"


class PPTImageFree(BaseModel):
    type: str = 'img_without_placeholder'
    path: str 
    crop_circle: bool = False
    backup_path : str = None
    # only useful in table cells
    height: Emu = None
    width: Emu = None 
    top: Union[Centipoints, Pt, Mm, Cm, Emu, Inches] 
    left: Union[Centipoints, Pt, Mm, Cm, Emu, Inches]


class Flaticon(BaseModel):
    type : str = 'flaticon'
    query : str  
    icon_code : str = None
    path : str = None
    backup_path : str = Field(default_factory = lambda: random.choice(placeholder_flaticons))
    # only useful in table cells
    height: Emu = None
    width: Emu = None 
    vertical_position: Union[Centipoints, Pt, Mm, Cm, Emu, Inches, str] = None
    horizontal_position: Union[Centipoints, Pt, Mm, Cm, Emu, Inches, str] = None
    bg_color: Union[list, tuple] = None

class PPTImgAndText(BaseModel):
    """Only to be used in Table as cells"""
    type : str = 'text_and_img' 
    text : PPTText
    image : PPTImage
    bg_color: Union[list, tuple] = None
    

class PPTChartType(str, Enum):
    BAR = 'bar_chart'
    LINE = 'line_chart'
    COLUMN_CHART = 'column_chart'
    AREA_CHART = 'area_chart'
    COLUMN_STACK_CHART = 'column_stack_chart'
    BAR_STACK_100_CHART = 'bar_stack_100_chart'


class PPTShapeSolidFill(BaseModel):
    type: str = 'solid_fill'
    color : list[int]


class PPTGradientStop(BaseModel):
    position: float
    color: list[int]


class PPTShapeGradientFill(BaseModel):
    # minimum two stops needed
    type: str = 'gradient_fill'
    stop_0 : PPTGradientStop = PPTGradientStop(position=0, color=[112, 148, 123])
    stop_1 : PPTGradientStop = PPTGradientStop(position=1, color=[112, 148, 123])
    intermediate_stops : list[PPTGradientStop] = []
    gradient_angle : int = 0

class PPTShape(BaseModel):
    type : str = 'shape'
    shape_type : MSO_SHAPE = MSO_SHAPE.RECTANGLE
    fill : Union[PPTShapeSolidFill, PPTShapeGradientFill] = PPTShapeSolidFill(color=[255, 255, 255])
    shadow : bool = False
    border_color : List[int] = None
    border_width : Union[Inches, Cm, Pt, Emu] = None
    height :  Union[Inches, Cm, Pt, Emu]
    width :  Union[Inches, Cm, Pt, Emu]
    left :  Union[Inches, Cm, Pt, Emu]
    top :  Union[Inches, Cm, Pt, Emu]


class PPTChart(BaseModel):
    type: PPTChartType 
    categories: list 
    series : List[list]
    labels : list = None
    minimum_scale: int = None
    maximum_scale: int = None
    number_format : str = None
    title : str = ''
    label_color : list[list] = None
    legend : bool = False
    labels_size : int = 8
    label_position : XL_LABEL_POSITION = XL_LABEL_POSITION.OUTSIDE_END
    series_color : list[list] = None
    bar_color : list[int] = None
    line_color : list[int] = None
    point_color : list[int] = None
    area_color: Union[list[int], PPTShapeGradientFill] = None
    negative_bar_color : list[int] = None
    has_data_labels: bool = True
    has_series_axis : bool = None
    has_category_axis : bool = None
    gap_width: int = None
    category_axis_label_position: XL_TICK_LABEL_POSITION = None # if set to XL_TICK_LABEL_POSITION.NONE hides labels
    value_axis_label_position:XL_TICK_LABEL_POSITION =  None

