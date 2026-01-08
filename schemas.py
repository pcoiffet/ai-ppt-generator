"""
Pydantic models for presentation structure validation.
"""
from typing import List, Optional, Union, Literal
from pydantic import BaseModel, Field, field_validator, model_validator


class TextFormatting(BaseModel):
    bold: bool = False
    italic: bool = False
    color: Optional[str] = Field(None, description="Hex color code #RRGGBB")
    size: Optional[float] = None


class TextRun(BaseModel):
    text: str
    formatting: Optional[TextFormatting] = None
    hyperlink: Optional[str] = None


class BulletPoint(BaseModel):
    text: str
    level: int = Field(default=0, ge=0, le=5, description="Indentation level 0-5")
    formatting: Optional[TextFormatting] = None


class TableData(BaseModel):
    headers: List[str] = Field(..., min_length=1)
    rows: List[List[Union[str, int, float]]] = Field(..., min_length=1)
    style: Optional[str] = None


class ChartSeries(BaseModel):
    name: str
    data: List[Union[int, float]]


class ChartData(BaseModel):
    type: Literal['column', 'line', 'pie'] = 'column'
    categories: List[str] = Field(..., min_length=1)
    series: List[ChartSeries] = Field(..., min_length=1)


class ImageData(BaseModel):
    path: str = Field(..., description="Descriptive keywords for image search")
    position: Literal['left', 'right', 'full'] = 'right'


class SlideContent(BaseModel):
    title: str = Field(..., min_length=1, max_length=200)
    content: Optional[Union[str, List[TextRun]]] = None
    bullet_points: Optional[List[Union[str, BulletPoint]]] = None
    table: Optional[TableData] = None
    chart: Optional[ChartData] = None
    image: Optional[ImageData] = None
    layout: Optional[str] = Field(None, description="Layout hint: content_only, image_left, image_right, etc.")

    @field_validator('content', mode='before')
    @classmethod
    def normalize_content(cls, v):
        if isinstance(v, dict) and 'runs' in v:
            return v['runs']
        return v
    
    @field_validator('bullet_points', mode='before')
    @classmethod
    def normalize_bullet_points(cls, v):
        if isinstance(v, list):
            return [
                BulletPoint(text=item) if isinstance(item, str) else item
                for item in v
            ]
        return v
    
    @model_validator(mode='after')
    def validate_has_content(self):
        """Ensure slide has at least one content element."""
        has_content = any([
            self.content,
            self.bullet_points,
            self.table,
            self.chart,
            self.image
        ])
        if not has_content:
            raise ValueError(f"Slide '{self.title}' must have at least one content element")
        return self


class PresentationInput(BaseModel):
    """Main presentation model."""
    title: str = Field(..., min_length=1, max_length=200)
    subtitle: Optional[str] = None
    author: Optional[str] = None
    subject: Optional[str] = None
    slides: List[SlideContent] = Field(..., min_length=1)


class PPTGenerationError(Exception):
    """Error during PowerPoint generation."""
    def __init__(self, message: str, details: dict = None):
        self.message = message
        self.details = details or {}
        super().__init__(message)
