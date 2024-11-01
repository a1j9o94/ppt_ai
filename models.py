from pydantic import BaseModel, Field, field_validator
from typing import List, Optional, Literal
from enum import Enum

class LayoutType(str, Enum):
    COLUMN = "columns"
    ROW = "rows"

class Section(BaseModel):
    header: str
    content: List[str]
    size: Optional[float] = None

    @field_validator('size')
    @classmethod
    def validate_size(cls, v):
        if v is not None and (v <= 0 or v > 100):
            raise ValueError("Size must be between 0 and 100")
        return v

class SlideModel(BaseModel):
    slide_title: str
    layout: Optional[LayoutType] = None
    columns: Optional[int] = None
    rows: Optional[int] = None
    sections: Optional[List[Section]] = Field(default_factory=list)
    ppt_name: Optional[str] = None

    @field_validator('columns', 'rows')
    @classmethod
    def validate_dimension(cls, v):
        if v is not None and v <= 0:
            raise ValueError("Dimension must be greater than 0")
        return v

    @field_validator('sections')
    @classmethod
    def validate_sections(cls, v, info):
        if not v:
            return v

        values = info.data
        # Check if both rows and columns are specified
        if values.get('rows') and values.get('columns'):
            raise ValueError("Cannot specify both rows and columns")

        # Validate section count matches specified dimension
        if values.get('columns') and len(v) != values['columns']:
            raise ValueError(f"Number of sections must match number of columns ({values['columns']})")
        if values.get('rows') and len(v) != values['rows']:
            raise ValueError(f"Number of sections must match number of rows ({values['rows']})")

        # Validate section sizes
        sizes = [section.size for section in v if section.size is not None]
        if sizes:
            if len(sizes) != len(v):
                raise ValueError("Cannot mix sized and unsized sections")
            total_size = sum(sizes)
            if total_size < 98 or total_size > 102:
                raise ValueError("Section sizes must sum to approximately 100%")

        return v

class PresentationModel(BaseModel):
    name: str
    title: str = "New Presentation"
    slides: List[SlideModel] = Field(default_factory=list)

class PresentationRequest(BaseModel):
    presentation: PresentationModel