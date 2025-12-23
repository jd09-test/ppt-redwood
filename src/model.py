from pydantic import BaseModel, Field
from typing import List, Optional, Literal

class Placeholder(BaseModel):
    placeholder_name: str
    content: List[str] | None 

class Slide(BaseModel):
    slide_number: int = Field(..., gt=0)
    layout: str 
    placeholders: List[Placeholder]

class PresentationContent(BaseModel):
    filename: str
    theme_mode: Literal['light', 'dark']
    slides: List[Slide]