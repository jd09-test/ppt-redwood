from pydantic import BaseModel, Field
from typing import List, Optional, Literal

class Placeholder(BaseModel):
    placeholder_name: str
    content: List[str] | None 

class Slide(BaseModel):
    slide_number: int = Field(..., gt=0)
    layout: str  # Choose a suitable layout name that fits the content, strictly from the available "layout_name" list below
    placeholders: List[Placeholder] 
    speaker_notes: Optional[str] = None  # Optional speaker notes for this slide

class PresentationContent(BaseModel):
    filename: str
    theme_mode: Literal['light', 'dark']
    slides: List[Slide]
