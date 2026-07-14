"""Pydantic models for the rendering API."""

from pydantic import BaseModel
from typing import Optional


class RenderRequest(BaseModel):
    template_id: Optional[int] = None
    xlsx_path: Optional[str] = None
    output_dir: Optional[str] = None


class FieldModel(BaseModel):
    id: str
    label: str
    left_px: float
    top_px: float
    width_px: float
    height_px: float
    type: str = "text"
    required: bool = False


class RenderResponse(BaseModel):
    page_width: int
    page_height: int
    background_image: str
    debug_image: Optional[str] = None
    fields: list[FieldModel]
