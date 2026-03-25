"""Pydantic schemas for MAP Premium PPTX Renderer."""
from typing import Optional, Any
from pydantic import BaseModel, Field


class RenderRequest(BaseModel):
    """Request body for /render-pptx endpoint."""
    template_id: str = Field(default="medai_dark", description="Template theme ID")
    meta: dict = Field(description="MAP metadata: drug, indication, year, country, etc.")
    sections: list[dict] = Field(description="Array of section objects from Step 4 review JSON")


class SectionData(BaseModel):
    """One section from the MAP generator output."""
    id: str
    label: str
    icon: Optional[str] = ""
    data: dict = Field(default_factory=dict)


class ValidateTemplateRequest(BaseModel):
    """Request body for /validate-template endpoint."""
    template_id: str


class ValidateTemplateResponse(BaseModel):
    """Response from template validation."""
    valid: bool
    template_id: str
    missing_shapes: list[str] = []
    found_shapes: list[str] = []
    message: str = ""


class HealthResponse(BaseModel):
    service: str = "medai-pptx-premium"
    status: str = "ok"
    version: str = "1.0.0"
    templates_available: list[str] = []
