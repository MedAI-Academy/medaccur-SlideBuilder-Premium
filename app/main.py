"""MedAI Premium PPTX Renderer — FastAPI Service.

Endpoints:
  POST /render-pptx     → Render MAP data into premium PPTX (returns file)
  POST /validate-template → Check if a template has all required shapes
  GET  /health           → Health check with available templates
  GET  /templates        → List available templates with metadata
"""
import time
from fastapi import FastAPI, HTTPException
from fastapi.responses import Response, JSONResponse
from fastapi.middleware.cors import CORSMiddleware

from .schemas import RenderRequest, ValidateTemplateRequest, ValidateTemplateResponse, HealthResponse
from .renderer import render_pptx, PALETTES
from .template_manager import list_available_themes, validate_template, get_full_template_path

app = FastAPI(
    title="MedAI Premium PPTX Renderer",
    description="Server-side PPTX rendering for MedAI MAP Generator Premium",
    version="1.0.0",
)

# CORS — allow calls from Netlify and localhost
app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "https://medai-dashboard.netlify.app",
        "http://localhost:*",
        "http://127.0.0.1:*",
        "*",  # During development; restrict in production
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/health", response_model=HealthResponse)
async def health():
    themes_with_full = [t for t in list_available_themes() if get_full_template_path(t)]
    return HealthResponse(
        templates_available=list(PALETTES.keys()),
        inplace_themes=themes_with_full,
    )


@app.get("/templates")
async def list_templates():
    """List all available template themes with metadata."""
    templates = []
    for key, pal in PALETTES.items():
        has_full = get_full_template_path(key) is not None
        templates.append({
            "id": key,
            "name": {
                "dark": "MedAI Dark",
                "light": "MedAI Clean Light",
                "gray": "McKinsey / BCG Style",
                "pharma": "Pharma Corporate Blue",
                "premium": "Black & Gold Premium",
                "normal": "MedAI Normal",
                "gold": "MedAI Gold",
                "aquarell": "MedAI Aquarell",
            }.get(key, key),
            "colors": {
                "bg": f"#{pal['bg']}",
                "accent": f"#{pal['accent']}",
                "teal": f"#{pal['teal']}",
                "text": f"#{pal['text']}",
            },
            "has_designer_templates": key in list_available_themes(),
            "inplace_ready": has_full,
            "render_mode": "inplace" if has_full else "programmatic",
        })
    return {"templates": templates}


@app.post("/render-pptx")
async def render_pptx_endpoint(request: RenderRequest):
    """
    Render a premium PPTX from MAP data.

    Accepts the same JSON format as the "Save JSON" button in map_generator.html.
    Returns a .pptx file as binary response.
    """
    t0 = time.time()

    if not request.sections:
        raise HTTPException(status_code=400, detail="No sections provided")

    meta = request.meta
    if not meta.get("drug"):
        raise HTTPException(status_code=400, detail="meta.drug is required")

    try:
        pptx_bytes = render_pptx(
            meta=meta,
            sections=request.sections,
            template_id=request.template_id,
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Render failed: {str(e)}")

    elapsed = time.time() - t0
    drug_name = meta.get("drug", "MAP").replace(" ", "_")
    year = meta.get("year", "")
    filename = f"{drug_name}_Premium_{year}.pptx"

    return Response(
        content=pptx_bytes,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"',
            "X-Render-Time": f"{elapsed:.2f}s",
            "X-Sections": str(len(request.sections)),
            "X-Template": request.template_id,
        },
    )


@app.post("/validate-template", response_model=ValidateTemplateResponse)
async def validate_template_endpoint(request: ValidateTemplateRequest):
    """Validate that a custom template has all required named shapes."""
    # For now, validate built-in themes
    result = validate_template(request.template_id, "_title")
    return ValidateTemplateResponse(**result)
