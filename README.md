# MedAI Premium PPTX Renderer

Server-side PPTX rendering service for MedAI MAP Generator Premium.
Produces designer-quality presentations using python-pptx with matplotlib charts.

## Architecture

```
map_generator_premium.html (Browser)
    │
    ├── Step 1-4: Same as Standard (Claude API → JSON)
    │
    └── Step 5: POST /render-pptx ──→ This Service (Railway)
                                          ├── python-pptx rendering
                                          ├── matplotlib charts (KM, bars, SWOT)
                                          ├── Template-based (Phase 2)
                                          └── ← Returns .pptx file
```

## Quick Start (Local)

```bash
pip install -r requirements.txt
uvicorn app.main:app --reload --port 8080
```

Then open: http://localhost:8080/docs (Swagger UI)

## API Endpoints

### `POST /render-pptx`
Accepts the same JSON format as "Save JSON" in map_generator.html.

```json
{
  "template_id": "medai_dark",
  "meta": {
    "drug": "Belantamab Mafodotin",
    "indication": "Multiple Myeloma",
    "year": 2027,
    "country": "Germany",
    "model": "claude-opus-4-6"
  },
  "sections": [
    {"id": "executive_summary", "label": "Executive Summary", "icon": "📋", "data": {...}},
    ...
  ]
}
```

Returns: `.pptx` file as binary download.

### `GET /templates`
Lists available themes with color palettes.

### `POST /validate-template`
Checks if a template has all required named shapes.

### `GET /health`
Health check.

## Deploy to Railway

1. Push this directory to a GitHub repo
2. Connect to Railway: `railway link`
3. Deploy: `railway up`
4. Note the URL: `https://medai-pptx-premium-production.up.railway.app`
5. Update `map_generator_premium.html` with the URL

## Template System

See `templates/TEMPLATE_SPEC.md` for the complete designer guide.

**Phase 1 (now):** Programmatic rendering — python-pptx builds slides from code.
Already higher quality than PptxGenJS because of proper font embedding and XML control.

**Phase 2 (next):** Designer templates — real .pptx files with named shapes.
A designer creates beautiful slides in PowerPoint. The renderer fills them with data.

**Phase 3 (enterprise):** Custom template upload — customers provide their own .pptx.

## Available Themes

| ID | Name | Background | Accent |
|----|------|-----------|--------|
| dark | MedAI Dark | Navy #0D2B4E | Teal + Purple |
| light | Clean Light | White #FFFFFF | Blue + Teal |
| gray | Consulting | Anthracite #34495E | Red + Teal |
| pharma | Pharma Blue | Light Blue #F0F9FF | Royal Blue |
| premium | Black & Gold | Black #111111 | Gold |

## Integration with Frontend

In `map_generator_premium.html`, the `premiumExport()` function calls this service:

```javascript
const res = await fetch('https://YOUR-RAILWAY-URL/render-pptx', {
  method: 'POST',
  headers: {'Content-Type': 'application/json'},
  body: JSON.stringify({
    template_id: S.premiumTemplate,
    meta: { drug: S.drug, indication: S.indication, year: S.year, country: S.country, model: S.model },
    sections: S.generatedContent
  })
});
const blob = await res.blob();
// Trigger download...
```
