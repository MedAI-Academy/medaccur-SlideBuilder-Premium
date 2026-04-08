# medaccur SlideBuilder Premium

Apache POI slide rendering engine for medaccur. Takes JSON input, fills PowerPoint templates with data, outputs PPTX.

## Stack
- Java 17 + Spring Boot 3.2
- Apache POI 5.4 (poi-ooxml-full)
- Deployed on Railway via Dockerfile

## API

### `POST /api/v1/render-deck`
Renders a complete slide deck. Returns binary PPTX.

### `GET /api/v1/health`
Returns engine status, template info, available layouts.

## Architecture
```
medaccur_master.pptx (25+ layouts with {{placeholders}})
    ↓
TemplateRenderService opens template, creates slides from layouts
    ↓
PlaceholderService replaces {{key}} → value (preserving formatting)
    ↓
ChartDataService updates embedded Excel charts (by alt-text name)
    ↓
FigureService inserts PNG figures (KM curves, Forest plots)
    ↓
Output: finished PPTX
```

## Rules
1. **NEVER modify template colors or themes**
2. **NEVER programmatically create layouts**
3. **ONLY replace text + insert images + update chart data**
4. **Template is READ-ONLY** — design changes only in PowerPoint
