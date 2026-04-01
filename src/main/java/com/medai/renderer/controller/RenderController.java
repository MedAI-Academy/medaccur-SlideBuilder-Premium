package com.medai.renderer.controller;

import com.medai.renderer.model.RenderRequest;
import com.medai.renderer.service.PptxRenderService;
import com.medai.renderer.service.TemplateRenderService;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.util.Map;

@RestController
public class RenderController {

    private static final Logger log = LoggerFactory.getLogger(RenderController.class);
    private static final String PPTX_CONTENT_TYPE =
        "application/vnd.openxmlformats-officedocument.presentationml.presentation";

    private final PptxRenderService renderService;
    private final TemplateRenderService templateService;

    public RenderController(PptxRenderService renderService, TemplateRenderService templateService) {
        this.renderService = renderService;
        this.templateService = templateService;
    }

    // ═══════════════════════════════════════════════════════════
    // NEW: Template-based rendering (medaccur custom templates)
    // ═══════════════════════════════════════════════════════════

    @PostMapping("/api/v1/render-deck")
    public ResponseEntity<byte[]> renderDeck(@RequestBody RenderRequest request) {
        long start = System.currentTimeMillis();
        log.info("Template render-deck: slides={}, drug={}",
            request.getSlides() != null ? request.getSlides().size() : 0,
            request.meta("drug"));

        try {
            if (request.getSlides() == null || request.getSlides().isEmpty()) {
                return ResponseEntity.badRequest()
                    .body("{\"error\": \"No slides provided\"}".getBytes());
            }

            byte[] pptxBytes = templateService.renderDeck(request);
            String filename = buildFilename(request);

            long elapsed = System.currentTimeMillis() - start;
            log.info("Template render complete: {} slides, {}KB, {}ms",
                request.getSlides().size(), pptxBytes.length / 1024, elapsed);

            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.parseMediaType(PPTX_CONTENT_TYPE));
            headers.setContentDispositionFormData("attachment", filename);
            headers.setContentLength(pptxBytes.length);
            headers.add("Access-Control-Expose-Headers", "Content-Disposition");

            return new ResponseEntity<>(pptxBytes, headers, HttpStatus.OK);

        } catch (Exception e) {
            log.error("Template render failed", e);
            String errorJson = "{\"error\": \"" + e.getMessage().replace("\"", "'") + "\"}";
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                .body(errorJson.getBytes());
        }
    }

    // ═══════════════════════════════════════════════════════════
    // EXISTING: Programmatic rendering (legacy — kept for other modules)
    // ═══════════════════════════════════════════════════════════

    @PostMapping("/api/v1/render")
    public ResponseEntity<byte[]> render(@RequestBody RenderRequest request) {
        long start = System.currentTimeMillis();
        log.info("Programmatic render: module={}, slides={}",
            request.getModule(),
            request.getSlides() != null ? request.getSlides().size() : 0);

        try {
            if (request.getSlides() == null || request.getSlides().isEmpty()) {
                return ResponseEntity.badRequest()
                    .body("{\"error\": \"No slides provided\"}".getBytes());
            }

            byte[] pptxBytes = renderService.render(request);
            String filename = buildFilename(request);

            long elapsed = System.currentTimeMillis() - start;
            log.info("Render complete: {} slides, {}KB, {}ms",
                request.getSlides().size(), pptxBytes.length / 1024, elapsed);

            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.parseMediaType(PPTX_CONTENT_TYPE));
            headers.setContentDispositionFormData("attachment", filename);
            headers.setContentLength(pptxBytes.length);
            headers.add("Access-Control-Expose-Headers", "Content-Disposition");

            return new ResponseEntity<>(pptxBytes, headers, HttpStatus.OK);

        } catch (Exception e) {
            log.error("Render failed", e);
            String errorJson = "{\"error\": \"" + e.getMessage().replace("\"", "'") + "\"}";
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                .body(errorJson.getBytes());
        }
    }

    @PostMapping("/render")
    public ResponseEntity<byte[]> legacyRender(@RequestBody RenderRequest request) {
        return render(request);
    }

    // ═══════════════════════════════════════════════════════════
    // HEALTH & INFO
    // ═══════════════════════════════════════════════════════════

    @GetMapping("/api/v1/health")
    public ResponseEntity<Map<String, Object>> health() {
        return ResponseEntity.ok(Map.of(
            "status", "ok",
            "engine", "Apache POI 5.4 + medaccur Templates",
            "version", "2.0.0",
            "endpoints", Map.of(
                "template_render", "/api/v1/render-deck",
                "programmatic_render", "/api/v1/render",
                "legacy", "/render"
            ),
            "features", Map.of(
                "templateCloneSwap", true,
                "kaplanMeier", true,
                "autoSelectLayout", true,
                "confidenceScore", true,
                "widescreen", true
            )
        ));
    }

    @GetMapping("/api/v1/templates")
    public ResponseEntity<Map<String, Object>> templates() {
        return ResponseEntity.ok(Map.of(
            "engine", "medaccur-clone-swap",
            "templateLayouts", new String[]{
                "TITLE", "SECTION_DIVIDER", "FOREST_PLOT", "SWIMMER_PLOT",
                "EXECUTIVE_SUMMARY", "COMPETITIVE_LANDSCAPE",
                "SWOT_CLASSIC", "SWOT_HORIZONTAL", "SWOT_DIAMOND",
                "DISEASE_OVERVIEW", "SCIENTIFIC_NARRATIVE", "PIVOTAL_STUDIES",
                "PREVALENCE_KPI", "GUIDELINES_TABLE", "UNMET_NEEDS", "MARKET_ACCESS",
                "WATERFALL_PLOT",
                "IMPERATIVES_TEMPLE_3", "IMPERATIVES_TEMPLE_4", "IMPERATIVES_TEMPLE_5",
                "IMPERATIVES_NUMBERED_4", "IMPERATIVES_NUMBERED_COMPACT",
                "TACTICAL_PLAN_FULL", "TACTICAL_PLAN_COMPACT"
            },
            "totalLayouts", 24,
            "autoSelect", Map.of(
                "strategic_imperatives", "3/4/5 pillars → matching template",
                "tactical_plan", "≤5 workstreams → compact, ≥6 → full",
                "swot", "default=classic, alt=horizontal/diamond"
            )
        ));
    }

    @GetMapping("/")
    public ResponseEntity<Map<String, String>> root() {
        return ResponseEntity.ok(Map.of(
            "service", "medaccur PPTX Renderer",
            "version", "2.0.0 (Java/Apache POI + Template Clone+Swap)",
            "docs", "/api/v1/templates",
            "health", "/api/v1/health",
            "render", "POST /api/v1/render-deck"
        ));
    }

    private String buildFilename(RenderRequest request) {
        String drug = request.meta("drug").replaceAll("[^a-zA-Z0-9]", "_");
        String module = request.getModule() != null ? request.getModule() : "MAP";
        String year = request.meta("year");
        if (drug.isEmpty()) drug = "medaccur";
        if (year.isEmpty()) year = String.valueOf(java.time.Year.now().getValue());
        String country = request.meta("country").replaceAll("[^a-zA-Z0-9]", "_");
        if (!country.isEmpty()) {
            return drug + "_" + module + "_" + country + "_" + year + ".pptx";
        }
        return drug + "_" + module + "_" + year + ".pptx";
    }
}
