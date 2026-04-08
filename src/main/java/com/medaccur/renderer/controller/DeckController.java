package com.medaccur.renderer.controller;

import com.medaccur.renderer.model.RenderRequest;
import com.medaccur.renderer.service.TemplateRenderService;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.http.*;
import org.springframework.web.bind.annotation.*;

import java.util.Map;

@RestController
@RequestMapping("/api/v1")
public class DeckController {

    private static final Logger log = LoggerFactory.getLogger(DeckController.class);
    private final TemplateRenderService renderService;

    public DeckController(TemplateRenderService renderService) {
        this.renderService = renderService;
    }

    @PostMapping("/render-deck")
    public ResponseEntity<byte[]> renderDeck(@RequestBody RenderRequest request) {
        long start = System.currentTimeMillis();
        log.info("Render request: {} slides", request.getSlides().size());

        try {
            byte[] pptxBytes = renderService.renderDeck(request);
            long elapsed = System.currentTimeMillis() - start;
            log.info("Rendered in {}ms ({} KB)", elapsed, pptxBytes.length / 1024);

            String filename = buildFilename(request);
            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.parseMediaType(
                "application/vnd.openxmlformats-officedocument.presentationml.presentation"));
            headers.setContentDisposition(ContentDisposition.attachment().filename(filename).build());
            headers.setContentLength(pptxBytes.length);

            return new ResponseEntity<>(pptxBytes, headers, HttpStatus.OK);

        } catch (Exception e) {
            log.error("Render failed: {}", e.getMessage(), e);
            return ResponseEntity.status(500)
                .contentType(MediaType.APPLICATION_JSON)
                .body(("{\"error\":\"" + e.getMessage().replace("\"", "'") + "\"}").getBytes());
        }
    }

    @GetMapping("/health")
    public ResponseEntity<Map<String, Object>> health() {
        return ResponseEntity.ok(Map.of(
            "status", "ok",
            "engine", "Apache POI 5.4",
            "version", "3.0.0",
            "templateLoaded", renderService.isTemplateLoaded(),
            "layoutCount", renderService.getLayoutCount(),
            "layouts", renderService.getAvailableLayouts()
        ));
    }

    private String buildFilename(RenderRequest request) {
        String drug = request.getMetadata().getOrDefault("drug", "Export");
        String module = request.getMetadata().getOrDefault("module", "deck");
        return drug.replaceAll("[^a-zA-Z0-9]", "_") + "_" + module + ".pptx";
    }
}
