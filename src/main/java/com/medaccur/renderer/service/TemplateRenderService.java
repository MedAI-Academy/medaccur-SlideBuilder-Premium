package com.medaccur.renderer.service;

import com.medaccur.renderer.model.RenderRequest;
import com.medaccur.renderer.model.SlideSpec;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.xslf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.io.Resource;
import org.springframework.stereotype.Service;

import jakarta.annotation.PostConstruct;
import java.awt.Dimension;
import java.io.*;
import java.util.*;

/**
 * Core rendering service. Opens the master template, creates slides from
 * SlideLayouts, fills placeholders, updates charts, inserts figures.
 *
 * CRITICAL RULES (from 6 weeks of failures):
 * 1. NEVER modify template colors or themes
 * 2. NEVER programmatically create layouts
 * 3. ONLY replace text + insert images + update chart data
 * 4. Template is READ-ONLY — design changes only in PowerPoint
 */
@Service
public class TemplateRenderService {

    private static final Logger log = LoggerFactory.getLogger(TemplateRenderService.class);

    @Value("classpath:templates/medaccur_master.pptx")
    private Resource masterTemplateResource;

    private final PlaceholderService placeholderService;
    private final FigureService figureService;
    private final ChartDataService chartDataService;

    private byte[] masterTemplateBytes;
    private List<String> availableLayouts = new ArrayList<>();

    public TemplateRenderService(PlaceholderService placeholderService,
                                  FigureService figureService,
                                  ChartDataService chartDataService) {
        this.placeholderService = placeholderService;
        this.figureService = figureService;
        this.chartDataService = chartDataService;
    }

    @PostConstruct
    public void init() {
        try {
            ZipSecureFile.setMaxFileCount(10000);
            ZipSecureFile.setMinInflateRatio(0.001);

            masterTemplateBytes = masterTemplateResource.getInputStream().readAllBytes();

            // Index all layout names
            try (XMLSlideShow pptx = new XMLSlideShow(new ByteArrayInputStream(masterTemplateBytes))) {
                for (XSLFSlideMaster master : pptx.getSlideMasters()) {
                    for (XSLFSlideLayout layout : master.getSlideLayouts()) {
                        String name = layout.getName();
                        if (name != null && !name.isEmpty()) {
                            availableLayouts.add(name);
                            log.info("  Layout: {}", name);
                        }
                    }
                }
            }
            log.info("Master template loaded: {} bytes, {} layouts", 
                     masterTemplateBytes.length, availableLayouts.size());

        } catch (Exception e) {
            log.warn("Master template not loaded: {}", e.getMessage());
        }
    }

    /**
     * Render a complete deck from a RenderRequest.
     * 
     * Strategy: Open the master template as our working copy.
     * For each slide in the request, create a new slide from the matching layout.
     * Then remove the original blank slide(s) that came with the template.
     */
    public byte[] renderDeck(RenderRequest request) throws Exception {
        if (masterTemplateBytes == null) {
            throw new IllegalStateException("Master template not loaded");
        }

        // Open a fresh copy of the master template as our output
        XMLSlideShow pptx = new XMLSlideShow(new ByteArrayInputStream(masterTemplateBytes));

        // Remember how many slides the template originally has (to delete later)
        int originalSlideCount = pptx.getSlides().size();

        // Create slides from requested layouts
        for (int i = 0; i < request.getSlides().size(); i++) {
            SlideSpec spec = request.getSlides().get(i);
            String layoutName = spec.getLayout();
            log.info("Slide {}/{}: layout='{}'", i + 1, request.getSlides().size(), layoutName);

            try {
                XSLFSlideLayout layout = findLayout(pptx, layoutName);
                if (layout == null) {
                    log.warn("  Layout '{}' not found — skipping", layoutName);
                    continue;
                }

                // Create new slide from layout — preserves ALL design from template
                XSLFSlide slide = pptx.createSlide(layout);

                // Merge metadata + slide content
                Map<String, String> merged = new HashMap<>();
                if (request.getMetadata() != null) merged.putAll(request.getMetadata());
                if (spec.getContent() != null) merged.putAll(spec.getContent());

                // Step 1: Replace {{placeholders}}
                int replaced = placeholderService.replacePlaceholders(slide, merged);
                log.info("  {} placeholders replaced", replaced);

                // Step 2: Remove unfilled {{placeholders}}
                placeholderService.removeUnfilledPlaceholders(slide);

                // Step 3: Update embedded chart data (Excel + XML cache)
                if (spec.hasChartData()) {
                    chartDataService.updateAllCharts(slide, spec.getChartData());
                }

                // Step 4: Insert figure PNGs (KM curves, Forest plots)
                if (spec.hasFigure()) {
                    figureService.insertFigure(slide, spec.getFigurePng());
                }

                // Step 5: Enable auto-shrink
                placeholderService.enableAutoShrink(slide);

            } catch (Exception e) {
                log.error("  Slide {} failed: {}", i + 1, e.getMessage(), e);
            }
        }

        // Remove original template slides (they were just carriers for layouts)
        for (int i = originalSlideCount - 1; i >= 0; i--) {
            pptx.removeSlide(i);
        }

        // Write output
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        pptx.write(out);
        pptx.close();

        log.info("Deck rendered: {} slides, {} KB", 
                 request.getSlides().size(), out.size() / 1024);
        return out.toByteArray();
    }

    /**
     * Find a SlideLayout by name across all SlideMasters.
     * Case-insensitive match.
     */
    private XSLFSlideLayout findLayout(XMLSlideShow pptx, String layoutName) {
        for (XSLFSlideMaster master : pptx.getSlideMasters()) {
            for (XSLFSlideLayout layout : master.getSlideLayouts()) {
                if (layout.getName() != null && 
                    layout.getName().equalsIgnoreCase(layoutName)) {
                    return layout;
                }
            }
        }
        return null;
    }

    public boolean isTemplateLoaded() {
        return masterTemplateBytes != null;
    }

    public int getLayoutCount() {
        return availableLayouts.size();
    }

    public List<String> getAvailableLayouts() {
        return Collections.unmodifiableList(availableLayouts);
    }
}
