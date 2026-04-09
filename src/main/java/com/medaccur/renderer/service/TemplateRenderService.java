package com.medaccur.renderer.service;

import com.medaccur.renderer.model.RenderRequest;
import com.medaccur.renderer.model.SlideSpec;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.openxml4j.opc.TargetMode;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.xslf.usermodel.*;
import org.openxmlformats.schemas.presentationml.x2006.main.CTGroupShape;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.io.Resource;
import org.springframework.stereotype.Service;

import jakarta.annotation.PostConstruct;
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
 *
 * KEY FIX (v3.1): pptx.createSlide(layout) creates an EMPTY slide.
 * Layout shapes (containing {{placeholders}}) are NOT copied to the slide.
 * We must import them via XML deep-copy BEFORE placeholder replacement.
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
     * Strategy:
     * PASS 1 — Create all slides from layouts and import layout shapes to slide level.
     * CLEANUP — Clear shapes from used layouts (prevents visual duplication).
     * PASS 2 — Replace placeholders, update charts, insert figures on each slide.
     * FINAL  — Remove original template slides and write output.
     */
    public byte[] renderDeck(RenderRequest request) throws Exception {
        if (masterTemplateBytes == null) {
            throw new IllegalStateException("Master template not loaded");
        }

        XMLSlideShow pptx = new XMLSlideShow(new ByteArrayInputStream(masterTemplateBytes));
        int originalSlideCount = pptx.getSlides().size();

        // ── PASS 1: Create slides and import layout shapes ──────────
        Set<XSLFSlideLayout> usedLayouts = new LinkedHashSet<>();
        List<XSLFSlide> newSlides = new ArrayList<>();
        List<SlideSpec> matchedSpecs = new ArrayList<>();

        for (int i = 0; i < request.getSlides().size(); i++) {
            SlideSpec spec = request.getSlides().get(i);
            String layoutName = spec.getLayout();
            log.info("Slide {}/{}: layout='{}'", i + 1, request.getSlides().size(), layoutName);

            XSLFSlideLayout layout = findLayout(pptx, layoutName);
            if (layout == null) {
                log.warn("  Layout '{}' not found — skipping", layoutName);
                continue;
            }

            // Create slide from layout (inherits design but shapes are empty)
            XSLFSlide slide = pptx.createSlide(layout);

            // CRITICAL FIX: Copy all shapes from layout to slide level via XML
            int imported = importLayoutShapesToSlide(slide, layout);
            log.info("  {} shapes imported from layout to slide", imported);

            usedLayouts.add(layout);
            newSlides.add(slide);
            matchedSpecs.add(spec);
        }

        // ── CLEANUP: Clear shapes from used layouts (prevent double-rendering) ──
        for (XSLFSlideLayout layout : usedLayouts) {
            clearLayoutShapes(layout);
            log.debug("  Cleared shapes from layout '{}'", layout.getName());
        }

        // ── PASS 2: Fill content on each slide ─────────────────────
        for (int i = 0; i < newSlides.size(); i++) {
            XSLFSlide slide = newSlides.get(i);
            SlideSpec spec = matchedSpecs.get(i);

            // Merge metadata + slide-specific content
            Map<String, String> merged = new HashMap<>();
            if (request.getMetadata() != null) merged.putAll(request.getMetadata());
            if (spec.getContent() != null) merged.putAll(spec.getContent());

            // Step 1: Replace {{placeholders}}
            int replaced = placeholderService.replacePlaceholders(slide, merged);
            log.info("  Slide {}: {} placeholders replaced", i + 1, replaced);

            // Step 2: Remove unfilled {{placeholders}}
            int cleaned = placeholderService.removeUnfilledPlaceholders(slide);
            if (cleaned > 0) log.info("  Slide {}: {} unfilled placeholders removed", i + 1, cleaned);

            // Step 3: Update embedded chart data
            if (spec.hasChartData()) {
                chartDataService.updateAllCharts(slide, spec.getChartData());
            }

            // Step 4: Insert figure PNGs (KM curves, Forest plots)
            if (spec.hasFigure()) {
                figureService.insertFigure(slide, spec.getFigurePng());
            }

            // Step 5: Enable auto-shrink
            placeholderService.enableAutoShrink(slide);
        }

        // ── FINAL: Remove original template slides ──────────────────
        for (int i = originalSlideCount - 1; i >= 0; i--) {
            pptx.removeSlide(i);
        }

        // Write output
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        pptx.write(out);
        pptx.close();

        log.info("Deck rendered: {} slides, {} KB",
                 newSlides.size(), out.size() / 1024);
        return out.toByteArray();
    }

    // ════════════════════════════════════════════════════════════
    // LAYOUT → SLIDE SHAPE IMPORT (the critical fix)
    // ════════════════════════════════════════════════════════════

    /**
     * Deep-copy all shapes from a layout's XML spTree to the slide's XML spTree.
     *
     * When Apache POI creates a slide via createSlide(layout), the slide is EMPTY.
     * Layout shapes are only "inherited" visually — they're not on the slide and
     * therefore NOT accessible via slide.getShapes(). This means PlaceholderService
     * can never find or replace {{placeholders}}, and ChartDataService can never
     * find embedded charts to update their Excel data.
     *
     * This method copies ALL shape types:
     * - sp       (TextBoxes, AutoShapes — contain {{placeholders}})
     * - grpSp    (Group shapes — may contain nested text)
     * - cxnSp    (Connectors)
     * - graphicFrame (Embedded Excel charts — need relationship copy!)
     * - pic      (Embedded images — need relationship copy!)
     *
     * For graphicFrame and pic, we also copy the OPC relationships from the
     * layout part to the slide part so that chart/image references resolve.
     *
     * @return number of shapes imported
     */
    private int importLayoutShapesToSlide(XSLFSlide slide, XSLFSlideLayout layout) {
        int count = 0;
        try {
            CTGroupShape layoutTree = layout.getXmlObject().getCSld().getSpTree();
            CTGroupShape slideTree = slide.getXmlObject().getCSld().getSpTree();

            // Copy sp elements (TextBoxes, AutoShapes with {{placeholders}})
            for (int i = 0; i < layoutTree.sizeOfSpArray(); i++) {
                slideTree.addNewSp().set(layoutTree.getSpArray(i).copy());
                count++;
            }

            // Copy group shapes (may contain nested TextBoxes)
            for (int i = 0; i < layoutTree.sizeOfGrpSpArray(); i++) {
                slideTree.addNewGrpSp().set(layoutTree.getGrpSpArray(i).copy());
                count++;
            }

            // Copy connector shapes
            for (int i = 0; i < layoutTree.sizeOfCxnSpArray(); i++) {
                slideTree.addNewCxnSp().set(layoutTree.getCxnSpArray(i).copy());
                count++;
            }

            // Copy graphicFrame elements (embedded Excel charts, tables)
            for (int i = 0; i < layoutTree.sizeOfGraphicFrameArray(); i++) {
                slideTree.addNewGraphicFrame().set(layoutTree.getGraphicFrameArray(i).copy());
                count++;
            }

            // Copy pic elements (embedded images, logos)
            for (int i = 0; i < layoutTree.sizeOfPicArray(); i++) {
                slideTree.addNewPic().set(layoutTree.getPicArray(i).copy());
                count++;
            }

            // Copy ALL relationships from layout to slide (charts, images, etc.)
            // GraphicFrame/pic XML contains r:id="rIdX" references that must resolve
            // on the slide part, not just the layout part.
            copyRelationships(layout, slide);

        } catch (Exception e) {
            log.error("Shape import failed: {}", e.getMessage(), e);
        }
        return count;
    }

    /**
     * Copy OPC package relationships from layout to slide.
     * This ensures that r:id references in copied graphicFrame/pic XML
     * resolve correctly on the slide part.
     *
     * Copies: chart relations, image relations, and any other embedded parts.
     */
    private void copyRelationships(XSLFSlideLayout layout, XSLFSlide slide) {
        try {
            for (POIXMLDocumentPart.RelationPart rp : layout.getRelationParts()) {
                POIXMLDocumentPart part = rp.getDocumentPart();
                PackageRelationship rel = rp.getRelationship();
                String relId = rel.getId();

                // Skip if this relationship already exists on the slide
                try {
                    if (slide.getPackagePart().getRelationship(relId) != null) continue;
                } catch (Exception ignore) {}

                // Copy the relationship: same ID, same target, same type
                slide.getPackagePart().addRelationship(
                    part.getPackagePart().getPartName(),
                    TargetMode.INTERNAL,
                    rel.getRelationshipType(),
                    relId
                );
                log.debug("  Copied relationship {} → {} ({})",
                    relId, part.getPackagePart().getPartName(), 
                    part.getClass().getSimpleName());
            }
        } catch (Exception e) {
            log.warn("Relationship copy partial failure: {}", e.getMessage());
        }
    }

    /**
     * Clear all shapes from a layout's spTree after they've been copied to slides.
     * This prevents visual duplication (layout shapes showing behind slide shapes).
     *
     * SAFE because we work on a FRESH COPY of the template per render.
     * The original template file is never modified.
     */
    private void clearLayoutShapes(XSLFSlideLayout layout) {
        try {
            CTGroupShape tree = layout.getXmlObject().getCSld().getSpTree();

            while (tree.sizeOfSpArray() > 0) tree.removeSp(0);
            while (tree.sizeOfGrpSpArray() > 0) tree.removeGrpSp(0);
            while (tree.sizeOfCxnSpArray() > 0) tree.removeCxnSp(0);
            while (tree.sizeOfGraphicFrameArray() > 0) tree.removeGraphicFrame(0);
            while (tree.sizeOfPicArray() > 0) tree.removePic(0);

        } catch (Exception e) {
            log.error("Layout shape clear failed: {}", e.getMessage(), e);
        }
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
