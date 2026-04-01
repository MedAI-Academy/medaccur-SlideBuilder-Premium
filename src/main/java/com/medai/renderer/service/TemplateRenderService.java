package com.medai.renderer.service;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.medai.renderer.model.RenderRequest;
import com.medai.renderer.model.SlideData;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xslf.usermodel.*;
import org.openxmlformats.schemas.presentationml.x2006.main.CTSlide;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Service;

import jakarta.annotation.PostConstruct;
import java.io.*;
import java.nio.file.*;
import java.util.*;

/**
 * Template-based PPTX renderer — Clone+Swap approach.
 *
 * Flow:
 * 1. Receives a "deck recipe" — ordered list of {layout, placeholders}
 * 2. For each entry: opens the correct template PPTX, clones the correct slide
 * 3. Replaces all {{placeholder}} text runs (preserving formatting)
 * 4. Returns the assembled PPTX
 *
 * Templates are stored in classpath:templates/medaccur/
 * Manifest: classpath:templates/medaccur/medaccur_manifest.json
 */
@Service
public class TemplateRenderService {

    private static final Logger log = LoggerFactory.getLogger(TemplateRenderService.class);
    private static final ObjectMapper mapper = new ObjectMapper();
    private static final String TEMPLATE_DIR = "templates/medaccur/";

    @Autowired(required = false)
    private ChartService chartService;

    // layout_id → {file, slide} from manifest
    private Map<String, LayoutRef> layoutMap;
    // section_id → layout_id from manifest
    private Map<String, String> sectionMap;
    // auto-select rules
    private Map<String, Map<String, String>> autoSelectRules;

    private record LayoutRef(String file, int slide) {}

    @PostConstruct
    @SuppressWarnings("unchecked")
    public void loadManifest() {
        try {
            InputStream is = new ClassPathResource(TEMPLATE_DIR + "medaccur_manifest.json").getInputStream();
            Map<String, Object> manifest = mapper.readValue(is, Map.class);

            // Parse layout_map
            layoutMap = new LinkedHashMap<>();
            Map<String, Map<String, Object>> layouts = (Map<String, Map<String, Object>>) manifest.get("layout_map");
            for (var entry : layouts.entrySet()) {
                String layoutId = entry.getKey();
                String file = (String) entry.getValue().get("file");
                int slideNum = ((Number) entry.getValue().get("slide")).intValue();
                layoutMap.put(layoutId, new LayoutRef(file, slideNum));
            }

            // Parse section_to_layout
            sectionMap = new LinkedHashMap<>();
            Map<String, String> sections = (Map<String, String>) manifest.get("section_to_layout");
            sectionMap.putAll(sections);

            // Parse auto_select_rules
            autoSelectRules = new LinkedHashMap<>();
            Map<String, Map<String, String>> rules = (Map<String, Map<String, String>>) manifest.get("auto_select_rules");
            if (rules != null) autoSelectRules.putAll(rules);

            log.info("medaccur manifest loaded: {} layouts, {} section mappings", layoutMap.size(), sectionMap.size());
        } catch (Exception e) {
            log.error("Failed to load medaccur manifest", e);
            layoutMap = Map.of();
            sectionMap = Map.of();
            autoSelectRules = Map.of();
        }
    }

    /**
     * Render a complete deck from a recipe.
     *
     * @param request RenderRequest with slides containing section IDs and placeholder data
     * @return PPTX bytes
     */
    public byte[] renderDeck(RenderRequest request) throws Exception {
        List<SlideData> recipe = request.getSlides();
        if (recipe == null || recipe.isEmpty()) {
            throw new IllegalArgumentException("No slides in recipe");
        }

        log.info("Rendering deck: {} slides, drug={}", recipe.size(), request.meta("drug"));

        // Create output presentation (empty, widescreen)
        XMLSlideShow output = new XMLSlideShow();
        output.setPageSize(new java.awt.Dimension(960, 540)); // 13.33" x 7.50" at 72dpi

        for (int i = 0; i < recipe.size(); i++) {
            SlideData sd = recipe.get(i);
            String sectionId = sd.getId() != null ? sd.getId() : "";
            Map<String, Object> placeholders = sd.getContent() != null ? sd.getContent() : Map.of();

            // Resolve layout ID
            String layoutId = resolveLayout(sectionId, sd.getLayout(), placeholders);
            LayoutRef ref = layoutMap.get(layoutId);
            if (ref == null) {
                log.warn("Unknown layout '{}' for section '{}' — skipping", layoutId, sectionId);
                continue;
            }

            log.debug("Slide {}: section={}, layout={}, file={}, slide={}",
                    i + 1, sectionId, layoutId, ref.file, ref.slide);

            // Clone slide from template into output
            cloneSlideFromTemplate(output, ref.file, ref.slide);

            // Replace placeholders in the last added slide
            XSLFSlide outputSlide = output.getSlides().get(output.getSlides().size() - 1);
            replacePlaceholders(outputSlide, placeholders);

            // Add global metadata as placeholders
            Map<String, Object> globalPlaceholders = new LinkedHashMap<>();
            globalPlaceholders.put("drug_label", request.meta("drug") + " · " + request.meta("indication"));
            globalPlaceholders.put("phase_badge", request.meta("phase"));
            globalPlaceholders.put("scope_badge", request.meta("country") + " · " + request.meta("year"));
            globalPlaceholders.put("confidence", request.meta("confidence"));
            replacePlaceholders(outputSlide, globalPlaceholders);
        }

        // Write output
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        output.write(baos);
        output.close();

        byte[] result = baos.toByteArray();
        log.info("Deck rendered: {} slides, {}KB", recipe.size(), result.length / 1024);
        return result;
    }

    /**
     * Resolve the layout ID for a section, handling auto-select rules.
     */
    private String resolveLayout(String sectionId, String explicitLayout, Map<String, Object> data) {
        // If explicit layout provided, use it
        if (explicitLayout != null && !explicitLayout.isEmpty() && layoutMap.containsKey(explicitLayout)) {
            return explicitLayout;
        }

        // Look up section → layout mapping
        String mapped = sectionMap.getOrDefault(sectionId, "");

        // Handle auto-select
        if (mapped.startsWith("_AUTO_SELECT_")) {
            return autoSelect(sectionId, data);
        }

        return mapped.isEmpty() ? "DISEASE_OVERVIEW" : mapped; // fallback
    }

    /**
     * Auto-select layout based on data content.
     */
    @SuppressWarnings("unchecked")
    private String autoSelect(String sectionId, Map<String, Object> data) {
        Map<String, String> rules = autoSelectRules.getOrDefault(sectionId, Map.of());

        if ("strategic_imperatives".equals(sectionId)) {
            // Count pillars
            List<?> pillars = data.get("pillars") instanceof List ? (List<?>) data.get("pillars") : List.of();
            int n = pillars.size();
            if (n <= 3) return rules.getOrDefault("3_pillars", "IMPERATIVES_TEMPLE_3");
            if (n == 4) return rules.getOrDefault("4_pillars", "IMPERATIVES_TEMPLE_4");
            return rules.getOrDefault("5_pillars", "IMPERATIVES_TEMPLE_5");
        }

        if ("tactical_plan".equals(sectionId)) {
            List<?> rows = data.get("workstreams") instanceof List ? (List<?>) data.get("workstreams") : List.of();
            if (rows.size() <= 5) return rules.getOrDefault("5_or_fewer", "TACTICAL_PLAN_COMPACT");
            return rules.getOrDefault("6_or_more", "TACTICAL_PLAN_FULL");
        }

        if ("swot".equals(sectionId)) {
            return rules.getOrDefault("default", "SWOT_CLASSIC");
        }

        return rules.getOrDefault("fallback", "DISEASE_OVERVIEW");
    }

    /**
     * Clone a slide from a template file into the output presentation.
     * Uses XML-level deep copy to preserve all shapes, formatting, backgrounds.
     */
    private void cloneSlideFromTemplate(XMLSlideShow output, String templateFile, int slideNumber)
            throws Exception {
        // Load template
        InputStream tplStream = new ClassPathResource(TEMPLATE_DIR + templateFile).getInputStream();
        XMLSlideShow template = new XMLSlideShow(tplStream);

        List<XSLFSlide> templateSlides = template.getSlides();
        int idx = slideNumber - 1; // manifest uses 1-indexed

        if (idx < 0 || idx >= templateSlides.size()) {
            template.close();
            throw new IllegalArgumentException(
                    "Template '" + templateFile + "' has " + templateSlides.size() +
                    " slides but requested slide " + slideNumber);
        }

        XSLFSlide sourceSlide = templateSlides.get(idx);

        // Import the slide layout first
        XSLFSlideLayout sourceLayout = sourceSlide.getSlideLayout();
        XSLFSlide newSlide = output.createSlide();

        // Deep copy via XML
        CTSlide srcCt = sourceSlide.getXmlObject();
        CTSlide dstCt = newSlide.getXmlObject();
        dstCt.set(srcCt.copy());

        // Copy relationships (images, embedded objects)
        for (var rel : sourceSlide.getRelations()) {
            if (rel instanceof XSLFPictureData picData) {
                // Re-add picture to output
                XSLFPictureData newPic = output.addPicture(
                        picData.getData(), picData.getType());
                // Update relationship
                newSlide.getPackagePart().addRelationship(
                        newPic.getPackagePart().getPartName(),
                        org.apache.poi.openxml4j.opc.TargetMode.INTERNAL,
                        rel.getPackageRelationship().getRelationshipType(),
                        rel.getPackageRelationship().getId());
            }
        }

        template.close();
    }

    /**
     * Replace all {{placeholder}} text in all shapes of a slide.
     * Preserves run-level formatting (font, size, color, bold).
     */
    private void replacePlaceholders(XSLFSlide slide, Map<String, Object> data) {
        if (data == null || data.isEmpty()) return;

        for (XSLFShape shape : slide.getShapes()) {
            if (shape instanceof XSLFTextShape textShape) {
                replaceInTextShape(textShape, data);
            } else if (shape instanceof XSLFGroupShape group) {
                // Recurse into grouped shapes
                for (XSLFShape child : group.getShapes()) {
                    if (child instanceof XSLFTextShape childText) {
                        replaceInTextShape(childText, data);
                    }
                }
            }
        }
    }

    private void replaceInTextShape(XSLFTextShape textShape, Map<String, Object> data) {
        for (XSLFTextParagraph para : textShape.getTextParagraphs()) {
            for (XSLFTextRun run : para.getTextRuns()) {
                String text = run.getRawText();
                if (text == null || !text.contains("{{")) continue;

                for (var entry : data.entrySet()) {
                    String placeholder = "{{" + entry.getKey() + "}}";
                    if (text.contains(placeholder)) {
                        String value = entry.getValue() != null ? entry.getValue().toString() : "";
                        text = text.replace(placeholder, value);
                    }
                }
                run.setText(text);
            }
        }
    }
}
