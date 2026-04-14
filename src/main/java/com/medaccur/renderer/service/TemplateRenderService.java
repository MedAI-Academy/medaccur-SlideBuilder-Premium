package com.medaccur.renderer.service;

import com.medaccur.renderer.model.RenderRequest;
import com.medaccur.renderer.model.SlideSpec;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.openxml4j.opc.PackagePart;
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
 * v3.7: Added ForestPlotShapeService for native forest plot rendering.
 *
 * Charts updated on layout -> then ALL shapes (including charts)
 * copied to slide with relationships -> layout cleared.
 */
@Service
public class TemplateRenderService {

    private static final Logger log = LoggerFactory.getLogger(TemplateRenderService.class);

    @Value("classpath:templates/medaccur_master.pptx")
    private Resource masterTemplateResource;

    private final PlaceholderService placeholderService;
    private final FigureService figureService;
    private final ChartDataService chartDataService;
    private final ForestPlotShapeService forestPlotShapeService;  // NEW

    private byte[] masterTemplateBytes;
    private List<String> availableLayouts = new ArrayList<>();

    public TemplateRenderService(PlaceholderService placeholderService,
                                  FigureService figureService,
                                  ChartDataService chartDataService,
                                  ForestPlotShapeService forestPlotShapeService) {  // NEW
        this.placeholderService = placeholderService;
        this.figureService = figureService;
        this.chartDataService = chartDataService;
        this.forestPlotShapeService = forestPlotShapeService;  // NEW
    }

    @PostConstruct
    public void init() {
        try {
            ZipSecureFile.setMaxFileCount(10000);
            ZipSecureFile.setMinInflateRatio(0.001);

            masterTemplateBytes = masterTemplateResource.getInputStream().readAllBytes();

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

    public byte[] renderDeck(RenderRequest request) throws Exception {
        if (masterTemplateBytes == null) {
            throw new IllegalStateException("Master template not loaded");
        }

        XMLSlideShow pptx = new XMLSlideShow(new ByteArrayInputStream(masterTemplateBytes));
        int originalSlideCount = pptx.getSlides().size();

        Set<XSLFSlideLayout> usedLayouts = new LinkedHashSet<>();
        List<XSLFSlide> newSlides = new ArrayList<>();
        List<SlideSpec> matchedSpecs = new ArrayList<>();

        for (int i = 0; i < request.getSlides().size(); i++) {
            SlideSpec spec = request.getSlides().get(i);
            String layoutName = spec.getLayout();
            log.info("Slide {}/{}: layout='{}'", i + 1, request.getSlides().size(), layoutName);

            XSLFSlideLayout layout = findLayout(pptx, layoutName);
            if (layout == null) {
                log.warn("  Layout '{}' not found -- skipping", layoutName);
                continue;
            }

            // Step A: Update chart data on layout BEFORE creating slide
            if (spec.hasChartData()) {
                chartDataService.updateAllCharts(layout, spec.getChartData());
            }

            // Step B: Create slide from layout
            XSLFSlide slide = pptx.createSlide(layout);

            // Step C: Copy ALL shapes from layout to slide (text + charts + images)
            int imported = importAllShapes(slide, layout);
            log.info("  {} shapes imported to slide", imported);

            // Step D: Copy chart/image relationships from layout to slide
            copyRelationships(layout, slide);

            usedLayouts.add(layout);
            newSlides.add(slide);
            matchedSpecs.add(spec);
        }

        // Clear shapes from used layouts (prevent double-rendering)
        for (XSLFSlideLayout layout : usedLayouts) {
            clearAllLayoutShapes(layout);
        }

        // Fill content on each slide
        for (int i = 0; i < newSlides.size(); i++) {
            XSLFSlide slide = newSlides.get(i);
            SlideSpec spec = matchedSpecs.get(i);

            Map<String, String> merged = new HashMap<>();
            if (request.getMetadata() != null) merged.putAll(request.getMetadata());
            if (spec.getContent() != null) merged.putAll(spec.getContent());

            int replaced = placeholderService.replacePlaceholders(slide, merged);
            log.info("  Slide {}: {} placeholders replaced", i + 1, replaced);

            placeholderService.removeUnfilledPlaceholders(slide);

            if (spec.hasFigure()) {
                figureService.insertFigure(slide, spec.getFigurePng());
            }

            // NEW: Draw native forest plot shapes
            if (spec.hasForestPlot()) {
                forestPlotShapeService.addForestPlot(slide, spec.getForestPlotPoints());
                log.info("  Slide {}: forest plot shapes drawn", i + 1);
            }

            placeholderService.enableAutoShrink(slide);
        }

        // Remove original template slides
        for (int i = originalSlideCount - 1; i >= 0; i--) {
            pptx.removeSlide(i);
        }

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        pptx.write(out);
        pptx.close();

        log.info("Deck rendered: {} slides, {} KB", newSlides.size(), out.size() / 1024);
        return out.toByteArray();
    }

    // ════════════════════════════════════════════════
    // SHAPE IMPORT: ALL types including charts
    // ════════════════════════════════════════════════

    private int importAllShapes(XSLFSlide slide, XSLFSlideLayout layout) {
        int count = 0;
        try {
            CTGroupShape src = layout.getXmlObject().getCSld().getSpTree();
            CTGroupShape dst = slide.getXmlObject().getCSld().getSpTree();

            for (int i = 0; i < src.sizeOfSpArray(); i++) {
                dst.addNewSp().set(src.getSpArray(i).copy());
                count++;
            }
            for (int i = 0; i < src.sizeOfGrpSpArray(); i++) {
                dst.addNewGrpSp().set(src.getGrpSpArray(i).copy());
                count++;
            }
            for (int i = 0; i < src.sizeOfCxnSpArray(); i++) {
                dst.addNewCxnSp().set(src.getCxnSpArray(i).copy());
                count++;
            }
            for (int i = 0; i < src.sizeOfGraphicFrameArray(); i++) {
                dst.addNewGraphicFrame().set(src.getGraphicFrameArray(i).copy());
                count++;
                log.debug("  Copied graphicFrame (chart/table)");
            }
            for (int i = 0; i < src.sizeOfPicArray(); i++) {
                dst.addNewPic().set(src.getPicArray(i).copy());
                count++;
            }
        } catch (Exception e) {
            log.error("Shape import failed: {}", e.getMessage(), e);
        }
        return count;
    }

    private void copyRelationships(XSLFSlideLayout layout, XSLFSlide slide) {
        try {
            PackagePart slidePart = slide.getPackagePart();

            for (POIXMLDocumentPart.RelationPart rp : layout.getRelationParts()) {
                PackageRelationship rel = rp.getRelationship();
                String relId = rel.getId();
                String relType = rel.getRelationshipType();

                if (!relType.contains("chart") && !relType.contains("image")) continue;

                boolean exists = false;
                try {
                    if (slidePart.getRelationship(relId) != null) exists = true;
                } catch (Exception ignore) {}

                if (!exists) {
                    PackagePart targetPart = rp.getDocumentPart().getPackagePart();
                    slidePart.addRelationship(
                        targetPart.getPartName(),
                        TargetMode.INTERNAL,
                        relType,
                        relId
                    );
                    log.debug("  Rel copied: {} -> {}", relId, targetPart.getPartName());
                }
            }
        } catch (Exception e) {
            log.warn("Relationship copy issue: {}", e.getMessage());
        }
    }

    private void clearAllLayoutShapes(XSLFSlideLayout layout) {
        try {
            CTGroupShape tree = layout.getXmlObject().getCSld().getSpTree();
            while (tree.sizeOfSpArray() > 0) tree.removeSp(0);
            while (tree.sizeOfGrpSpArray() > 0) tree.removeGrpSp(0);
            while (tree.sizeOfCxnSpArray() > 0) tree.removeCxnSp(0);
            while (tree.sizeOfGraphicFrameArray() > 0) tree.removeGraphicFrame(0);
            while (tree.sizeOfPicArray() > 0) tree.removePic(0);
        } catch (Exception e) {
            log.error("Layout clear failed: {}", e.getMessage(), e);
        }
    }

    private XSLFSlideLayout findLayout(XMLSlideShow pptx, String layoutName) {
        for (XSLFSlideMaster master : pptx.getSlideMasters()) {
            for (XSLFSlideLayout layout : master.getSlideLayouts()) {
                if (layout.getName() != null && layout.getName().equalsIgnoreCase(layoutName)) {
                    return layout;
                }
            }
        }
        return null;
    }

    public boolean isTemplateLoaded() { return masterTemplateBytes != null; }
    public int getLayoutCount() { return availableLayouts.size(); }
    public List<String> getAvailableLayouts() { return Collections.unmodifiableList(availableLayouts); }
}
