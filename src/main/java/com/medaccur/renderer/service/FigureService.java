package com.medaccur.renderer.service;

import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.xslf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import java.awt.geom.Rectangle2D;
import java.util.Base64;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Inserts PNG figures (KM curves, Forest plots) into CHART_AREA shapes.
 * Finds shapes by alt-text containing "CHART_AREA".
 * Preserves aspect ratio. Removes the placeholder shape after insertion.
 */
@Service
public class FigureService {

    private static final Logger log = LoggerFactory.getLogger(FigureService.class);
    private static final Pattern ALT_TEXT_PATTERN = Pattern.compile("descr=\"([^\"]*)\"");

    /**
     * Insert a base64-encoded PNG into the first CHART_AREA shape on the slide
     * that does NOT contain an embedded chart (i.e., is a pure placeholder rectangle).
     */
    public void insertFigure(XSLFSlide slide, String base64Png) {
        if (base64Png == null || base64Png.isEmpty()) return;

        // Find CHART_AREA shape that is NOT an embedded chart
        XSLFShape targetShape = null;
        for (XSLFShape shape : slide.getShapes()) {
            if (shape instanceof XSLFGraphicFrame) continue; // skip embedded charts
            String altText = extractAltText(shape);
            if ("CHART_AREA".equalsIgnoreCase(altText)) {
                targetShape = shape;
                break;
            }
        }

        if (targetShape == null) {
            log.warn("No CHART_AREA placeholder found (non-chart) — figure not inserted");
            return;
        }

        try {
            Rectangle2D anchor = targetShape.getAnchor();
            byte[] imageBytes = Base64.getDecoder().decode(base64Png);

            XMLSlideShow pptx = (XMLSlideShow) slide.getSlideShow();
            XSLFPictureData picData = pptx.addPicture(imageBytes, PictureData.PictureType.PNG);
            XSLFPictureShape pic = slide.createPicture(picData);
            pic.setAnchor(anchor); // exact same position as placeholder

            // Remove the placeholder rectangle
            slide.removeShape(targetShape);

            log.info("Figure inserted: {:.0f}x{:.0f} at ({:.1f},{:.1f}) — {} KB",
                     anchor.getWidth(), anchor.getHeight(),
                     anchor.getX(), anchor.getY(),
                     imageBytes.length / 1024);

        } catch (Exception e) {
            log.error("Figure insertion failed: {}", e.getMessage());
        }
    }

    /**
     * Insert a figure into a SPECIFIC named CHART_AREA.
     * Used when a slide has multiple chart areas (e.g., KM + Forest on same slide).
     */
    public void insertFigureByName(XSLFSlide slide, String areaName, String base64Png) {
        if (base64Png == null || base64Png.isEmpty()) return;

        XSLFShape targetShape = null;
        for (XSLFShape shape : slide.getShapes()) {
            if (shape instanceof XSLFGraphicFrame) continue;
            String altText = extractAltText(shape);
            if (areaName.equalsIgnoreCase(altText)) {
                targetShape = shape;
                break;
            }
        }

        if (targetShape == null) {
            log.warn("Named area '{}' not found — figure not inserted", areaName);
            return;
        }

        try {
            Rectangle2D anchor = targetShape.getAnchor();
            byte[] imageBytes = Base64.getDecoder().decode(base64Png);

            XMLSlideShow pptx = (XMLSlideShow) slide.getSlideShow();
            XSLFPictureData picData = pptx.addPicture(imageBytes, PictureData.PictureType.PNG);
            XSLFPictureShape pic = slide.createPicture(picData);
            pic.setAnchor(anchor);

            slide.removeShape(targetShape);
            log.info("Figure '{}' inserted — {} KB", areaName, imageBytes.length / 1024);

        } catch (Exception e) {
            log.error("Figure '{}' insertion failed: {}", areaName, e.getMessage());
        }
    }

    private String extractAltText(XSLFShape shape) {
        try {
            String xml = shape.getXmlObject().toString();
            Matcher m = ALT_TEXT_PATTERN.matcher(xml);
            if (m.find()) return m.group(1);
        } catch (Exception e) { /* ignore */ }
        return "";
    }
}
