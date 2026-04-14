package com.medaccur.renderer.service;

import com.medaccur.renderer.model.ForestPlotPoint;
import org.apache.poi.xslf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import java.awt.Color;
import java.awt.geom.Rectangle2D;
import java.util.List;

/**
 * Draws native PowerPoint shapes for forest plots: diamonds, squares, CI whiskers,
 * and a vertical reference line at HR=1.0.
 *
 * Shapes are aligned with the FORESTPLOT_light template table rows.
 * This replaces the matplotlib PNG approach for perfect alignment.
 *
 * Coordinate system: points (1 inch = 72 points).
 * Slide dimensions: 13.33" x 7.5" (widescreen 16:9).
 */
@Service
public class ForestPlotShapeService {

    private static final Logger log = LoggerFactory.getLogger(ForestPlotShapeService.class);

    // ═══════════════════════════════════════════════════
    // GRID CONSTANTS — must match FORESTPLOT_light layout
    // ═══════════════════════════════════════════════════

    // Plot area (right side of slide, next to the table)
    private static final double PLOT_LEFT   = 5.80;   // inches from left edge
    private static final double PLOT_RIGHT  = 12.40;  // inches from left edge
    private static final double PLOT_WIDTH  = PLOT_RIGHT - PLOT_LEFT;  // ~6.6 inches

    // Vertical positioning of rows
    private static final double FIRST_ROW_Y = 1.15;   // inches: center of "Overall" row
    private static final double ROW_HEIGHT  = 0.30;   // inches per row

    // Log scale range
    private static final double LOG_MIN = 0.05;   // left edge of plot = HR 0.05
    private static final double LOG_MAX = 5.00;   // right edge of plot = HR 5.0
    private static final double LOG_RANGE = Math.log10(LOG_MAX) - Math.log10(LOG_MIN);

    // Shape sizes
    private static final double SQUARE_SIZE    = 0.12;  // inches (side of subgroup square)
    private static final double DIAMOND_HEIGHT = 0.20;  // inches (half-height of overall diamond)
    private static final double DIAMOND_WIDTH  = 0.15;  // inches (half-width of overall diamond)
    private static final double WHISKER_WIDTH  = 1.5;   // points (line thickness)

    // Colors
    private static final Color COLOR_PRIMARY    = new Color(0x7C, 0x6F, 0xFF);  // medaccur purple
    private static final Color COLOR_REFLINE    = new Color(0x94, 0xA3, 0xB8);  // slate gray
    private static final Color COLOR_WHISKER    = new Color(0x47, 0x55, 0x69);  // dark gray

    /**
     * Draw the complete forest plot graphic on a slide.
     * Call this AFTER placeholders have been replaced (so table text is visible).
     */
    public void addForestPlot(XSLFSlide slide, List<ForestPlotPoint> points) {
        if (points == null || points.isEmpty()) {
            log.warn("No forest plot points provided");
            return;
        }

        log.info("Drawing forest plot: {} rows", points.size());

        // 1. Draw vertical reference line at HR = 1.0
        drawReferenceLine(slide);

        // 2. Draw each data point
        for (ForestPlotPoint point : points) {
            if (point.isHeader()) {
                log.debug("  Row {}: header — skip", point.getRowIndex());
                continue;
            }

            double yCenter = getRowY(point.getRowIndex());

            if (point.getHr() <= 0 || point.getCiLower() <= 0 || point.getCiUpper() <= 0) {
                log.debug("  Row {}: invalid values — skip", point.getRowIndex());
                continue;
            }

            // Draw CI whisker line (horizontal)
            drawWhisker(slide, point.getCiLower(), point.getCiUpper(), yCenter);

            if (point.isOverall()) {
                // Draw diamond for overall result
                drawDiamond(slide, point.getHr(), point.getCiLower(), point.getCiUpper(), yCenter);
                log.debug("  Row {}: DIAMOND hr={} ci=[{}-{}]",
                    point.getRowIndex(), point.getHr(), point.getCiLower(), point.getCiUpper());
            } else {
                // Draw square for subgroup
                drawSquare(slide, point.getHr(), yCenter);
                log.debug("  Row {}: SQUARE hr={} ci=[{}-{}]",
                    point.getRowIndex(), point.getHr(), point.getCiLower(), point.getCiUpper());
            }
        }

        // 3. Draw "Favours" labels at bottom
        drawFavoursLabels(slide);

        // 4. Draw x-axis tick marks
        drawAxisTicks(slide);

        log.info("Forest plot complete: {} data points drawn",
            points.stream().filter(p -> !p.isHeader()).count());
    }

    // ═══════════════════════════════════════════
    // DRAWING METHODS
    // ═══════════════════════════════════════════

    /**
     * Vertical dashed reference line at HR = 1.0
     */
    private void drawReferenceLine(XSLFSlide slide) {
        double xPts = hrToX(1.0);
        double yTopPts = in2pt(FIRST_ROW_Y - 0.15);
        double yBotPts = in2pt(FIRST_ROW_Y + 19 * ROW_HEIGHT + 0.10);

        // POI doesn't support dashed lines easily, so use a thin rectangle
        XSLFAutoShape line = slide.createAutoShape();
        line.setShapeType(org.apache.poi.sl.usermodel.ShapeType.RECT);
        line.setAnchor(new Rectangle2D.Double(xPts - 0.5, yTopPts, 1.0, yBotPts - yTopPts));
        line.setFillColor(COLOR_REFLINE);
        line.setLineWidth(0);
    }

    /**
     * Horizontal CI whisker line from ciLower to ciUpper
     */
    private void drawWhisker(XSLFSlide slide, double ciLower, double ciUpper, double yCenter) {
        // Clamp to visible range
        double xLeft = hrToX(Math.max(ciLower, LOG_MIN));
        double xRight = hrToX(Math.min(ciUpper, LOG_MAX));
        double yPts = in2pt(yCenter);

        XSLFAutoShape whisker = slide.createAutoShape();
        whisker.setShapeType(org.apache.poi.sl.usermodel.ShapeType.RECT);
        whisker.setAnchor(new Rectangle2D.Double(
            xLeft, yPts - WHISKER_WIDTH / 2,
            xRight - xLeft, WHISKER_WIDTH
        ));
        whisker.setFillColor(COLOR_WHISKER);
        whisker.setLineWidth(0);

        // Left cap (small vertical tick)
        if (ciLower >= LOG_MIN) {
            drawTick(slide, xLeft, yPts, 4.0);
        }
        // Right cap (small vertical tick)
        if (ciUpper <= LOG_MAX) {
            drawTick(slide, xRight, yPts, 4.0);
        }
    }

    /**
     * Small vertical tick mark at end of whisker
     */
    private void drawTick(XSLFSlide slide, double xPts, double yPts, double halfHeight) {
        XSLFAutoShape tick = slide.createAutoShape();
        tick.setShapeType(org.apache.poi.sl.usermodel.ShapeType.RECT);
        tick.setAnchor(new Rectangle2D.Double(
            xPts - 0.5, yPts - halfHeight,
            1.0, halfHeight * 2
        ));
        tick.setFillColor(COLOR_WHISKER);
        tick.setLineWidth(0);
    }

    /**
     * Filled square at HR point estimate (for subgroups)
     */
    private void drawSquare(XSLFSlide slide, double hr, double yCenter) {
        double xPts = hrToX(hr);
        double size = in2pt(SQUARE_SIZE);
        double yPts = in2pt(yCenter);

        XSLFAutoShape square = slide.createAutoShape();
        square.setShapeType(org.apache.poi.sl.usermodel.ShapeType.RECT);
        square.setAnchor(new Rectangle2D.Double(
            xPts - size / 2, yPts - size / 2,
            size, size
        ));
        square.setFillColor(COLOR_PRIMARY);
        square.setLineWidth(0);
    }

    /**
     * Diamond shape at HR point estimate (for overall result).
     * Diamond width spans from ciLower to ciUpper.
     */
    private void drawDiamond(XSLFSlide slide, double hr, double ciLower, double ciUpper, double yCenter) {
        double xCenter = hrToX(hr);
        double xLeft = hrToX(Math.max(ciLower, LOG_MIN));
        double xRight = hrToX(Math.min(ciUpper, LOG_MAX));
        double yPts = in2pt(yCenter);
        double halfH = in2pt(DIAMOND_HEIGHT);

        // Diamond as a polygon: 4 points (left, top, right, bottom)
        // POI doesn't have native polygon, so we use a rotated rectangle approximation
        // Better approach: use XSLFFreeformShape if available, otherwise overlay two triangles

        // Create diamond using a small rotated rect
        // Actually, let's use 4 triangles approach with AutoShapes
        // Simplest: use DIAMOND shape type
        XSLFAutoShape diamond = slide.createAutoShape();
        diamond.setShapeType(org.apache.poi.sl.usermodel.ShapeType.DIAMOND);
        diamond.setAnchor(new Rectangle2D.Double(
            xLeft, yPts - halfH,
            xRight - xLeft, halfH * 2
        ));
        diamond.setFillColor(COLOR_PRIMARY);
        diamond.setLineColor(COLOR_PRIMARY);
        diamond.setLineWidth(0.5);
    }

    /**
     * "Favours Experimental" and "Favours Control" labels below the plot
     */
    private void drawFavoursLabels(XSLFSlide slide) {
        double yPts = in2pt(FIRST_ROW_Y + 19 * ROW_HEIGHT + 0.25);
        double refX = hrToX(1.0);

        // Left label: "Favours Experimental"
        // (actual text comes from template placeholder — this is just a backup arrow indicator)
        double leftCenterX = (in2pt(PLOT_LEFT) + refX) / 2;
        drawArrow(slide, refX - 10, yPts, in2pt(PLOT_LEFT) + 20, yPts);

        // Right label: "Favours Control"
        double rightCenterX = (refX + in2pt(PLOT_RIGHT)) / 2;
        drawArrow(slide, refX + 10, yPts, in2pt(PLOT_RIGHT) - 20, yPts);
    }

    /**
     * Simple arrow line (thin rectangle)
     */
    private void drawArrow(XSLFSlide slide, double x1, double y1, double x2, double y2) {
        double xMin = Math.min(x1, x2);
        double w = Math.abs(x2 - x1);

        XSLFAutoShape arrow = slide.createAutoShape();
        arrow.setShapeType(org.apache.poi.sl.usermodel.ShapeType.RECT);
        arrow.setAnchor(new Rectangle2D.Double(xMin, y1 - 0.5, w, 1.0));
        arrow.setFillColor(COLOR_REFLINE);
        arrow.setLineWidth(0);
    }

    /**
     * X-axis tick marks at standard HR values
     */
    private void drawAxisTicks(XSLFSlide slide) {
        double[] hrTicks = {0.1, 0.2, 0.5, 1.0, 2.0, 5.0};
        double yTop = in2pt(FIRST_ROW_Y + 19 * ROW_HEIGHT + 0.05);
        double tickH = 5.0; // points

        for (double hr : hrTicks) {
            if (hr < LOG_MIN || hr > LOG_MAX) continue;
            double xPts = hrToX(hr);

            // Tick mark
            XSLFAutoShape tick = slide.createAutoShape();
            tick.setShapeType(org.apache.poi.sl.usermodel.ShapeType.RECT);
            tick.setAnchor(new Rectangle2D.Double(xPts - 0.5, yTop, 1.0, tickH));
            tick.setFillColor(COLOR_WHISKER);
            tick.setLineWidth(0);

            // Label
            XSLFTextBox label = slide.createTextBox();
            label.setAnchor(new Rectangle2D.Double(
                xPts - in2pt(0.30), yTop + tickH + 1,
                in2pt(0.60), in2pt(0.18)
            ));
            label.clearText();
            XSLFTextParagraph p = label.addNewTextParagraph();
            p.setTextAlign(org.apache.poi.sl.usermodel.TextParagraph.TextAlign.CENTER);
            XSLFTextRun r = p.addNewTextRun();
            // Format: show "1.0" not "1"
            String txt = hr == (int) hr ? String.format("%.1f", hr) : String.valueOf(hr);
            r.setText(txt);
            r.setFontSize(7.0);
            r.setFontFamily("Calibri");
            r.setFontColor(COLOR_WHISKER);
        }
    }

    // ═══════════════════════════════════════════
    // COORDINATE HELPERS
    // ═══════════════════════════════════════════

    /**
     * Convert HR value to x-position in points (log scale).
     */
    private double hrToX(double hr) {
        double logPos = (Math.log10(hr) - Math.log10(LOG_MIN)) / LOG_RANGE;
        logPos = Math.max(0, Math.min(1, logPos)); // clamp 0-1
        return in2pt(PLOT_LEFT + logPos * PLOT_WIDTH);
    }

    /**
     * Get the y-center of a row (in inches) based on row index.
     * Row 0 = Overall, Row 1 = first header, etc.
     */
    private double getRowY(int rowIndex) {
        return FIRST_ROW_Y + rowIndex * ROW_HEIGHT;
    }

    /**
     * Inches to points.
     */
    private static double in2pt(double inches) {
        return inches * 72.0;
    }
}
