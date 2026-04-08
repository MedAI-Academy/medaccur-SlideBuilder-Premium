package com.medaccur.renderer.service;

import com.medaccur.renderer.model.ComparisonBar;
import com.medaccur.renderer.model.GanttBar;
import org.apache.poi.util.Units;
import org.apache.poi.xslf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import java.awt.Color;
import java.awt.geom.Rectangle2D;
import java.util.List;

@Service
public class NativeShapeService {

    private static final Logger log = LoggerFactory.getLogger(NativeShapeService.class);

    // Grid constants for TACTICAL_TIMELINE layout (in inches)
    // These MUST match the master template exactly
    private static final double GRID_LEFT = 2.20;
    private static final double GRID_WIDTH = 11.00;
    private static final double GRID_TOP = 1.50;
    private static final double ROW_HEIGHT = 1.10;
    private static final double BAR_HEIGHT = 0.32;
    private static final double BAR_OFFSET = 0.25; // offset from top of row
    private static final double BAR_RADIUS = 4.0;  // corner radius in points

    /**
     * Add Gantt bars as native AutoShapes within the template grid.
     * Each bar is a rounded rectangle positioned by month range.
     */
    public void addGanttBars(XSLFSlide slide, List<GanttBar> bars) {
        for (GanttBar bar : bars) {
            double x = GRID_LEFT + (bar.getStartMonth() - 1.0) / 12.0 * GRID_WIDTH;
            double w = (bar.getEndMonth() - bar.getStartMonth() + 1.0) / 12.0 * GRID_WIDTH;
            double y = GRID_TOP + bar.getWorkstreamIndex() * ROW_HEIGHT + BAR_OFFSET;

            // Create rounded rectangle
            XSLFAutoShape rect = slide.createAutoShape();
            rect.setShapeType(org.apache.poi.sl.usermodel.ShapeType.ROUND_RECT);
            rect.setAnchor(new Rectangle2D.Double(
                inchesToPoints(x), inchesToPoints(y),
                inchesToPoints(w), inchesToPoints(BAR_HEIGHT)
            ));
            rect.setFillColor(parseColor(bar.getColor()));
            rect.setLineWidth(0); // no border

            // Add label text inside bar
            if (bar.getLabel() != null && !bar.getLabel().isEmpty()) {
                rect.clearText();
                XSLFTextParagraph p = rect.addNewTextParagraph();
                p.setTextAlign(org.apache.poi.sl.usermodel.TextParagraph.TextAlign.CENTER);
                XSLFTextRun r = p.addNewTextRun();
                r.setText(bar.getLabel());
                r.setFontColor(Color.WHITE);
                r.setFontSize(8.0);
                r.setFontFamily("Calibri");
                try {
                    rect.setTextAutofit(org.apache.poi.sl.usermodel.TextShape.TextAutofit.NORMAL);
                } catch (Exception e) { /* ignore */ }
            }

            log.debug("  Gantt bar: ws={} months={}-{} '{}'", 
                bar.getWorkstreamIndex(), bar.getStartMonth(), bar.getEndMonth(), bar.getLabel());
        }
    }

    /**
     * Add comparison bars (e.g., Median PFS Drug vs Control).
     * Vertical bars side by side within CHART_AREA.
     */
    public void addComparisonBars(XSLFSlide slide, List<ComparisonBar> bars) {
        if (bars == null || bars.isEmpty()) return;

        // Find CHART_AREA bounds (or use default position)
        Rectangle2D area = findChartArea(slide);
        if (area == null) {
            // Default position (right column of pivotal study)
            area = new Rectangle2D.Double(
                inchesToPoints(9.0), inchesToPoints(1.5),
                inchesToPoints(4.0), inchesToPoints(5.0)
            );
        }

        double maxVal = bars.stream().mapToDouble(ComparisonBar::getValue).max().orElse(1.0);
        double barWidth = area.getWidth() / (bars.size() * 2.0); // 50% bars, 50% gap
        double gap = barWidth * 0.3;

        for (int i = 0; i < bars.size(); i++) {
            ComparisonBar bar = bars.get(i);
            double ratio = bar.getValue() / maxVal;
            double h = ratio * area.getHeight() * 0.80;
            double x = area.getX() + i * (barWidth + gap);
            double y = area.getY() + area.getHeight() - h;

            XSLFAutoShape rect = slide.createAutoShape();
            rect.setShapeType(org.apache.poi.sl.usermodel.ShapeType.ROUND_RECT);
            rect.setAnchor(new Rectangle2D.Double(x, y, barWidth, h));
            rect.setFillColor(parseColor(bar.getColor()));
            rect.setLineWidth(0);

            // Value label on top of bar
            XSLFTextBox labelBox = slide.createTextBox();
            labelBox.setAnchor(new Rectangle2D.Double(x, y - inchesToPoints(0.35), barWidth, inchesToPoints(0.30)));
            labelBox.clearText();
            XSLFTextParagraph p = labelBox.addNewTextParagraph();
            p.setTextAlign(org.apache.poi.sl.usermodel.TextParagraph.TextAlign.CENTER);
            XSLFTextRun r = p.addNewTextRun();
            String valueText = String.format("%.1f", bar.getValue());
            if (bar.getUnit() != null) valueText += " " + bar.getUnit();
            r.setText(valueText);
            r.setFontSize(11.0);
            r.setFontFamily("Calibri");
            r.setBold(true);
            r.setFontColor(parseColor(bar.getColor()));

            log.debug("  Comparison bar: '{}' = {} ({}%)", bar.getLabel(), bar.getValue(), 
                (int)(ratio * 100));
        }
    }

    private Rectangle2D findChartArea(XSLFSlide slide) {
        for (XSLFShape shape : slide.getShapes()) {
            String name = shape.getShapeName();
            if (name != null && name.toUpperCase().contains("CHART_AREA")) {
                return shape.getAnchor();
            }
        }
        return null;
    }

    private static double inchesToPoints(double inches) {
        return inches * 72.0;
    }

    private static Color parseColor(String hex) {
        if (hex == null || hex.isEmpty()) return new Color(0x7C, 0x6F, 0xFF);
        hex = hex.replace("#", "");
        return new Color(
            Integer.parseInt(hex.substring(0, 2), 16),
            Integer.parseInt(hex.substring(2, 4), 16),
            Integer.parseInt(hex.substring(4, 6), 16)
        );
    }
}
