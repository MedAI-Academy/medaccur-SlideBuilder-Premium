package com.medai.renderer.util;

import com.medai.renderer.template.ThemeConfig;
import org.apache.poi.sl.usermodel.TextParagraph;
import org.apache.poi.xslf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.main.*;
import org.openxmlformats.schemas.presentationml.x2006.main.CTBackground;
import org.openxmlformats.schemas.presentationml.x2006.main.CTBackgroundProperties;

import java.awt.*;
import java.awt.geom.Rectangle2D;

/**
 * Utility methods for Apache POI XSLF (PPTX) manipulation.
 * Handles backgrounds, shapes, text formatting — the low-level operations
 * that caused bugs in python-pptx (namespace issues, noFill, etc.)
 */
public class PptxUtils {

    // Conversion: inches to EMU (English Metric Units — POI's native unit)
    private static final long EMU_PER_INCH = 914400L;
    private static final long EMU_PER_PT   = 12700L;

    /** Convert inches to EMU */
    public static long emu(double inches) {
        return Math.round(inches * EMU_PER_INCH);
    }

    /** Convert points to EMU */
    public static long ptEmu(double points) {
        return Math.round(points * EMU_PER_PT);
    }

    // ═══════════════════════════════════════════════════════════
    // SLIDE BACKGROUND — THE FIX FOR THE PYTHON-PPTX BUG
    // ═══════════════════════════════════════════════════════════

    /**
     * Set slide background color correctly.
     * This is the core fix — python-pptx added inline xmlns declarations
     * on <p:bg> which PowerPoint ignores. Apache POI handles this correctly
     * by placing namespaces on the root <p:sld> element.
     */
    public static void setBackground(XSLFSlide slide, String hexColor) {
        CTBackground bg = slide.getXmlObject().getCSld().addNewBg();
        CTBackgroundProperties bgPr = bg.addNewBgPr();

        // Create solid fill with the correct color
        CTSolidColorFillProperties solidFill = bgPr.addNewSolidFill();
        CTSRgbColor srgb = solidFill.addNewSrgbClr();
        srgb.setVal(hexToBytes(hexColor));

        // Required: add empty effect list
        bgPr.addNewEffectLst();
    }

    // ═══════════════════════════════════════════════════════════
    // SHAPES — Rectangles, accent bars, cards
    // ═══════════════════════════════════════════════════════════

    /** Add a filled rectangle shape */
    public static XSLFAutoShape addRect(XSLFSlide slide,
                                         double x, double y, double w, double h,
                                         String fillHex) {
        XSLFAutoShape shape = slide.createAutoShape();
        shape.setShapeType(org.apache.poi.sl.usermodel.ShapeType.RECT);
        shape.setAnchor(new Rectangle2D.Double(x * 72, y * 72, w * 72, h * 72));
        shape.setFillColor(ThemeConfig.hex(fillHex));
        shape.setLineWidth(0.0); // No border
        return shape;
    }

    /** Add header bar (navy bar + teal accent stripe) */
    public static void addHeader(XSLFSlide slide, String title, String subtitle) {
        // Navy header bar
        addRect(slide, 0, 0, ThemeConfig.SLIDE_W, ThemeConfig.HEADER_H, ThemeConfig.HEX_NAVY);

        // Teal accent stripe on left
        addRect(slide, 0, 0, 0.18, ThemeConfig.HEADER_H, ThemeConfig.HEX_TEAL);

        // Title text
        addText(slide, title,
            0.35, 0, 9.5, ThemeConfig.HEADER_H,
            ThemeConfig.SIZE_HEADING - 2, ThemeConfig.FONT_TITLE, ThemeConfig.HEX_WHITE,
            true, TextParagraph.TextAlign.LEFT);

        // Subtitle (right-aligned)
        if (subtitle != null && !subtitle.isEmpty()) {
            addText(slide, subtitle,
                10.0, 0, 3.1, ThemeConfig.HEADER_H,
                ThemeConfig.SIZE_SMALL - 1, ThemeConfig.FONT_BODY, ThemeConfig.HEX_MUTED,
                false, TextParagraph.TextAlign.RIGHT);
        }
    }

    /** Add footer with disclaimer + references */
    public static void addFooter(XSLFSlide slide, String confidenceBadge) {
        String disclaimer = "⚠ AI-generated — Verify before external use · MedAI Suite";
        addText(slide, disclaimer,
            ThemeConfig.CONTENT_X, ThemeConfig.FOOTER_Y, ThemeConfig.CONTENT_W, ThemeConfig.FOOTER_H,
            ThemeConfig.SIZE_FOOTER, ThemeConfig.FONT_BODY, ThemeConfig.HEX_DIM,
            false, TextParagraph.TextAlign.LEFT);

        if (confidenceBadge != null && !confidenceBadge.isEmpty()) {
            addText(slide, confidenceBadge,
                ThemeConfig.CONTENT_X, ThemeConfig.FOOTER_Y - 0.22, ThemeConfig.CONTENT_W, 0.22,
                ThemeConfig.SIZE_FOOTER, ThemeConfig.FONT_MONO, ThemeConfig.HEX_TEAL,
                false, TextParagraph.TextAlign.RIGHT);
        }
    }

    // ═══════════════════════════════════════════════════════════
    // TEXT — The core text creation method
    // ═══════════════════════════════════════════════════════════

    /** Add a text box with full formatting control */
    public static XSLFTextBox addText(XSLFSlide slide, String text,
                                       double x, double y, double w, double h,
                                       double fontSize, String fontName, String colorHex,
                                       boolean bold, TextParagraph.TextAlign align) {
        XSLFTextBox textBox = slide.createTextBox();
        textBox.setAnchor(new Rectangle2D.Double(x * 72, y * 72, w * 72, h * 72));
        textBox.setWordWrap(true);
        textBox.setVerticalAlignment(org.apache.poi.sl.usermodel.VerticalAlignment.MIDDLE);

        // Enable auto-shrink to fit
        textBox.setTextAutofit(XSLFTextShape.TextAutofit.NORMAL);

        XSLFTextParagraph para = textBox.getTextParagraphs().get(0);
        para.setTextAlign(align);

        XSLFTextRun run = para.addNewTextRun();
        run.setText(text);
        run.setFontSize(fontSize);
        run.setFontFamily(fontName);
        run.setFontColor(ThemeConfig.hex(colorHex));
        run.setBold(bold);

        return textBox;
    }

    /** Add multi-line text with bullet points */
    public static XSLFTextBox addBulletText(XSLFSlide slide, String[] items,
                                             double x, double y, double w, double h,
                                             double fontSize, String colorHex) {
        XSLFTextBox textBox = slide.createTextBox();
        textBox.setAnchor(new Rectangle2D.Double(x * 72, y * 72, w * 72, h * 72));
        textBox.setWordWrap(true);
        textBox.setTextAutofit(XSLFTextShape.TextAutofit.NORMAL);

        // Remove default paragraph text
        textBox.getTextParagraphs().get(0).getTextRuns().forEach(r -> r.setText(""));

        for (int i = 0; i < items.length; i++) {
            XSLFTextParagraph para;
            if (i == 0) {
                para = textBox.getTextParagraphs().get(0);
            } else {
                para = textBox.addNewTextParagraph();
            }
            para.setTextAlign(TextParagraph.TextAlign.LEFT);
            para.setBullet(true);
            para.setIndentLevel(0);

            XSLFTextRun run = para.addNewTextRun();
            run.setText(items[i]);
            run.setFontSize(fontSize);
            run.setFontFamily(ThemeConfig.FONT_BODY);
            run.setFontColor(ThemeConfig.hex(colorHex));
        }
        return textBox;
    }

    // ═══════════════════════════════════════════════════════════
    // TABLES
    // ═══════════════════════════════════════════════════════════

    /** Create a styled data table with smart column widths */
    public static XSLFTable addTable(XSLFSlide slide,
                                      String[] headers, String[][] rows,
                                      double x, double y, double w, double rowH) {
        int numRows = rows.length + 1; // +1 for header
        int numCols = headers.length;

        // Smart column width allocation based on content
        double[] colWidths = computeColumnWidths(headers, rows, w);

        XSLFTable table = slide.createTable(numRows, numCols);
        table.setAnchor(new Rectangle2D.Double(x * 72, y * 72, w * 72, rowH * numRows * 72));

        // Set column widths
        for (int c = 0; c < numCols; c++) {
            table.setColumnWidth(c, colWidths[c] * 72);
        }

        // Header row
        for (int c = 0; c < numCols; c++) {
            XSLFTableCell cell = table.getCell(0, c);
            setCellText(cell, headers[c], ThemeConfig.SIZE_SMALL, ThemeConfig.HEX_WHITE, true);
            setCellFill(cell, ThemeConfig.HEX_ACCENT);
        }

        // Data rows
        for (int r = 0; r < rows.length; r++) {
            String bgHex = (r % 2 == 0) ? ThemeConfig.HEX_SURFACE : ThemeConfig.HEX_NAVY;
            for (int c = 0; c < Math.min(rows[r].length, numCols); c++) {
                XSLFTableCell cell = table.getCell(r + 1, c);
                setCellText(cell, rows[r][c], ThemeConfig.SIZE_SMALL - 1, ThemeConfig.HEX_TEXT, false);
                setCellFill(cell, bgHex);
            }
        }

        return table;
    }

    /** Compute proportional column widths based on content length */
    private static double[] computeColumnWidths(String[] headers, String[][] rows, double totalW) {
        int n = headers.length;
        double[] weights = new double[n];

        // Base weight from header length
        for (int c = 0; c < n; c++) {
            weights[c] = Math.max(3.0, headers[c].length());
        }

        // Adjust weights based on max content length in each column (sample first 5 rows)
        for (int r = 0; r < Math.min(rows.length, 5); r++) {
            for (int c = 0; c < Math.min(rows[r].length, n); c++) {
                double contentLen = rows[r][c] != null ? rows[r][c].length() : 0;
                weights[c] = Math.max(weights[c], contentLen * 0.8);
            }
        }

        // Known narrow columns (by header name)
        for (int c = 0; c < n; c++) {
            String h = headers[c].toLowerCase();
            if (h.equals("n") || h.equals("phase") || h.equals("status") || h.equals("si")
                || h.equals("budget") || h.equals("orr") || h.equals("year")) {
                weights[c] = Math.min(weights[c], 8.0);  // Cap narrow columns
            }
        }

        // Normalize to total width with minimum 0.8" per column
        double sum = 0;
        for (double w2 : weights) sum += w2;
        double[] widths = new double[n];
        double minCol = 0.8;
        double remaining = totalW;
        int capped = 0;

        // First pass: cap minimum widths
        for (int c = 0; c < n; c++) {
            double proposed = (weights[c] / sum) * totalW;
            if (proposed < minCol) {
                widths[c] = minCol;
                remaining -= minCol;
                capped++;
            }
        }

        // Second pass: distribute remaining width
        double uncappedSum = 0;
        for (int c = 0; c < n; c++) {
            if (widths[c] == 0) uncappedSum += weights[c];
        }
        for (int c = 0; c < n; c++) {
            if (widths[c] == 0) {
                widths[c] = (weights[c] / uncappedSum) * remaining;
            }
        }

        return widths;
    }

    private static void setCellText(XSLFTableCell cell, String text,
                                     double fontSize, String colorHex, boolean bold) {
        cell.clearText();
        cell.setVerticalAlignment(org.apache.poi.sl.usermodel.VerticalAlignment.TOP);
        // Add cell padding for breathing room
        cell.setTopInset(4.0);
        cell.setBottomInset(4.0);
        cell.setLeftInset(6.0);
        cell.setRightInset(6.0);
        XSLFTextParagraph para = cell.addNewTextParagraph();
        XSLFTextRun run = para.addNewTextRun();
        run.setText(text != null ? text : "");
        run.setFontSize(fontSize);
        run.setFontFamily(ThemeConfig.FONT_BODY);
        run.setFontColor(ThemeConfig.hex(colorHex));
        run.setBold(bold);
    }

    private static void setCellFill(XSLFTableCell cell, String hexColor) {
        cell.setFillColor(ThemeConfig.hex(hexColor));
    }

    // ═══════════════════════════════════════════════════════════
    // IMAGES — Embed PNG (for charts, logos)
    // ═══════════════════════════════════════════════════════════

    /** Embed a PNG image from byte array */
    public static void addImage(XSLFSlide slide, XMLSlideShow pptx, byte[] imageData,
                                 double x, double y, double w, double h) {
        XSLFPictureData pictureData = pptx.addPicture(imageData,
            org.apache.poi.sl.usermodel.PictureData.PictureType.PNG);
        XSLFPictureShape picture = slide.createPicture(pictureData);
        picture.setAnchor(new Rectangle2D.Double(x * 72, y * 72, w * 72, h * 72));
    }

    // ═══════════════════════════════════════════════════════════
    // UTILITY
    // ═══════════════════════════════════════════════════════════

    /** Convert hex string (e.g. "0B1A3B") to byte array */
    public static byte[] hexToBytes(String hex) {
        return new byte[] {
            (byte) Integer.parseInt(hex.substring(0, 2), 16),
            (byte) Integer.parseInt(hex.substring(2, 4), 16),
            (byte) Integer.parseInt(hex.substring(4, 6), 16)
        };
    }
}
