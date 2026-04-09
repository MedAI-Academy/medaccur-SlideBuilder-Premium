package com.medaccur.renderer.service;

import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.xslf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.chart.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import java.io.OutputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Updates embedded Excel chart data in PowerPoint layouts.
 *
 * v3.3 STRATEGY (Gemini insight + our learnings):
 *   1. Update cell values in the embedded XSSFWorkbook
 *   2. Write workbook back to the chart's package part
 *   3. DELETE the XML cache (unsetNumCache / unsetStrCache)
 *      → Forces PowerPoint to regenerate chart from Excel on open
 *
 * Previous failures:
 * - v3.1: Write-back corrupted xlsx → repair popup + black charts
 * - v3.2: Cache-only → PowerPoint ignores cache, renders from Excel
 * - v3.3: Write-back + DELETE cache → PowerPoint MUST read fresh Excel
 */
@Service
public class ChartDataService {

    private static final Logger log = LoggerFactory.getLogger(ChartDataService.class);
    private static final Pattern ALT_TEXT = Pattern.compile("descr=\"([^\"]*)\"");
    private static final Pattern REL_ID = Pattern.compile("r:id=\"([^\"]*)\"");

    @SuppressWarnings("unchecked")
    public void updateAllCharts(XSLFSheet sheet, Map<String, Object> chartData) {
        if (chartData == null || chartData.isEmpty()) return;

        // Build relId → XSLFChart map
        Map<String, XSLFChart> chartsByRelId = new HashMap<>();
        for (POIXMLDocumentPart.RelationPart rp : sheet.getRelationParts()) {
            if (rp.getDocumentPart() instanceof XSLFChart) {
                chartsByRelId.put(rp.getRelationship().getId(), (XSLFChart) rp.getDocumentPart());
            }
        }

        if (chartsByRelId.isEmpty()) {
            log.debug("No charts on sheet");
            return;
        }
        log.info("Found {} chart(s), {} chartData key(s)", chartsByRelId.size(), chartData.size());

        for (XSLFShape shape : sheet.getShapes()) {
            if (!(shape instanceof XSLFGraphicFrame)) continue;

            String altText = extractAltText(shape);
            if (altText.isEmpty()) continue;

            Object data = chartData.get(altText);
            if (data == null) continue;

            String relId = extractChartRelId(shape);
            if (relId == null) {
                log.warn("Chart '{}' has no r:id", altText);
                continue;
            }

            XSLFChart chart = chartsByRelId.get(relId);
            if (chart == null) {
                log.warn("No chart for relId '{}' ('{}')", relId, altText);
                continue;
            }

            Map<String, Object> chartSpec = (Map<String, Object>) data;
            List<Map<String, Object>> rows = (List<Map<String, Object>>) chartSpec.get("rows");
            if (rows == null || rows.isEmpty()) continue;

            try {
                // ── Step 1: Update Excel cells ──
                XSSFWorkbook wb = chart.getWorkbook();
                XSSFSheet excelSheet = wb.getSheetAt(0);
                updateExcelCells(excelSheet, rows);

                // ── Step 2: Write workbook back to package part ──
                writeBackWorkbook(chart, wb);

                // ── Step 3: DELETE the XML cache → force PPT refresh ──
                deleteCaches(chart);

                log.info("Chart '{}': {} rows updated, cache deleted (PPT will refresh)", 
                         altText, rows.size());

            } catch (Exception e) {
                log.error("Chart '{}' failed: {}", altText, e.getMessage(), e);
            }
        }
    }

    // ════════════════════════════════════════════════
    // STEP 1: UPDATE EXCEL CELLS
    // ════════════════════════════════════════════════

    private void updateExcelCells(XSSFSheet sheet, List<Map<String, Object>> rows) {
        for (int i = 0; i < rows.size(); i++) {
            XSSFRow row = sheet.getRow(i + 1);
            if (row == null) row = sheet.createRow(i + 1);
            Map<String, Object> data = rows.get(i);

            // Col A: Label
            setCell(row, 0, data.get("label"));
            // Col B: Value
            setCell(row, 1, data.get("value"));
            // Col C: Value2 (optional, for dual-series)
            if (data.containsKey("value2")) {
                setCell(row, 2, data.get("value2"));
            }
        }

        // Clear excess rows (template may have had more rows)
        for (int i = rows.size() + 1; i <= 20; i++) {
            XSSFRow row = sheet.getRow(i);
            if (row == null) continue;
            for (int c = 0; c < 4; c++) {
                XSSFCell cell = row.getCell(c);
                if (cell != null) {
                    if (c == 0) cell.setCellValue("");
                    else cell.setCellValue(0);
                }
            }
        }
    }

    // ════════════════════════════════════════════════
    // STEP 2: WRITE WORKBOOK BACK TO PACKAGE PART
    // ════════════════════════════════════════════════

    private void writeBackWorkbook(XSLFChart chart, XSSFWorkbook wb) {
        try {
            for (POIXMLDocumentPart.RelationPart rp : chart.getRelationParts()) {
                String ct = rp.getDocumentPart().getPackagePart().getContentType();
                if (ct.contains("spreadsheetml") || ct.contains("excel")) {
                    try (OutputStream os = rp.getDocumentPart().getPackagePart().getOutputStream()) {
                        wb.write(os);
                    }
                    log.debug("  Workbook written back");
                    return;
                }
            }
            log.warn("  No workbook part found for write-back");
        } catch (Exception e) {
            log.error("  Write-back failed: {}", e.getMessage());
        }
    }

    // ════════════════════════════════════════════════
    // STEP 3: DELETE XML CACHE (the Gemini insight!)
    // ════════════════════════════════════════════════

    /**
     * Delete all numCache and strCache from chart series.
     * This forces PowerPoint to re-read from the embedded Excel on open,
     * which now contains our updated values.
     *
     * Without this, PowerPoint uses the stale cache and shows old bars.
     */
    private void deleteCaches(XSLFChart chart) {
        try {
            CTChart ctChart = chart.getCTChart();
            CTPlotArea plotArea = ctChart.getPlotArea();

            // Bar charts
            for (int ci = 0; ci < plotArea.sizeOfBarChartArray(); ci++) {
                CTBarSer[] series = plotArea.getBarChartArray(ci).getSerArray();
                for (CTBarSer ser : series) {
                    if (ser.getVal() != null && ser.getVal().getNumRef() != null) {
                        ser.getVal().getNumRef().unsetNumCache();
                    }
                    if (ser.getCat() != null) {
                        if (ser.getCat().getStrRef() != null) {
                            ser.getCat().getStrRef().unsetStrCache();
                        }
                        if (ser.getCat().getNumRef() != null) {
                            ser.getCat().getNumRef().unsetNumCache();
                        }
                    }
                }
            }

            // Line charts
            for (int ci = 0; ci < plotArea.sizeOfLineChartArray(); ci++) {
                CTLineSer[] series = plotArea.getLineChartArray(ci).getSerArray();
                for (CTLineSer ser : series) {
                    if (ser.getVal() != null && ser.getVal().getNumRef() != null) {
                        ser.getVal().getNumRef().unsetNumCache();
                    }
                    if (ser.getCat() != null) {
                        if (ser.getCat().getStrRef() != null) {
                            ser.getCat().getStrRef().unsetStrCache();
                        }
                    }
                }
            }

            // Pie charts
            for (int ci = 0; ci < plotArea.sizeOfPieChartArray(); ci++) {
                CTPieSer[] series = plotArea.getPieChartArray(ci).getSerArray();
                for (CTPieSer ser : series) {
                    if (ser.getVal() != null && ser.getVal().getNumRef() != null) {
                        ser.getVal().getNumRef().unsetNumCache();
                    }
                    if (ser.getCat() != null) {
                        if (ser.getCat().getStrRef() != null) {
                            ser.getCat().getStrRef().unsetStrCache();
                        }
                    }
                }
            }

            // Scatter charts (future: embedded KM curves)
            for (int ci = 0; ci < plotArea.sizeOfScatterChartArray(); ci++) {
                CTScatterSer[] series = plotArea.getScatterChartArray(ci).getSerArray();
                for (CTScatterSer ser : series) {
                    if (ser.getYVal() != null && ser.getYVal().getNumRef() != null) {
                        ser.getYVal().getNumRef().unsetNumCache();
                    }
                    if (ser.getXVal() != null) {
                        if (ser.getXVal().getNumRef() != null) {
                            ser.getXVal().getNumRef().unsetNumCache();
                        }
                    }
                }
            }

            log.debug("  All caches deleted");

        } catch (Exception e) {
            log.warn("  Cache deletion partial: {}", e.getMessage());
        }
    }

    // ════════════════════════════════════════════════
    // HELPERS
    // ════════════════════════════════════════════════

    private void setCell(XSSFRow row, int col, Object value) {
        XSSFCell cell = row.getCell(col);
        if (cell == null) cell = row.createCell(col);
        if (value instanceof Number) {
            cell.setCellValue(((Number) value).doubleValue());
        } else if (value instanceof String) {
            cell.setCellValue((String) value);
        } else {
            cell.setCellValue(value != null ? value.toString() : "");
        }
    }

    private String extractAltText(XSLFShape shape) {
        try {
            Matcher m = ALT_TEXT.matcher(shape.getXmlObject().toString());
            if (m.find()) return m.group(1);
        } catch (Exception e) { /* ignore */ }
        return "";
    }

    private String extractChartRelId(XSLFShape shape) {
        try {
            String xml = shape.getXmlObject().toString();
            if (xml.contains("chart")) {
                Matcher m = REL_ID.matcher(xml);
                if (m.find()) return m.group(1);
            }
        } catch (Exception e) { /* ignore */ }
        return null;
    }
}
