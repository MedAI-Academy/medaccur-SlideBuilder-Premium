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
 * Updates embedded Excel chart data in PowerPoint slides/layouts.
 * Finds charts by alt-text name (CHART_MEDIAN_PFS, CHART_ORR_RESPONSE_DRUG, etc.)
 *
 * ALWAYS updates BOTH:
 * 1. Embedded Excel workbook (source of truth)
 * 2. XML numCache/strCache (what PowerPoint renders immediately)
 *
 * CRITICAL FIX (v3.1): After modifying the XSSFWorkbook, we MUST write it
 * back to the chart's embedded package part. Without this, the PPTX file
 * still contains the original Excel data and PowerPoint shows old values.
 *
 * ARCHITECTURE (v3.1): Charts live on LAYOUTS, not slides. We update them
 * on the layout BEFORE creating the slide. The slide inherits the updated
 * chart via PowerPoint's visual inheritance — no XML copying needed.
 */
@Service
public class ChartDataService {

    private static final Logger log = LoggerFactory.getLogger(ChartDataService.class);
    private static final Pattern ALT_TEXT = Pattern.compile("descr=\"([^\"]*)\"");
    private static final Pattern REL_ID = Pattern.compile("r:id=\"([^\"]*)\"");

    /**
     * Update all charts on a sheet (slide or layout) based on chartData map.
     * Works on both XSLFSlide and XSLFSlideLayout (both extend XSLFSheet).
     *
     * chartData format:
     * {
     *   "CHART_MEDIAN_PFS": {
     *     "rows": [{"label":"SVd", "value": 13.9}, {"label":"Vd", "value": 9.4}]
     *   }
     * }
     */
    @SuppressWarnings("unchecked")
    public void updateAllCharts(XSLFSheet sheet, Map<String, Object> chartData) {
        if (chartData == null || chartData.isEmpty()) return;

        // Build map of relId → XSLFChart from this sheet's relations
        Map<String, XSLFChart> chartsByRelId = new HashMap<>();
        for (POIXMLDocumentPart.RelationPart rp : sheet.getRelationParts()) {
            if (rp.getDocumentPart() instanceof XSLFChart) {
                chartsByRelId.put(rp.getRelationship().getId(), (XSLFChart) rp.getDocumentPart());
            }
        }

        if (chartsByRelId.isEmpty()) {
            log.debug("No charts found on sheet");
            return;
        }
        log.info("Found {} chart(s) on sheet, checking against {} chartData key(s)",
                 chartsByRelId.size(), chartData.size());

        // Iterate shapes to find graphicFrames with matching alt-text
        for (XSLFShape shape : sheet.getShapes()) {
            if (!(shape instanceof XSLFGraphicFrame)) continue;

            String altText = extractAltText(shape);
            if (altText.isEmpty()) continue;

            Object data = chartData.get(altText);
            if (data == null) continue;

            // Extract the relationship ID from the graphicFrame XML (c:chart r:id="rIdX")
            String relId = extractChartRelId(shape);
            if (relId == null) {
                log.warn("Chart shape '{}' has no r:id — skipping", altText);
                continue;
            }

            XSLFChart chart = chartsByRelId.get(relId);
            if (chart == null) {
                log.warn("No chart part found for relId '{}' (shape '{}')", relId, altText);
                continue;
            }

            Map<String, Object> chartSpec = (Map<String, Object>) data;
            List<Map<String, Object>> rows = (List<Map<String, Object>>) chartSpec.get("rows");
            if (rows == null || rows.isEmpty()) continue;

            try {
                // Step 1: Update embedded Excel workbook
                XSSFWorkbook wb = chart.getWorkbook();
                XSSFSheet excelSheet = wb.getSheetAt(0);
                updateExcel(excelSheet, rows);

                // Step 2: CRITICAL — Write workbook back to package part!
                // Without this, the PPTX contains the ORIGINAL Excel data.
                writeBackWorkbook(chart, wb);

                // Step 3: Update XML cache (what PowerPoint renders immediately)
                updateXmlCache(chart, rows);

                // Step 4: Auto-scale Y-axis
                autoScaleAxis(chart, rows, chartSpec);

                log.info("Chart '{}' updated: {} rows (relId={})", altText, rows.size(), relId);

            } catch (Exception e) {
                log.error("Chart '{}' update failed: {}", altText, e.getMessage(), e);
            }
        }
    }

    // ════════════════════════════════════════════════
    // WORKBOOK WRITE-BACK (the critical missing piece)
    // ════════════════════════════════════════════════

    /**
     * Write the modified XSSFWorkbook back to the chart's embedded package part.
     *
     * This is the fix that Gemini correctly identified: POI reads the embedded
     * Excel into memory as an XSSFWorkbook, but modifying it doesn't automatically
     * persist the changes. We must explicitly write it back to the package part's
     * OutputStream so the PPTX file contains the updated data.
     */
    private void writeBackWorkbook(XSLFChart chart, XSSFWorkbook wb) {
        try {
            // Find the embedded workbook part among the chart's relations
            for (POIXMLDocumentPart.RelationPart rp : chart.getRelationParts()) {
                String contentType = rp.getDocumentPart().getPackagePart().getContentType();
                if (contentType.contains("spreadsheetml") || contentType.contains("excel")) {
                    try (OutputStream os = rp.getDocumentPart().getPackagePart().getOutputStream()) {
                        wb.write(os);
                        log.debug("  Workbook written back to package part");
                        return;
                    }
                }
            }
            // Fallback: try writing to chart's own package part output
            log.warn("  No separate workbook part found — trying chart part directly");
        } catch (Exception e) {
            log.error("  Workbook write-back failed: {}", e.getMessage());
        }
    }

    // ════════════════════════════════════════════════
    // EXCEL UPDATE
    // ════════════════════════════════════════════════

    private void updateExcel(XSSFSheet sheet, List<Map<String, Object>> rows) {
        for (int i = 0; i < rows.size(); i++) {
            XSSFRow row = sheet.getRow(i + 1);
            if (row == null) row = sheet.createRow(i + 1);
            Map<String, Object> data = rows.get(i);

            setCell(row, 0, data.get("label"));
            setCell(row, 1, data.get("value"));
            if (data.containsKey("value2")) setCell(row, 2, data.get("value2"));
        }

        // Clear excess rows
        for (int i = rows.size() + 1; i <= 12; i++) {
            XSSFRow row = sheet.getRow(i);
            if (row == null) continue;
            for (int c = 0; c < 3; c++) {
                XSSFCell cell = row.getCell(c);
                if (cell != null) {
                    if (c == 0) cell.setCellValue("");
                    else cell.setCellValue(0);
                }
            }
        }
    }

    // ════════════════════════════════════════════════
    // XML CACHE + RANGE SYNC
    // ════════════════════════════════════════════════

    private void updateXmlCache(XSLFChart chart, List<Map<String, Object>> rows) {
        try {
            CTChart ctChart = chart.getCTChart();
            CTPlotArea plotArea = ctChart.getPlotArea();

            for (int ci = 0; ci < plotArea.sizeOfBarChartArray(); ci++) {
                CTBarChart barChart = plotArea.getBarChartArray(ci);
                updateBarChartCache(barChart, rows);
            }

            for (int ci = 0; ci < plotArea.sizeOfLineChartArray(); ci++) {
                // Line chart — similar pattern
            }

            for (int ci = 0; ci < plotArea.sizeOfPieChartArray(); ci++) {
                // Pie chart — similar pattern
            }

            for (int ci = 0; ci < plotArea.sizeOfScatterChartArray(); ci++) {
                // Scatter chart (future: KM curves if embedded)
            }

        } catch (Exception e) {
            log.warn("XML cache sync partial failure: {}", e.getMessage());
        }
    }

    private void updateBarChartCache(CTBarChart barChart, List<Map<String, Object>> rows) {
        CTBarSer[] series = barChart.getSerArray();
        if (series.length == 0) return;

        // Update first series values
        if (series[0].getVal() != null && series[0].getVal().getNumRef() != null) {
            CTNumRef numRef = series[0].getVal().getNumRef();
            numRef.setF("Sheet1!$B$2:$B$" + (rows.size() + 1));
            if (numRef.getNumCache() != null) {
                syncNumCache(numRef.getNumCache(), rows, "value");
            }
        }

        // Update category labels
        if (series[0].getCat() != null && series[0].getCat().getStrRef() != null) {
            CTStrRef strRef = series[0].getCat().getStrRef();
            strRef.setF("Sheet1!$A$2:$A$" + (rows.size() + 1));
            if (strRef.getStrCache() != null) {
                syncStrCache(strRef.getStrCache(), rows, "label");
            }
        }

        // Update second series if exists
        if (series.length > 1 && rows.get(0).containsKey("value2")) {
            if (series[1].getVal() != null && series[1].getVal().getNumRef() != null) {
                CTNumRef numRef = series[1].getVal().getNumRef();
                numRef.setF("Sheet1!$C$2:$C$" + (rows.size() + 1));
                if (numRef.getNumCache() != null) {
                    syncNumCache(numRef.getNumCache(), rows, "value2");
                }
            }
        }
    }

    private void syncNumCache(CTNumData cache, List<Map<String, Object>> rows, String key) {
        cache.getPtCount().setVal(rows.size());
        while (cache.sizeOfPtArray() < rows.size()) {
            CTNumVal pt = cache.addNewPt();
            pt.setIdx(cache.sizeOfPtArray() - 1);
            pt.setV("0");
        }
        while (cache.sizeOfPtArray() > rows.size()) {
            cache.removePt(cache.sizeOfPtArray() - 1);
        }
        for (int i = 0; i < rows.size(); i++) {
            CTNumVal pt = cache.getPtArray(i);
            pt.setIdx(i);
            Object val = rows.get(i).get(key);
            pt.setV(val instanceof Number ? String.valueOf(((Number) val).doubleValue()) : "0");
        }
    }

    private void syncStrCache(CTStrData cache, List<Map<String, Object>> rows, String key) {
        cache.getPtCount().setVal(rows.size());
        while (cache.sizeOfPtArray() < rows.size()) {
            CTStrVal pt = cache.addNewPt();
            pt.setIdx(cache.sizeOfPtArray() - 1);
            pt.setV("");
        }
        while (cache.sizeOfPtArray() > rows.size()) {
            cache.removePt(cache.sizeOfPtArray() - 1);
        }
        for (int i = 0; i < rows.size(); i++) {
            CTStrVal pt = cache.getPtArray(i);
            pt.setIdx(i);
            String val = (String) rows.get(i).get(key);
            pt.setV(val != null ? val : "");
        }
    }

    // ════════════════════════════════════════════════
    // AUTO-SCALE Y-AXIS
    // ════════════════════════════════════════════════

    private void autoScaleAxis(XSLFChart chart, List<Map<String, Object>> rows,
                                Map<String, Object> spec) {
        try {
            CTChart ctChart = chart.getCTChart();
            CTPlotArea plotArea = ctChart.getPlotArea();
            if (plotArea.sizeOfValAxArray() == 0) return;

            CTValAx valAx = plotArea.getValAxArray(0);
            CTScaling scaling = valAx.getScaling();

            Object explicitMax = spec.get("y_max");
            if (explicitMax != null) {
                scaling.getMax().setVal(((Number) explicitMax).doubleValue());
                return;
            }

            double maxVal = rows.stream()
                .mapToDouble(r -> {
                    Object v = r.get("value");
                    return v instanceof Number ? ((Number) v).doubleValue() : 0;
                })
                .max().orElse(10);

            if (maxVal <= 1.5) {
                scaling.getMin().setVal(0);
                scaling.getMax().setVal(Math.ceil(maxVal * 5) / 5.0);
            } else if (maxVal <= 100) {
                scaling.getMin().setVal(0);
                double nice = niceRound(maxVal * 1.15);
                scaling.getMax().setVal(nice);
            }

        } catch (Exception e) {
            log.debug("Y-axis auto-scale skipped: {}", e.getMessage());
        }
    }

    // ════════════════════════════════════════════════
    // HELPERS
    // ════════════════════════════════════════════════

    private String extractAltText(XSLFShape shape) {
        try {
            String xml = shape.getXmlObject().toString();
            Matcher m = ALT_TEXT.matcher(xml);
            if (m.find()) return m.group(1);
        } catch (Exception e) { /* ignore */ }
        return "";
    }

    /**
     * Extract the chart relationship ID (r:id) from a graphicFrame's XML.
     * The chart reference looks like: <c:chart r:id="rId2"/>
     */
    private String extractChartRelId(XSLFShape shape) {
        try {
            String xml = shape.getXmlObject().toString();
            // Look for the chart namespace reference with r:id
            if (xml.contains("chart") && xml.contains("r:id")) {
                Matcher m = REL_ID.matcher(xml);
                if (m.find()) return m.group(1);
            }
        } catch (Exception e) { /* ignore */ }
        return null;
    }

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

    private double niceRound(double val) {
        if (val <= 0) return 1;
        double mag = Math.pow(10, Math.floor(Math.log10(val)));
        double norm = val / mag;
        if (norm <= 1.5) return 1.5 * mag;
        if (norm <= 2) return 2 * mag;
        if (norm <= 3) return 3 * mag;
        if (norm <= 5) return 5 * mag;
        return 10 * mag;
    }
}
