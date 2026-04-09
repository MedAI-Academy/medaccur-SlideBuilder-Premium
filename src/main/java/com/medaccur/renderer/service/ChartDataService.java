package com.medaccur.renderer.service;

import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.xslf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.chart.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Updates embedded chart data by replacing Excel references with inline literals.
 *
 * v3.4 STRATEGY: Replace numRef/strRef (Excel links) with numLit/strLit
 * (inline values directly in chart XML). PowerPoint reads literals directly.
 *
 * Previous failures:
 * - v3.1/3.3: wb.write(os) corrupts embedded xlsx → PPT can't open
 * - v3.2: Cache-only → PPT ignores cache, renders from Excel
 * - v3.4: numLit/strLit → NO Excel contact, values in XML, PPT reads directly
 *
 * Trade-off: Double-clicking chart to edit won't show updated data in Excel
 * (Excel still has template data). For generated presentations this is fine.
 */
@Service
public class ChartDataService {

    private static final Logger log = LoggerFactory.getLogger(ChartDataService.class);
    private static final Pattern ALT_TEXT = Pattern.compile("descr=\"([^\"]*)\"");
    private static final Pattern REL_ID = Pattern.compile("r:id=\"([^\"]*)\"");

    @SuppressWarnings("unchecked")
    public void updateAllCharts(XSLFSheet sheet, Map<String, Object> chartData) {
        if (chartData == null || chartData.isEmpty()) return;

        Map<String, XSLFChart> chartsByRelId = new HashMap<>();
        for (POIXMLDocumentPart.RelationPart rp : sheet.getRelationParts()) {
            if (rp.getDocumentPart() instanceof XSLFChart) {
                chartsByRelId.put(rp.getRelationship().getId(), (XSLFChart) rp.getDocumentPart());
            }
        }

        if (chartsByRelId.isEmpty()) return;
        log.info("Found {} chart(s), {} chartData key(s)", chartsByRelId.size(), chartData.size());

        for (XSLFShape shape : sheet.getShapes()) {
            if (!(shape instanceof XSLFGraphicFrame)) continue;

            String altText = extractAltText(shape);
            if (altText.isEmpty()) continue;

            Object data = chartData.get(altText);
            if (data == null) continue;

            String relId = extractChartRelId(shape);
            if (relId == null) continue;

            XSLFChart chart = chartsByRelId.get(relId);
            if (chart == null) continue;

            Map<String, Object> chartSpec = (Map<String, Object>) data;
            List<Map<String, Object>> rows = (List<Map<String, Object>>) chartSpec.get("rows");
            if (rows == null || rows.isEmpty()) continue;

            try {
                replaceWithLiterals(chart, rows);
                log.info("Chart '{}': replaced with {} literal values", altText, rows.size());
            } catch (Exception e) {
                log.error("Chart '{}' failed: {}", altText, e.getMessage(), e);
            }
        }
    }

    // ════════════════════════════════════════════════
    // REPLACE numRef/strRef WITH numLit/strLit
    // ════════════════════════════════════════════════

    private void replaceWithLiterals(XSLFChart chart, List<Map<String, Object>> rows) {
        CTPlotArea plotArea = chart.getCTChart().getPlotArea();

        for (int ci = 0; ci < plotArea.sizeOfBarChartArray(); ci++) {
            CTBarSer[] series = plotArea.getBarChartArray(ci).getSerArray();
            for (int si = 0; si < series.length; si++) {
                CTBarSer ser = series[si];

                // Replace categories with strLit
                if (ser.getCat() != null && si == 0) {
                    replaceCategories(ser.getCat(), rows);
                }

                // Replace values with numLit
                if (ser.getVal() != null) {
                    String key = (si == 0) ? "value" : "value2";
                    replaceValues(ser.getVal(), rows, key);
                }
            }
        }

        for (int ci = 0; ci < plotArea.sizeOfLineChartArray(); ci++) {
            CTLineSer[] series = plotArea.getLineChartArray(ci).getSerArray();
            for (int si = 0; si < series.length; si++) {
                if (series[si].getCat() != null && si == 0) {
                    replaceCategories(series[si].getCat(), rows);
                }
                if (series[si].getVal() != null) {
                    replaceValues(series[si].getVal(), rows, si == 0 ? "value" : "value2");
                }
            }
        }

        for (int ci = 0; ci < plotArea.sizeOfPieChartArray(); ci++) {
            CTPieSer[] series = plotArea.getPieChartArray(ci).getSerArray();
            for (int si = 0; si < series.length; si++) {
                if (series[si].getCat() != null && si == 0) {
                    replaceCategories(series[si].getCat(), rows);
                }
                if (series[si].getVal() != null) {
                    replaceValues(series[si].getVal(), rows, "value");
                }
            }
        }
    }

    /**
     * Replace category axis data source with inline string literals.
     * Removes strRef (Excel link) and creates strLit with values directly.
     */
    private void replaceCategories(CTAxDataSource cat, List<Map<String, Object>> rows) {
        // Remove existing references
        if (cat.isSetStrRef()) cat.unsetStrRef();
        if (cat.isSetNumRef()) cat.unsetNumRef();
        if (cat.isSetStrLit()) cat.unsetStrLit();
        if (cat.isSetNumLit()) cat.unsetNumLit();

        // Create inline string literals
        CTStrData lit = cat.addNewStrLit();
        lit.addNewPtCount().setVal(rows.size());
        for (int i = 0; i < rows.size(); i++) {
            CTStrVal pt = lit.addNewPt();
            pt.setIdx(i);
            Object label = rows.get(i).get("label");
            pt.setV(label != null ? label.toString() : "");
        }
    }

    /**
     * Replace numeric value data source with inline number literals.
     * Removes numRef (Excel link) and creates numLit with values directly.
     */
    private void replaceValues(CTNumDataSource val, List<Map<String, Object>> rows, String key) {
        // Remove existing references
        if (val.isSetNumRef()) val.unsetNumRef();
        if (val.isSetNumLit()) val.unsetNumLit();

        // Create inline number literals
        CTNumData lit = val.addNewNumLit();
        lit.addNewPtCount().setVal(rows.size());
        for (int i = 0; i < rows.size(); i++) {
            CTNumVal pt = lit.addNewPt();
            pt.setIdx(i);
            Object v = rows.get(i).get(key);
            if (v instanceof Number) {
                pt.setV(String.valueOf(((Number) v).doubleValue()));
            } else {
                pt.setV("0");
            }
        }
    }

    // ════════════════════════════════════════════════
    // HELPERS
    // ════════════════════════════════════════════════

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
