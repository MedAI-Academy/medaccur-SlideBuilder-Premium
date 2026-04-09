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
 * Updates embedded chart data in PowerPoint layouts/slides.
 * Finds charts by alt-text name (CHART_MEDIAN_PFS, CHART_ORR_RESPONSE_DRUG, etc.)
 *
 * v3.2 STRATEGY: Update ONLY the XML cache (numCache/strCache) — NOT the
 * embedded Excel workbook. POI's workbook write-back corrupts the xlsx
 * structure, causing PowerPoint's "repair" popup and black/empty charts.
 *
 * The XML cache is what PowerPoint renders on open. The Excel is only used
 * when double-clicking a chart to edit — acceptable for generated decks.
 *
 * Charts live on LAYOUTS. We update them BEFORE creating the slide.
 * The slide inherits the updated chart via visual inheritance.
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

        if (chartsByRelId.isEmpty()) {
            log.debug("No charts found on sheet");
            return;
        }
        log.info("Found {} chart(s), checking against {} key(s)", chartsByRelId.size(), chartData.size());

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
                updateXmlCache(chart, rows);
                log.info("Chart '{}' cache updated: {} rows", altText, rows.size());
            } catch (Exception e) {
                log.error("Chart '{}' failed: {}", altText, e.getMessage(), e);
            }
        }
    }

    private void updateXmlCache(XSLFChart chart, List<Map<String, Object>> rows) {
        CTChart ctChart = chart.getCTChart();
        CTPlotArea plotArea = ctChart.getPlotArea();

        for (int ci = 0; ci < plotArea.sizeOfBarChartArray(); ci++) {
            updateBarChart(plotArea.getBarChartArray(ci), rows);
        }
        for (int ci = 0; ci < plotArea.sizeOfLineChartArray(); ci++) {
            CTLineSer[] s = plotArea.getLineChartArray(ci).getSerArray();
            if (s.length > 0) updateGenericSeries(s[0].getCat(), s[0].getVal(), rows, "value");
        }
        for (int ci = 0; ci < plotArea.sizeOfPieChartArray(); ci++) {
            CTPieSer[] s = plotArea.getPieChartArray(ci).getSerArray();
            if (s.length > 0) updateGenericSeries(s[0].getCat(), s[0].getVal(), rows, "value");
        }
    }

    private void updateBarChart(CTBarChart barChart, List<Map<String, Object>> rows) {
        CTBarSer[] series = barChart.getSerArray();
        if (series.length == 0) return;

        // Categories (shared)
        if (series[0].getCat() != null && series[0].getCat().getStrRef() != null) {
            CTStrRef sr = series[0].getCat().getStrRef();
            sr.setF("Sheet1!$A$2:$A$" + (rows.size() + 1));
            if (sr.getStrCache() != null) syncStr(sr.getStrCache(), rows, "label");
        }

        // Series 1
        if (series[0].getVal() != null && series[0].getVal().getNumRef() != null) {
            CTNumRef nr = series[0].getVal().getNumRef();
            nr.setF("Sheet1!$B$2:$B$" + (rows.size() + 1));
            if (nr.getNumCache() != null) syncNum(nr.getNumCache(), rows, "value");
        }

        // Series 2
        if (series.length > 1 && rows.size() > 0 && rows.get(0).containsKey("value2")) {
            if (series[1].getVal() != null && series[1].getVal().getNumRef() != null) {
                CTNumRef nr = series[1].getVal().getNumRef();
                nr.setF("Sheet1!$C$2:$C$" + (rows.size() + 1));
                if (nr.getNumCache() != null) syncNum(nr.getNumCache(), rows, "value2");
            }
        }
    }

    private void updateGenericSeries(CTAxDataSource cat, CTNumDataSource val,
                                      List<Map<String, Object>> rows, String key) {
        if (cat != null && cat.getStrRef() != null) {
            cat.getStrRef().setF("Sheet1!$A$2:$A$" + (rows.size() + 1));
            if (cat.getStrRef().getStrCache() != null) syncStr(cat.getStrRef().getStrCache(), rows, "label");
        }
        if (val != null && val.getNumRef() != null) {
            val.getNumRef().setF("Sheet1!$B$2:$B$" + (rows.size() + 1));
            if (val.getNumRef().getNumCache() != null) syncNum(val.getNumRef().getNumCache(), rows, key);
        }
    }

    private void syncNum(CTNumData cache, List<Map<String, Object>> rows, String key) {
        cache.getPtCount().setVal(rows.size());
        while (cache.sizeOfPtArray() < rows.size()) { CTNumVal p = cache.addNewPt(); p.setIdx(cache.sizeOfPtArray()-1); p.setV("0"); }
        while (cache.sizeOfPtArray() > rows.size()) cache.removePt(cache.sizeOfPtArray()-1);
        for (int i = 0; i < rows.size(); i++) {
            cache.getPtArray(i).setIdx(i);
            Object v = rows.get(i).get(key);
            cache.getPtArray(i).setV(v instanceof Number ? String.valueOf(((Number)v).doubleValue()) : "0");
        }
    }

    private void syncStr(CTStrData cache, List<Map<String, Object>> rows, String key) {
        cache.getPtCount().setVal(rows.size());
        while (cache.sizeOfPtArray() < rows.size()) { CTStrVal p = cache.addNewPt(); p.setIdx(cache.sizeOfPtArray()-1); p.setV(""); }
        while (cache.sizeOfPtArray() > rows.size()) cache.removePt(cache.sizeOfPtArray()-1);
        for (int i = 0; i < rows.size(); i++) {
            cache.getPtArray(i).setIdx(i);
            String v = (String) rows.get(i).get(key);
            cache.getPtArray(i).setV(v != null ? v : "");
        }
    }

    private String extractAltText(XSLFShape shape) {
        try { Matcher m = ALT_TEXT.matcher(shape.getXmlObject().toString()); if (m.find()) return m.group(1); } catch (Exception e) {}
        return "";
    }

    private String extractChartRelId(XSLFShape shape) {
        try { String x = shape.getXmlObject().toString(); if (x.contains("chart")) { Matcher m = REL_ID.matcher(x); if (m.find()) return m.group(1); } } catch (Exception e) {}
        return null;
    }
}
