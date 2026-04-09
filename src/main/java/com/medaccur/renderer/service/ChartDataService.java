package com.medaccur.renderer.service;

import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.xslf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.chart.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import java.io.*;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.zip.*;

/**
 * v3.6: Raw xlsx ZIP edit with INLINE strings (bypass sharedStrings entirely)
 * + XML cache update + chart-to-slide migration support.
 *
 * Strategy:
 * 1. Raw-edit embedded xlsx: numbers in sheet1.xml, labels as inlineStr
 * 2. Update chart XML cache for immediate rendering
 * 3. Both Excel AND cache show correct data → PowerPoint is consistent
 *
 * Key change from v3.5: Labels use <c t="inlineStr"><is><t>SVd</t></is></c>
 * instead of shared string references. This bypasses the fragile sharedStrings
 * rebuild that failed for ORR Control charts.
 */
@Service
public class ChartDataService {

    private static final Logger log = LoggerFactory.getLogger(ChartDataService.class);
    private static final Pattern ALT_TEXT_PAT = Pattern.compile("descr=\"([^\"]*)\"");
    private static final Pattern REL_ID_PAT = Pattern.compile("r:id=\"([^\"]*)\"");

    @SuppressWarnings("unchecked")
    public void updateAllCharts(XSLFSheet sheet, Map<String, Object> chartData) {
        if (chartData == null || chartData.isEmpty()) return;

        Map<String, XSLFChart> chartsByRelId = new HashMap<>();
        Map<String, PackagePart> xlsxParts = new HashMap<>();

        for (POIXMLDocumentPart.RelationPart rp : sheet.getRelationParts()) {
            if (rp.getDocumentPart() instanceof XSLFChart) {
                XSLFChart chart = (XSLFChart) rp.getDocumentPart();
                String rid = rp.getRelationship().getId();
                chartsByRelId.put(rid, chart);
                for (POIXMLDocumentPart.RelationPart crp : chart.getRelationParts()) {
                    String ct = crp.getDocumentPart().getPackagePart().getContentType();
                    if (ct.contains("spreadsheet") || ct.contains("excel")) {
                        xlsxParts.put(rid, crp.getDocumentPart().getPackagePart());
                    }
                }
            }
        }

        if (chartsByRelId.isEmpty()) return;
        log.info("Found {} chart(s), {} xlsx parts", chartsByRelId.size(), xlsxParts.size());

        for (XSLFShape shape : sheet.getShapes()) {
            if (!(shape instanceof XSLFGraphicFrame)) continue;
            String altText = extractAltText(shape);
            if (altText.isEmpty()) continue;
            Object data = chartData.get(altText);
            if (data == null) continue;
            String relId = extractChartRelId(shape);
            if (relId == null) continue;
            XSLFChart chart = chartsByRelId.get(relId);
            PackagePart xlsx = xlsxParts.get(relId);
            if (chart == null) continue;

            Map<String, Object> spec = (Map<String, Object>) data;
            List<Map<String, Object>> rows = (List<Map<String, Object>>) spec.get("rows");
            if (rows == null || rows.isEmpty()) continue;

            try {
                if (xlsx != null) {
                    rawEditXlsx(xlsx, rows);
                    log.info("Chart '{}': xlsx updated", altText);
                }
                updateXmlCache(chart, rows);
                log.info("Chart '{}': cache updated ({} rows)", altText, rows.size());
            } catch (Exception e) {
                log.error("Chart '{}' failed: {}", altText, e.getMessage(), e);
            }
        }
    }

    // ════════════════════════════════════════════════
    // RAW XLSX EDIT
    // ════════════════════════════════════════════════

    private void rawEditXlsx(PackagePart xlsxPart, List<Map<String, Object>> rows) throws Exception {
        byte[] original;
        try (InputStream is = xlsxPart.getInputStream()) {
            original = is.readAllBytes();
        }

        ByteArrayOutputStream modified = new ByteArrayOutputStream();
        try (ZipInputStream zis = new ZipInputStream(new ByteArrayInputStream(original));
             ZipOutputStream zos = new ZipOutputStream(modified)) {

            ZipEntry entry;
            while ((entry = zis.getNextEntry()) != null) {
                byte[] content = zis.readAllBytes();
                String name = entry.getName();

                if (name.endsWith("/sheet1.xml") && name.contains("worksheets")) {
                    content = buildSheetXml(content, rows);
                } else if (name.contains("tables/table") && name.endsWith(".xml")) {
                    content = updateTableRange(content, rows.size());
                }
                // sharedStrings.xml: leave untouched — we use inlineStr now

                ZipEntry newEntry = new ZipEntry(name);
                zos.putNextEntry(newEntry);
                zos.write(content);
                zos.closeEntry();
            }
        }

        try (OutputStream os = xlsxPart.getOutputStream()) {
            os.write(modified.toByteArray());
        }
    }

    /**
     * Build new sheet1.xml with inline string labels and numeric values.
     * Uses <c t="inlineStr"><is><t>LABEL</t></is></c> for labels —
     * this bypasses sharedStrings entirely and works reliably.
     */
    private byte[] buildSheetXml(byte[] original, List<Map<String, Object>> rows) {
        String xml = new String(original, java.nio.charset.StandardCharsets.UTF_8);

        // Extract the part before sheetData and after sheetData
        int sdStart = xml.indexOf("<sheetData>");
        int sdEnd = xml.indexOf("</sheetData>") + "</sheetData>".length();
        if (sdStart < 0 || sdEnd < 0) return original;

        String before = xml.substring(0, sdStart);
        String after = xml.substring(sdEnd);

        // Update dimension
        int lastRow = rows.size() + 1;
        before = before.replaceFirst("ref=\"[^\"]+\"", "ref=\"A1:B" + lastRow + "\"");

        // Extract header series name from original row 1 col B
        // (e.g. "mPFS", "DRUG", "Series 1")
        String seriesName = "Series";
        Pattern headerPat = Pattern.compile("<row r=\"1\"[^>]*>.*?</row>", Pattern.DOTALL);
        Matcher hm = headerPat.matcher(xml);
        if (hm.find()) {
            String headerRow = hm.group();
            // Look for the B1 cell value — could be shared string ref or inline
            Pattern b1Val = Pattern.compile("<c r=\"B1\"[^>]*>.*?</c>", Pattern.DOTALL);
            Matcher b1m = b1Val.matcher(headerRow);
            if (b1m.find()) {
                // If t="s", it's a shared string index — try to read from sharedStrings
                // For simplicity, just keep B1 as-is by preserving the whole header row
            }
        }

        // Build new sheetData with inline strings
        StringBuilder sd = new StringBuilder("<sheetData>");

        // Row 1: header — preserve original header row exactly
        if (hm.find(0)) {
            sd.append(hm.group());
        } else {
            // Fallback: simple header
            sd.append("<row r=\"1\" spans=\"1:2\">");
            sd.append("<c r=\"A1\" t=\"inlineStr\"><is><t> </t></is></c>");
            sd.append("<c r=\"B1\" t=\"inlineStr\"><is><t>").append(seriesName).append("</t></is></c>");
            sd.append("</row>");
        }

        // Data rows: inline label + numeric value
        for (int i = 0; i < rows.size(); i++) {
            int rowNum = i + 2;
            String label = rows.get(i).get("label") != null ? rows.get(i).get("label").toString() : "";
            Object val = rows.get(i).get("value");
            double numVal = val instanceof Number ? ((Number) val).doubleValue() : 0.0;

            sd.append("<row r=\"").append(rowNum).append("\" spans=\"1:2\">");
            sd.append("<c r=\"A").append(rowNum).append("\" t=\"inlineStr\">");
            sd.append("<is><t>").append(escapeXml(label)).append("</t></is></c>");
            sd.append("<c r=\"B").append(rowNum).append("\">");
            sd.append("<v>").append(numVal).append("</v></c>");
            sd.append("</row>");
        }

        sd.append("</sheetData>");

        return (before + sd.toString() + after).getBytes(java.nio.charset.StandardCharsets.UTF_8);
    }

    private byte[] updateTableRange(byte[] original, int dataRows) {
        String xml = new String(original, java.nio.charset.StandardCharsets.UTF_8);
        xml = xml.replaceFirst("ref=\"[^\"]+\"", "ref=\"A1:B" + (dataRows + 1) + "\"");
        return xml.getBytes(java.nio.charset.StandardCharsets.UTF_8);
    }

    // ════════════════════════════════════════════════
    // XML CACHE UPDATE
    // ════════════════════════════════════════════════

    private void updateXmlCache(XSLFChart chart, List<Map<String, Object>> rows) {
        CTPlotArea plotArea = chart.getCTChart().getPlotArea();

        for (int ci = 0; ci < plotArea.sizeOfBarChartArray(); ci++) {
            CTBarSer[] series = plotArea.getBarChartArray(ci).getSerArray();
            for (int si = 0; si < series.length; si++) {
                if (si == 0 && series[si].getCat() != null) updateCatCache(series[si].getCat(), rows);
                if (series[si].getVal() != null) updateValCache(series[si].getVal(), rows, si == 0 ? "value" : "value2");
            }
        }
        for (int ci = 0; ci < plotArea.sizeOfLineChartArray(); ci++) {
            CTLineSer[] series = plotArea.getLineChartArray(ci).getSerArray();
            if (series.length > 0) {
                if (series[0].getCat() != null) updateCatCache(series[0].getCat(), rows);
                if (series[0].getVal() != null) updateValCache(series[0].getVal(), rows, "value");
            }
        }
        for (int ci = 0; ci < plotArea.sizeOfPieChartArray(); ci++) {
            CTPieSer[] series = plotArea.getPieChartArray(ci).getSerArray();
            if (series.length > 0) {
                if (series[0].getCat() != null) updateCatCache(series[0].getCat(), rows);
                if (series[0].getVal() != null) updateValCache(series[0].getVal(), rows, "value");
            }
        }
    }

    private void updateCatCache(CTAxDataSource cat, List<Map<String, Object>> rows) {
        try {
            if (cat.getStrRef() != null) {
                CTStrRef ref = cat.getStrRef();
                ref.setF("Sheet1!$A$2:$A$" + (rows.size() + 1));
                CTStrData cache = ref.isSetStrCache() ? ref.getStrCache() : ref.addNewStrCache();
                cache.setPtArray(new CTStrVal[0]);
                cache.getPtCount().setVal(rows.size());
                for (int i = 0; i < rows.size(); i++) {
                    CTStrVal pt = cache.addNewPt();
                    pt.setIdx(i);
                    pt.setV(rows.get(i).get("label") != null ? rows.get(i).get("label").toString() : "");
                }
            }
        } catch (Exception e) { log.debug("Cat cache skip: {}", e.getMessage()); }
    }

    private void updateValCache(CTNumDataSource val, List<Map<String, Object>> rows, String key) {
        try {
            if (val.getNumRef() != null) {
                CTNumRef ref = val.getNumRef();
                ref.setF("Sheet1!$B$2:$B$" + (rows.size() + 1));
                CTNumData cache = ref.isSetNumCache() ? ref.getNumCache() : ref.addNewNumCache();
                cache.setPtArray(new CTNumVal[0]);
                cache.getPtCount().setVal(rows.size());
                for (int i = 0; i < rows.size(); i++) {
                    CTNumVal pt = cache.addNewPt();
                    pt.setIdx(i);
                    Object v = rows.get(i).get(key);
                    pt.setV(v instanceof Number ? String.valueOf(((Number) v).doubleValue()) : "0");
                }
            }
        } catch (Exception e) { log.debug("Val cache skip: {}", e.getMessage()); }
    }

    // ════════════════════════════════════════════════
    // HELPERS
    // ════════════════════════════════════════════════

    private String extractAltText(XSLFShape shape) {
        try { Matcher m = ALT_TEXT_PAT.matcher(shape.getXmlObject().toString()); if (m.find()) return m.group(1); } catch (Exception e) {}
        return "";
    }

    private String extractChartRelId(XSLFShape shape) {
        try { String x = shape.getXmlObject().toString(); if (x.contains("chart")) { Matcher m = REL_ID_PAT.matcher(x); if (m.find()) return m.group(1); } } catch (Exception e) {}
        return null;
    }

    private String escapeXml(String s) {
        return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;");
    }
}
