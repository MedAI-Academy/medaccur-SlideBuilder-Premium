package com.medaccur.renderer.service;

import org.apache.poi.xslf.usermodel.*;
import org.apache.poi.sl.usermodel.TextShape.TextAutofit;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;

@Service
public class PlaceholderService {

    private static final Logger log = LoggerFactory.getLogger(PlaceholderService.class);
    private static final Pattern PLACEHOLDER = Pattern.compile("\\{\\{[a-zA-Z0-9_]+\\}\\}");

    /**
     * Replace all {{key}} placeholders in a slide with values from content map.
     * Preserves font, size, color, bold — only changes text.
     *
     * CRITICAL: PowerPoint splits text across multiple XSLFTextRun objects unpredictably.
     * We must join all runs to find {{key}}, then put result in first run and clear the rest.
     *
     * CRITICAL: Line break runs (<a:br/>) CANNOT have their text changed — POI throws
     * "You cannot change text of a line break". We must skip them.
     */
    public int replacePlaceholders(XSLFSlide slide, Map<String, String> content) {
        int count = 0;
        for (XSLFShape shape : slide.getShapes()) {
            if (shape instanceof XSLFTextShape) {
                count += replaceInTextShape((XSLFTextShape) shape, content);
            } else if (shape instanceof XSLFGroupShape) {
                for (XSLFShape child : ((XSLFGroupShape) shape).getShapes()) {
                    if (child instanceof XSLFTextShape) {
                        count += replaceInTextShape((XSLFTextShape) child, content);
                    }
                }
            } else if (shape instanceof XSLFTable) {
                // BUG 1 FIX: scan table cells for {{placeholders}}
                count += replaceInTable((XSLFTable) shape, content);
            }
        }
        return count;
    }

    private int replaceInTable(XSLFTable table, Map<String, String> content) {
        int count = 0;
        for (XSLFTableRow row : table.getRows()) {
            for (XSLFTableCell cell : row.getCells()) {
                for (XSLFTextParagraph para : cell.getTextParagraphs()) {
                    String fullText = para.getText();
                    if (fullText == null || !fullText.contains("{{")) continue;

                    String newText = fullText;
                    for (Map.Entry<String, String> e : content.entrySet()) {
                        String placeholder = "{{" + e.getKey() + "}}";
                        if (newText.contains(placeholder)) {
                            String value = e.getValue() != null ? e.getValue() : "";
                            newText = newText.replace(placeholder, value);
                            count++;
                        }
                    }

                    if (!newText.equals(fullText)) {
                        applyTextToRuns(para, newText);
                    }
                }
            }
        }
        return count;
    }

    private int replaceInTextShape(XSLFTextShape textShape, Map<String, String> content) {
        int count = 0;
        for (XSLFTextParagraph para : textShape.getTextParagraphs()) {
            String fullText = para.getText();
            if (fullText == null || !fullText.contains("{{")) continue;

            String newText = fullText;
            for (Map.Entry<String, String> e : content.entrySet()) {
                String placeholder = "{{" + e.getKey() + "}}";
                if (newText.contains(placeholder)) {
                    String value = e.getValue() != null ? e.getValue() : "";
                    newText = newText.replace(placeholder, value);
                    count++;
                }
            }

            if (!newText.equals(fullText)) {
                applyTextToRuns(para, newText);
            }
        }
        return count;
    }

    /**
     * Remove any remaining {{placeholder}} text (replace with empty string).
     * Called AFTER replacePlaceholders to clean up unfilled slots.
     */
    public int removeUnfilledPlaceholders(XSLFSlide slide) {
        int count = 0;
        for (XSLFShape shape : slide.getShapes()) {
            if (shape instanceof XSLFTextShape) {
                count += removeUnfilledInTextShape((XSLFTextShape) shape);
            } else if (shape instanceof XSLFGroupShape) {
                for (XSLFShape child : ((XSLFGroupShape) shape).getShapes()) {
                    if (child instanceof XSLFTextShape) {
                        count += removeUnfilledInTextShape((XSLFTextShape) child);
                    }
                }
            } else if (shape instanceof XSLFTable) {
                // BUG 1 FIX: also clean unfilled {{}} in table cells
                count += removeUnfilledInTable((XSLFTable) shape);
            }
        }
        return count;
    }

    private int removeUnfilledInTable(XSLFTable table) {
        int count = 0;
        for (XSLFTableRow row : table.getRows()) {
            for (XSLFTableCell cell : row.getCells()) {
                for (XSLFTextParagraph para : cell.getTextParagraphs()) {
                    String text = para.getText();
                    if (text == null || !text.contains("{{")) continue;

                    String cleaned = PLACEHOLDER.matcher(text).replaceAll("");
                    if (!cleaned.equals(text)) {
                        applyTextToRuns(para, cleaned.trim());
                        count++;
                    }
                }
            }
        }
        return count;
    }

    private int removeUnfilledInTextShape(XSLFTextShape textShape) {
        int count = 0;
        for (XSLFTextParagraph para : textShape.getTextParagraphs()) {
            String text = para.getText();
            if (text == null || !text.contains("{{")) continue;

            String cleaned = PLACEHOLDER.matcher(text).replaceAll("");
            if (!cleaned.equals(text)) {
                applyTextToRuns(para, cleaned.trim());
                count++;
            }
        }
        return count;
    }

    /**
     * Apply new text to a paragraph's runs, skipping line break runs.
     * Puts all text in the first editable run and clears the rest.
     * Preserves formatting of the first run.
     */
    private void applyTextToRuns(XSLFTextParagraph para, String newText) {
        List<XSLFTextRun> runs = para.getTextRuns();
        if (runs.isEmpty()) return;

        boolean firstSet = false;
        for (XSLFTextRun run : runs) {
            // Skip line break runs — POI throws if you try to setText on them
            if (isLineBreakRun(run)) continue;

            if (!firstSet) {
                run.setText(newText);
                firstSet = true;
            } else {
                run.setText("");
            }
        }
    }

    /**
     * Check if a run is a line break (<a:br/> element).
     * Line breaks always have text "\n" and cannot be modified.
     */
    private boolean isLineBreakRun(XSLFTextRun run) {
        try {
            String raw = run.getRawText();
            if ("\n".equals(raw)) return true;
        } catch (Exception e) {
            // If we can't read it, assume it's safe
        }
        // Also check by trying — if the XML element is a br, the class name hints at it
        try {
            String xmlName = run.getXmlObject().getDomNode().getLocalName();
            if ("br".equals(xmlName)) return true;
        } catch (Exception e) {
            // ignore
        }
        return false;
    }

    /**
     * Enable auto-shrink on all text shapes.
     * This prevents text overflow — PowerPoint automatically reduces font size.
     */
    public void enableAutoShrink(XSLFSlide slide) {
        for (XSLFShape shape : slide.getShapes()) {
            if (shape instanceof XSLFTextShape) {
                try {
                    ((XSLFTextShape) shape).setTextAutofit(TextAutofit.NORMAL);
                } catch (Exception e) {
                    // Some shapes don't support autofit — ignore
                }
            }
        }
    }
}
