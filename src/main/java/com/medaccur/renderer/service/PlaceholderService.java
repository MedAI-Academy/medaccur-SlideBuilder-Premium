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
     */
    public int replacePlaceholders(XSLFSlide slide, Map<String, String> content) {
        int count = 0;
        for (XSLFShape shape : slide.getShapes()) {
            if (!(shape instanceof XSLFTextShape)) continue;
            XSLFTextShape textShape = (XSLFTextShape) shape;

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
                    // Put new text in first run, clear remaining runs
                    // This preserves the formatting of the first run
                    List<XSLFTextRun> runs = para.getTextRuns();
                    if (!runs.isEmpty()) {
                        runs.get(0).setText(newText);
                        for (int i = 1; i < runs.size(); i++) {
                            runs.get(i).setText("");
                        }
                    }
                }
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
            if (!(shape instanceof XSLFTextShape)) continue;
            XSLFTextShape textShape = (XSLFTextShape) shape;

            for (XSLFTextParagraph para : textShape.getTextParagraphs()) {
                String text = para.getText();
                if (text == null || !text.contains("{{")) continue;

                String cleaned = PLACEHOLDER.matcher(text).replaceAll("");
                if (!cleaned.equals(text)) {
                    List<XSLFTextRun> runs = para.getTextRuns();
                    if (!runs.isEmpty()) {
                        runs.get(0).setText(cleaned.trim());
                        for (int i = 1; i < runs.size(); i++) {
                            runs.get(i).setText("");
                        }
                    }
                    count++;
                }
            }
        }
        return count;
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
