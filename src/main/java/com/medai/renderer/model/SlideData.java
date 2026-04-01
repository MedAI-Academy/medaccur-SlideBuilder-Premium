package com.medai.renderer.model;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import java.util.List;
import java.util.Map;

@JsonIgnoreProperties(ignoreUnknown = true)
public class SlideData {

    private String id;              // Unique slide identifier
    private String layout;          // TITLE, TOC, DIVIDER, CONTENT_FULL, TABLE, CHART_KM, SWOT, TIMELINE, etc.
    private String section;         // Section name for grouping
    private int sectionIndex;       // For divider numbering
    private Map<String, Object> content;  // Layout-specific content

    // --- Getters & Setters ---
    public String getId() { return id; }
    public void setId(String id) { this.id = id; }

    public String getLayout() { return layout; }
    public void setLayout(String layout) { this.layout = layout; }

    public String getSection() { return section; }
    public void setSection(String section) { this.section = section; }

    public int getSectionIndex() { return sectionIndex; }
    public void setSectionIndex(int sectionIndex) { this.sectionIndex = sectionIndex; }

    public Map<String, Object> getContent() { return content; }
    public void setContent(Map<String, Object> content) { this.content = content; }

    // --- Content helpers ---
    public String contentStr(String key) {
        if (content == null) return "";
        Object val = content.get(key);
        return val != null ? val.toString() : "";
    }

    @SuppressWarnings("unchecked")
    public List<Map<String, Object>> contentList(String key) {
        if (content == null) return List.of();
        Object val = content.get(key);
        if (val instanceof List) return (List<Map<String, Object>>) val;
        return List.of();
    }

    @SuppressWarnings("unchecked")
    public Map<String, Object> contentMap(String key) {
        if (content == null) return Map.of();
        Object val = content.get(key);
        if (val instanceof Map) return (Map<String, Object>) val;
        return Map.of();
    }
}
