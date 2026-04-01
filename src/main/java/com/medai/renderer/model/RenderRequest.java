package com.medai.renderer.model;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import java.util.List;
import java.util.Map;

@JsonIgnoreProperties(ignoreUnknown = true)
public class RenderRequest {

    private String module;          // "map", "slides", "narrative", "clinical"
    private String theme = "dark";  // "dark" or "light"
    private boolean widescreen = true;
    private Map<String, Object> metadata;
    private ConfidenceScore confidenceScore;
    private List<SlideData> slides;

    // --- Getters & Setters ---
    public String getModule() { return module; }
    public void setModule(String module) { this.module = module; }

    public String getTheme() { return theme; }
    public void setTheme(String theme) { this.theme = theme; }

    public boolean isWidescreen() { return widescreen; }
    public void setWidescreen(boolean widescreen) { this.widescreen = widescreen; }

    public Map<String, Object> getMetadata() { return metadata; }
    public void setMetadata(Map<String, Object> metadata) { this.metadata = metadata; }

    public ConfidenceScore getConfidenceScore() { return confidenceScore; }
    public void setConfidenceScore(ConfidenceScore confidenceScore) { this.confidenceScore = confidenceScore; }

    public List<SlideData> getSlides() { return slides; }
    public void setSlides(List<SlideData> slides) { this.slides = slides; }

    // Helper: get metadata value as String
    public String meta(String key) {
        if (metadata == null) return "";
        Object val = metadata.get(key);
        return val != null ? val.toString() : "";
    }
}
