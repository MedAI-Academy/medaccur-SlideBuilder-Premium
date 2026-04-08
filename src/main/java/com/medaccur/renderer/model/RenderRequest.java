package com.medaccur.renderer.model;

import java.util.List;
import java.util.Map;

public class RenderRequest {
    private Map<String, String> metadata;
    private List<SlideSpec> slides;

    public Map<String, String> getMetadata() { return metadata; }
    public void setMetadata(Map<String, String> metadata) { this.metadata = metadata; }
    public List<SlideSpec> getSlides() { return slides; }
    public void setSlides(List<SlideSpec> slides) { this.slides = slides; }
}
