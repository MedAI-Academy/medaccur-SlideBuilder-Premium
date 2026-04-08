package com.medaccur.renderer.model;

import java.util.List;
import java.util.Map;

public class SlideSpec {
    private String layout;                    // Layout name in master template
    private Map<String, String> content;      // {{key}} → value mappings
    private String figurePng;                 // Base64-encoded PNG (for CHART_AREA)
    private List<GanttBar> ganttBars;         // Native shape bars
    private List<ComparisonBar> comparisonBars;
    private Map<String, Object> chartData;    // Data for Chart Microservice call

    public String getLayout() { return layout; }
    public void setLayout(String layout) { this.layout = layout; }
    public Map<String, String> getContent() { return content; }
    public void setContent(Map<String, String> content) { this.content = content; }
    public String getFigurePng() { return figurePng; }
    public void setFigurePng(String figurePng) { this.figurePng = figurePng; }
    public List<GanttBar> getGanttBars() { return ganttBars; }
    public void setGanttBars(List<GanttBar> ganttBars) { this.ganttBars = ganttBars; }
    public List<ComparisonBar> getComparisonBars() { return comparisonBars; }
    public void setComparisonBars(List<ComparisonBar> comparisonBars) { this.comparisonBars = comparisonBars; }
    public Map<String, Object> getChartData() { return chartData; }
    public void setChartData(Map<String, Object> chartData) { this.chartData = chartData; }

    public boolean hasFigure() { return figurePng != null && !figurePng.isEmpty(); }
    public boolean hasGanttBars() { return ganttBars != null && !ganttBars.isEmpty(); }
    public boolean hasComparisonBars() { return comparisonBars != null && !comparisonBars.isEmpty(); }
    public boolean hasChartData() { return chartData != null && !chartData.isEmpty(); }
}
