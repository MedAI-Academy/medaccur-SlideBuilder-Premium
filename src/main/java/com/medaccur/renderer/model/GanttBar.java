package com.medaccur.renderer.model;

public class GanttBar {
    private int workstreamIndex;
    private String label;
    private int startMonth;  // 1-12
    private int endMonth;    // 1-12
    private String color;    // hex e.g. "#8B5CF6"
    
    public int getWorkstreamIndex() { return workstreamIndex; }
    public void setWorkstreamIndex(int i) { this.workstreamIndex = i; }
    public String getLabel() { return label; }
    public void setLabel(String l) { this.label = l; }
    public int getStartMonth() { return startMonth; }
    public void setStartMonth(int m) { this.startMonth = m; }
    public int getEndMonth() { return endMonth; }
    public void setEndMonth(int m) { this.endMonth = m; }
    public String getColor() { return color; }
    public void setColor(String c) { this.color = c; }
}
