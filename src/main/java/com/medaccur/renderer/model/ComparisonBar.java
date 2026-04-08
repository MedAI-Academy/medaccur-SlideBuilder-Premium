package com.medaccur.renderer.model;

public class ComparisonBar {
    private String label;
    private double value;
    private String color;
    private String unit;  // "mo", "%", etc.
    
    public String getLabel() { return label; }
    public void setLabel(String l) { this.label = l; }
    public double getValue() { return value; }
    public void setValue(double v) { this.value = v; }
    public String getColor() { return color; }
    public void setColor(String c) { this.color = c; }
    public String getUnit() { return unit; }
    public void setUnit(String u) { this.unit = u; }
}
