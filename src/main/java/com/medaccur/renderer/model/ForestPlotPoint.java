package com.medaccur.renderer.model;

/**
 * One row in a forest plot: either a data point (with HR + CI) or a header row.
 * Sent as part of the JSON "forestPlotPoints" array in SlideSpec.
 *
 * JSON example:
 * {
 *   "forestPlotPoints": [
 *     {"rowIndex": 0, "hr": 0.64, "ciLower": 0.50, "ciUpper": 0.82, "overall": true},
 *     {"rowIndex": 1, "header": true},
 *     {"rowIndex": 2, "hr": 0.68, "ciLower": 0.46, "ciUpper": 1.02},
 *     {"rowIndex": 3, "hr": 0.62, "ciLower": 0.46, "ciUpper": 0.85}
 *   ]
 * }
 */
public class ForestPlotPoint {

    private int rowIndex;       // 0-based row in the template table (Overall=0, first header=1, etc.)
    private double hr;          // Hazard Ratio point estimate
    private double ciLower;     // Lower bound of 95% CI
    private double ciUpper;     // Upper bound of 95% CI
    private boolean overall;    // true = draw diamond instead of square
    private boolean header;     // true = skip this row (category header, no data)

    public int getRowIndex() { return rowIndex; }
    public void setRowIndex(int rowIndex) { this.rowIndex = rowIndex; }

    public double getHr() { return hr; }
    public void setHr(double hr) { this.hr = hr; }

    public double getCiLower() { return ciLower; }
    public void setCiLower(double ciLower) { this.ciLower = ciLower; }

    public double getCiUpper() { return ciUpper; }
    public void setCiUpper(double ciUpper) { this.ciUpper = ciUpper; }

    public boolean isOverall() { return overall; }
    public void setOverall(boolean overall) { this.overall = overall; }

    public boolean isHeader() { return header; }
    public void setHeader(boolean header) { this.header = header; }
}
