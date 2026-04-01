package com.medai.renderer.service;

import com.medai.renderer.template.ThemeConfig;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.axis.NumberAxis;
import org.jfree.chart.plot.XYPlot;
import org.jfree.chart.renderer.xy.XYStepRenderer;
import org.jfree.chart.title.LegendTitle;
import org.jfree.data.xy.XYSeries;
import org.jfree.data.xy.XYSeriesCollection;
import org.springframework.stereotype.Service;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.text.NumberFormat;
import java.util.List;
import java.util.Map;

/**
 * Generates high-resolution chart images (PNG) using JFreeChart.
 * These are embedded into PPTX slides as pictures.
 *
 * Supported chart types:
 * - Kaplan-Meier survival curves (step function)
 * - Bar charts (efficacy comparisons)
 * - SWOT matrix (as a styled 2x2 grid image)
 */
@Service
public class ChartService {

    private static final int CHART_W = 1600;  // px — high-res for crisp PPTX
    private static final int CHART_H = 900;

    /**
     * Generate a Kaplan-Meier survival curve as PNG byte array.
     *
     * @param chartData Map containing:
     *   - arms: List of {name, color, timepoints[], survival[], atRisk[], median}
     *   - xlabel, ylabel, hazardRatio
     */
    @SuppressWarnings("unchecked")
    public byte[] generateKaplanMeier(Map<String, Object> chartData) throws Exception {
        List<Map<String, Object>> arms = (List<Map<String, Object>>) chartData.get("arms");
        String xlabel = (String) chartData.getOrDefault("xlabel", "Time (months)");
        String ylabel = (String) chartData.getOrDefault("ylabel", "Survival Probability");
        String hrText = (String) chartData.getOrDefault("hazardRatio", "");

        XYSeriesCollection dataset = new XYSeriesCollection();

        for (Map<String, Object> arm : arms) {
            String name = (String) arm.get("name");
            List<Number> timepoints = (List<Number>) arm.get("timepoints");
            List<Number> survival = (List<Number>) arm.get("survival");
            Number median = (Number) arm.get("median");

            XYSeries series = new XYSeries(name + (median != null ?
                " (median: " + median + " mo)" : ""));

            for (int i = 0; i < Math.min(timepoints.size(), survival.size()); i++) {
                series.add(timepoints.get(i).doubleValue(), survival.get(i).doubleValue());
            }
            dataset.addSeries(series);
        }

        // Create step-function chart (Kaplan-Meier style)
        JFreeChart chart = ChartFactory.createXYLineChart(
            null,  // No title (we add it as a PPTX text box)
            xlabel, ylabel, dataset
        );

        // Style the chart for dark theme
        XYPlot plot = chart.getXYPlot();
        plot.setBackgroundPaint(ThemeConfig.CLR_SURFACE);
        plot.setDomainGridlinePaint(new Color(74, 106, 154, 80));  // Subtle grid
        plot.setRangeGridlinePaint(new Color(74, 106, 154, 80));
        plot.setOutlinePaint(ThemeConfig.CLR_SURFACE);

        // Step renderer for KM curves
        XYStepRenderer renderer = new XYStepRenderer();
        for (int i = 0; i < arms.size(); i++) {
            Map<String, Object> arm = arms.get(i);
            String colorHex = ((String) arm.getOrDefault("color", "#7C6FFF")).replace("#", "");
            renderer.setSeriesPaint(i, ThemeConfig.hex(colorHex));
            renderer.setSeriesStroke(i, new BasicStroke(2.5f));
        }
        plot.setRenderer(renderer);

        // Axis styling
        NumberAxis domainAxis = (NumberAxis) plot.getDomainAxis();
        domainAxis.setLabelPaint(ThemeConfig.CLR_MUTED);
        domainAxis.setTickLabelPaint(ThemeConfig.CLR_MUTED);
        domainAxis.setLabelFont(new Font(ThemeConfig.FONT_BODY, Font.PLAIN, 14));

        NumberAxis rangeAxis = (NumberAxis) plot.getRangeAxis();
        rangeAxis.setRange(0.0, 1.05);
        rangeAxis.setNumberFormatOverride(NumberFormat.getPercentInstance());
        rangeAxis.setLabelPaint(ThemeConfig.CLR_MUTED);
        rangeAxis.setTickLabelPaint(ThemeConfig.CLR_MUTED);
        rangeAxis.setLabelFont(new Font(ThemeConfig.FONT_BODY, Font.PLAIN, 14));

        // Legend styling
        chart.setBackgroundPaint(ThemeConfig.CLR_SURFACE);
        LegendTitle legend = chart.getLegend();
        if (legend != null) {
            legend.setBackgroundPaint(ThemeConfig.CLR_SURFACE);
            legend.setItemPaint(ThemeConfig.CLR_TEXT);
            legend.setItemFont(new Font(ThemeConfig.FONT_BODY, Font.PLAIN, 13));
        }

        // Render to PNG
        BufferedImage image = chart.createBufferedImage(CHART_W, CHART_H);

        // Add hazard ratio annotation if provided
        if (hrText != null && !hrText.isEmpty()) {
            Graphics2D g2 = image.createGraphics();
            g2.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING,
                RenderingHints.VALUE_TEXT_ANTIALIAS_ON);
            g2.setFont(new Font(ThemeConfig.FONT_BODY, Font.BOLD, 16));
            g2.setColor(ThemeConfig.CLR_TEAL);
            g2.drawString(hrText, CHART_W - g2.getFontMetrics().stringWidth(hrText) - 20, 30);
            g2.dispose();
        }

        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        ImageIO.write(image, "png", baos);
        return baos.toByteArray();
    }

    /**
     * Generate a SWOT matrix as a styled 2x2 grid image.
     */
    public byte[] generateSwotMatrix(Map<String, Object> swotData) throws Exception {
        int W = 1600, H = 1000;
        BufferedImage image = new BufferedImage(W, H, BufferedImage.TYPE_INT_ARGB);
        Graphics2D g2 = image.createGraphics();
        g2.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_ON);
        g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);

        // Background
        g2.setColor(ThemeConfig.CLR_SURFACE);
        g2.fillRect(0, 0, W, H);

        // Quadrant colors and labels
        Color[] colors = {
            ThemeConfig.hex("1A5276"), ThemeConfig.hex("1A6B4A"),  // S, W (darker)
            ThemeConfig.hex("7D5A2F"), ThemeConfig.hex("7B2D3B")   // O, T (darker)
        };
        String[] labels = {"STRENGTHS", "WEAKNESSES", "OPPORTUNITIES", "THREATS"};
        String[] icons  = {"💪", "⚠️", "🎯", "🔥"};
        String[] keys   = {"strengths", "weaknesses", "opportunities", "threats"};

        int halfW = W / 2 - 8, halfH = H / 2 - 8;
        int[][] positions = {{4, 4}, {W/2 + 4, 4}, {4, H/2 + 4}, {W/2 + 4, H/2 + 4}};

        for (int i = 0; i < 4; i++) {
            int qx = positions[i][0], qy = positions[i][1];
            g2.setColor(colors[i]);
            g2.fillRoundRect(qx, qy, halfW, halfH, 12, 12);

            // Label
            g2.setFont(new Font(ThemeConfig.FONT_TITLE, Font.BOLD, 18));
            g2.setColor(ThemeConfig.CLR_WHITE);
            g2.drawString(icons[i] + "  " + labels[i], qx + 16, qy + 34);

            // Content items
            @SuppressWarnings("unchecked")
            List<String> items = (List<String>) swotData.getOrDefault(keys[i], List.of());
            g2.setFont(new Font(ThemeConfig.FONT_BODY, Font.PLAIN, 13));
            g2.setColor(ThemeConfig.CLR_TEXT);
            int ty = qy + 60;
            for (int j = 0; j < Math.min(items.size(), 6); j++) {
                String item = items.get(j);
                if (item.length() > 80) item = item.substring(0, 77) + "...";
                g2.drawString("• " + item, qx + 20, ty);
                ty += 22;
            }
        }

        g2.dispose();

        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        ImageIO.write(image, "png", baos);
        return baos.toByteArray();
    }
}
