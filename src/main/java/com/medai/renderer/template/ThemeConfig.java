package com.medai.renderer.template;

import java.awt.Color;

/**
 * MedAI Brand Design System — colors, fonts, and spacing constants.
 * All slide builders use this single source of truth.
 */
public class ThemeConfig {

    // ═══════════════════════════════════════════════════════════
    // COLORS (as java.awt.Color for JFreeChart + as hex for POI)
    // ═══════════════════════════════════════════════════════════

    // Primary backgrounds
    public static final String HEX_DARK    = "0B1A3B";  // Title, dividers
    public static final String HEX_NAVY    = "0D2B4E";  // Content dark BG
    public static final String HEX_SURFACE = "163060";  // Cards, boxes
    public static final String HEX_LIGHT   = "F0F4FF";  // Content light BG

    // Accent colors
    public static final String HEX_ACCENT  = "7C6FFF";  // Purple accent
    public static final String HEX_TEAL    = "22D3A5";  // Positive, success
    public static final String HEX_GOLD    = "F5C842";  // Warnings, KPIs
    public static final String HEX_ROSE    = "FF5F7E";  // Negative, alerts
    public static final String HEX_ORANGE  = "FF9F43";  // Secondary accent

    // Text colors
    public static final String HEX_WHITE   = "FFFFFF";
    public static final String HEX_TEXT    = "EAF0FF";  // Primary text on dark
    public static final String HEX_MUTED   = "7B9FD4";  // Secondary text
    public static final String HEX_DIM     = "4A6A9A";  // Tertiary text
    public static final String HEX_INK     = "1E293B";  // Text on light BG
    public static final String HEX_SLATE   = "64748B";  // Muted on light BG

    // java.awt.Color equivalents (for JFreeChart)
    public static final Color CLR_DARK    = hex(HEX_DARK);
    public static final Color CLR_NAVY    = hex(HEX_NAVY);
    public static final Color CLR_SURFACE = hex(HEX_SURFACE);
    public static final Color CLR_ACCENT  = hex(HEX_ACCENT);
    public static final Color CLR_TEAL    = hex(HEX_TEAL);
    public static final Color CLR_GOLD    = hex(HEX_GOLD);
    public static final Color CLR_ROSE    = hex(HEX_ROSE);
    public static final Color CLR_WHITE   = hex(HEX_WHITE);
    public static final Color CLR_TEXT    = hex(HEX_TEXT);
    public static final Color CLR_MUTED   = hex(HEX_MUTED);

    // Accent color array for cycling through charts/items
    public static final String[] ACCENT_CYCLE = {
        HEX_ACCENT, HEX_TEAL, HEX_GOLD, HEX_ROSE,
        HEX_ORANGE, "4DA6FF", "E879F9", "34D399"
    };

    // ═══════════════════════════════════════════════════════════
    // FONTS
    // ═══════════════════════════════════════════════════════════

    public static final String FONT_TITLE = "Calibri";
    public static final String FONT_BODY  = "Calibri";
    public static final String FONT_MONO  = "Consolas";

    // Font sizes (in points)
    public static final double SIZE_TITLE     = 36.0;
    public static final double SIZE_SUBTITLE  = 20.0;
    public static final double SIZE_HEADING   = 18.0;
    public static final double SIZE_BODY      = 13.0;
    public static final double SIZE_SMALL     = 10.0;
    public static final double SIZE_CAPTION   = 8.0;
    public static final double SIZE_FOOTER    = 7.0;

    // ═══════════════════════════════════════════════════════════
    // SLIDE DIMENSIONS (inches — widescreen 13.33" × 7.5")
    // ═══════════════════════════════════════════════════════════

    public static final double SLIDE_W = 13.333;  // Widescreen width
    public static final double SLIDE_H = 7.5;     // Widescreen height

    public static final double MARGIN   = 0.50;  // Outer margin
    public static final double CONTENT_X = 0.50;
    public static final double CONTENT_W = 12.33; // SLIDE_W - 2*MARGIN
    public static final double HEADER_H  = 0.82;  // Header bar height
    public static final double FOOTER_Y  = 7.05;  // Footer Y position
    public static final double FOOTER_H  = 0.35;  // Footer height
    public static final double CONTENT_Y = 1.10;  // Content start Y (below header)
    public static final double CONTENT_H = 5.80;  // Available content height

    // ═══════════════════════════════════════════════════════════
    // HELPERS
    // ═══════════════════════════════════════════════════════════

    /** Convert 6-char hex string to java.awt.Color */
    public static Color hex(String hex) {
        return new Color(
            Integer.parseInt(hex.substring(0, 2), 16),
            Integer.parseInt(hex.substring(2, 4), 16),
            Integer.parseInt(hex.substring(4, 6), 16)
        );
    }

    /** Get background color for a layout type */
    public static String bgColorFor(String layout, String theme) {
        if (layout == null) return "dark".equals(theme) ? HEX_NAVY : HEX_LIGHT;
        return switch (layout) {
            case "TITLE", "DIVIDER", "CONFIDENCE" -> HEX_DARK;
            case "TOC", "REFERENCES" -> HEX_NAVY;
            default -> "dark".equals(theme) ? HEX_NAVY : HEX_LIGHT;
        };
    }

    /** Get text color based on background */
    public static String textColorFor(String bgHex) {
        return switch (bgHex) {
            case "F0F4FF", "FFFFFF" -> HEX_INK;
            default -> HEX_TEXT;
        };
    }
}
