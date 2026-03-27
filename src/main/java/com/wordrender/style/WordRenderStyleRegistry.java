package com.wordrender.style;

import com.wordrender.autoconfigure.WordRenderProperties;
import com.wordrender.model.WordRenderReportStyle;
import com.wordrender.support.WordRenderFontResolver;
import java.util.EnumMap;
import java.util.Map;

public class WordRenderStyleRegistry {

    private final Map<WordRenderReportStyle, WordRenderStyleDefinition> styles =
        new EnumMap<WordRenderReportStyle, WordRenderStyleDefinition>(WordRenderReportStyle.class);
    private final WordRenderFontResolver fontResolver;

    public WordRenderStyleRegistry(WordRenderProperties properties, WordRenderFontResolver fontResolver) {
        this.fontResolver = fontResolver;
        String defaultFontFamily = fontResolver.getDefaultFontFamily();
        String defaultTitleColor = normalizeColor(properties.getColors().getTitleColor(), "1F1F1F");
        String defaultHeadingColor = normalizeColor(properties.getColors().getHeadingColor(), defaultTitleColor);
        String defaultAccentColor = normalizeColor(properties.getColors().getAccentColor(), "7F7F7F");
        styles.put(WordRenderReportStyle.DEFAULT, new WordRenderStyleDefinition(
            defaultFontFamily, properties.getDefaultFontSize(), defaultTitleColor, defaultHeadingColor, defaultAccentColor,
            24, 18, 15, 13, 12, 11, 11, 1440, 1440, 1200, 1200
        ));
        styles.put(WordRenderReportStyle.FORMAL, new WordRenderStyleDefinition(
            defaultFontFamily, properties.getDefaultFontSize(), defaultTitleColor, defaultHeadingColor, defaultAccentColor,
            26, 18, 16, 14, 13, 12, 12, 1700, 1700, 1500, 1500
        ));
        styles.put(WordRenderReportStyle.SIMPLE, new WordRenderStyleDefinition(
            defaultFontFamily, properties.getDefaultFontSize(), defaultTitleColor, defaultHeadingColor, defaultAccentColor,
            22, 16, 14, 12, 11, 10, 10, 1200, 1200, 1000, 1000
        ));
    }

    public WordRenderStyleDefinition getStyle(WordRenderReportStyle reportStyle) {
        WordRenderStyleDefinition definition = styles.get(reportStyle);
        if (definition == null) {
            return styles.get(WordRenderReportStyle.DEFAULT);
        }
        return definition;
    }

    public WordRenderStyleDefinition applyOverrides(WordRenderStyleDefinition baseStyle, String titleColor,
                                                    String headingColor, String accentColor) {
        if (baseStyle == null) {
            return null;
        }
        String resolvedTitleColor = normalizeColor(titleColor, baseStyle.getTitleColor());
        String resolvedHeadingColor = normalizeColor(headingColor, baseStyle.getHeadingColor());
        String resolvedAccentColor = normalizeColor(accentColor, baseStyle.getAccentColor());
        return new WordRenderStyleDefinition(
            baseStyle.getFontFamily(),
            baseStyle.getBodyFontSize(),
            resolvedTitleColor,
            resolvedHeadingColor,
            resolvedAccentColor,
            baseStyle.getCoverTitleFontSize(),
            baseStyle.getHeadingOneFontSize(),
            baseStyle.getHeadingTwoFontSize(),
            baseStyle.getHeadingThreeFontSize(),
            baseStyle.getHeadingFourFontSize(),
            baseStyle.getHeadingFiveFontSize(),
            baseStyle.getHeadingSixFontSize(),
            baseStyle.getLeftMarginTwips(),
            baseStyle.getRightMarginTwips(),
            baseStyle.getTopMarginTwips(),
            baseStyle.getBottomMarginTwips()
        );
    }

    public String getDefaultFontFamily() {
        return fontResolver.getDefaultFontFamily();
    }

    private String normalizeColor(String color, String fallback) {
        if (color == null || color.trim().isEmpty()) {
            return fallback;
        }
        String normalized = color.trim().replace("#", "").toUpperCase();
        if (!normalized.matches("[0-9A-F]{6}")) {
            throw new IllegalArgumentException("Invalid color value: " + color);
        }
        return normalized;
    }
}
