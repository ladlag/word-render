package com.wordrender.style;

import java.util.Collections;
import java.util.HashMap;
import java.util.Map;

public class WordRenderStyleDefinition {

    private final String fontFamily;
    private final int bodyFontSize;
    private final String titleColor;
    private final String headingColor;
    private final String accentColor;
    private final int coverTitleFontSize;
    private final int headingOneFontSize;
    private final int headingTwoFontSize;
    private final int headingThreeFontSize;
    private final int headingFourFontSize;
    private final int headingFiveFontSize;
    private final int headingSixFontSize;
    private final int leftMarginTwips;
    private final int rightMarginTwips;
    private final int topMarginTwips;
    private final int bottomMarginTwips;
    private final Map<Integer, String> headingStyleIds;

    public WordRenderStyleDefinition(String fontFamily, int bodyFontSize, String titleColor, String headingColor,
                                     String accentColor,
                                     int coverTitleFontSize, int headingOneFontSize, int headingTwoFontSize,
                                     int headingThreeFontSize, int headingFourFontSize, int headingFiveFontSize,
                                     int headingSixFontSize, int leftMarginTwips,
                                     int rightMarginTwips, int topMarginTwips, int bottomMarginTwips) {
        this(fontFamily, bodyFontSize, titleColor, headingColor, accentColor, coverTitleFontSize, headingOneFontSize,
            headingTwoFontSize, headingThreeFontSize, headingFourFontSize, headingFiveFontSize, headingSixFontSize,
            leftMarginTwips, rightMarginTwips, topMarginTwips, bottomMarginTwips, null);
    }

    public WordRenderStyleDefinition(String fontFamily, int bodyFontSize, String titleColor, String headingColor,
                                     String accentColor,
                                     int coverTitleFontSize, int headingOneFontSize, int headingTwoFontSize,
                                     int headingThreeFontSize, int headingFourFontSize, int headingFiveFontSize,
                                     int headingSixFontSize, int leftMarginTwips,
                                     int rightMarginTwips, int topMarginTwips, int bottomMarginTwips,
                                     Map<Integer, String> headingStyleIds) {
        this.fontFamily = fontFamily;
        this.bodyFontSize = bodyFontSize;
        this.titleColor = titleColor;
        this.headingColor = headingColor;
        this.accentColor = accentColor;
        this.coverTitleFontSize = coverTitleFontSize;
        this.headingOneFontSize = headingOneFontSize;
        this.headingTwoFontSize = headingTwoFontSize;
        this.headingThreeFontSize = headingThreeFontSize;
        this.headingFourFontSize = headingFourFontSize;
        this.headingFiveFontSize = headingFiveFontSize;
        this.headingSixFontSize = headingSixFontSize;
        this.leftMarginTwips = leftMarginTwips;
        this.rightMarginTwips = rightMarginTwips;
        this.topMarginTwips = topMarginTwips;
        this.bottomMarginTwips = bottomMarginTwips;
        this.headingStyleIds = headingStyleIds == null
            ? Collections.<Integer, String>emptyMap()
            : Collections.unmodifiableMap(new HashMap<Integer, String>(headingStyleIds));
    }

    public String getFontFamily() {
        return fontFamily;
    }

    public int getBodyFontSize() {
        return bodyFontSize;
    }

    public String getTitleColor() {
        return titleColor;
    }

    public String getHeadingColor() {
        return headingColor;
    }

    public String getAccentColor() {
        return accentColor;
    }

    public int getCoverTitleFontSize() {
        return coverTitleFontSize;
    }

    public int getHeadingOneFontSize() {
        return headingOneFontSize;
    }

    public int getHeadingTwoFontSize() {
        return headingTwoFontSize;
    }

    public int getHeadingThreeFontSize() {
        return headingThreeFontSize;
    }

    public int getHeadingFourFontSize() {
        return headingFourFontSize;
    }

    public int getHeadingFiveFontSize() {
        return headingFiveFontSize;
    }

    public int getHeadingSixFontSize() {
        return headingSixFontSize;
    }

    public int getLeftMarginTwips() {
        return leftMarginTwips;
    }

    public int getRightMarginTwips() {
        return rightMarginTwips;
    }

    public int getTopMarginTwips() {
        return topMarginTwips;
    }

    public int getBottomMarginTwips() {
        return bottomMarginTwips;
    }

    public String resolveHeadingStyleId(int level) {
        String mapped = headingStyleIds.get(Integer.valueOf(level));
        return mapped != null ? mapped : "Heading" + level;
    }

    public Map<Integer, String> getHeadingStyleIds() {
        return headingStyleIds;
    }

    public WordRenderStyleDefinition withFontFamily(String resolvedFontFamily) {
        return new WordRenderStyleDefinition(
            resolvedFontFamily,
            bodyFontSize,
            titleColor,
            headingColor,
            accentColor,
            coverTitleFontSize,
            headingOneFontSize,
            headingTwoFontSize,
            headingThreeFontSize,
            headingFourFontSize,
            headingFiveFontSize,
            headingSixFontSize,
            leftMarginTwips,
            rightMarginTwips,
            topMarginTwips,
            bottomMarginTwips,
            headingStyleIds
        );
    }

    public WordRenderStyleDefinition withHeadingStyleIds(Map<Integer, String> resolvedHeadingStyleIds) {
        return new WordRenderStyleDefinition(
            fontFamily,
            bodyFontSize,
            titleColor,
            headingColor,
            accentColor,
            coverTitleFontSize,
            headingOneFontSize,
            headingTwoFontSize,
            headingThreeFontSize,
            headingFourFontSize,
            headingFiveFontSize,
            headingSixFontSize,
            leftMarginTwips,
            rightMarginTwips,
            topMarginTwips,
            bottomMarginTwips,
            resolvedHeadingStyleIds
        );
    }
}
