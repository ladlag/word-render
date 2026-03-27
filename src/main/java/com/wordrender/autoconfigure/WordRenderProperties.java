package com.wordrender.autoconfigure;

import com.wordrender.model.WordRenderCacheType;
import com.wordrender.model.WordRenderPageSize;
import com.wordrender.model.WordRenderReportStyle;
import org.springframework.boot.context.properties.ConfigurationProperties;

@ConfigurationProperties(prefix = "word-render")
public class WordRenderProperties {

    private boolean enabled = true;
    private WordRenderReportStyle defaultStyle = WordRenderReportStyle.DEFAULT;
    private String defaultFontFamily = "Microsoft YaHei";
    private int defaultFontSize = 11;
    private WordRenderPageSize pageSize = WordRenderPageSize.A4;
    private Fonts fonts = new Fonts();
    private Colors colors = new Colors();
    private Watermark watermark = new Watermark();
    private Cache cache = new Cache();

    public boolean isEnabled() {
        return enabled;
    }

    public void setEnabled(boolean enabled) {
        this.enabled = enabled;
    }

    public WordRenderReportStyle getDefaultStyle() {
        return defaultStyle;
    }

    public void setDefaultStyle(WordRenderReportStyle defaultStyle) {
        this.defaultStyle = defaultStyle;
    }

    public String getDefaultFontFamily() {
        return defaultFontFamily;
    }

    public void setDefaultFontFamily(String defaultFontFamily) {
        this.defaultFontFamily = defaultFontFamily;
    }

    public int getDefaultFontSize() {
        return defaultFontSize;
    }

    public void setDefaultFontSize(int defaultFontSize) {
        this.defaultFontSize = defaultFontSize;
    }

    public WordRenderPageSize getPageSize() {
        return pageSize;
    }

    public void setPageSize(WordRenderPageSize pageSize) {
        this.pageSize = pageSize;
    }

    public Fonts getFonts() {
        return fonts;
    }

    public void setFonts(Fonts fonts) {
        this.fonts = fonts;
    }

    public Colors getColors() {
        return colors;
    }

    public void setColors(Colors colors) {
        this.colors = colors;
    }

    public Watermark getWatermark() {
        return watermark;
    }

    public void setWatermark(Watermark watermark) {
        this.watermark = watermark;
    }

    public Cache getCache() {
        return cache;
    }

    public void setCache(Cache cache) {
        this.cache = cache;
    }

    public String resolveDefaultFontFamily() {
        if (fonts != null && fonts.getPrimaryFamily() != null && !fonts.getPrimaryFamily().trim().isEmpty()) {
            return fonts.getPrimaryFamily().trim();
        }
        if (defaultFontFamily != null && !defaultFontFamily.trim().isEmpty()) {
            return defaultFontFamily.trim();
        }
        return "Noto Sans SC";
    }

    public static class Fonts {
        private String regularPath;
        private String boldPath;
        private String defaultFamily = "Microsoft YaHei, PingFang SC, Noto Sans SC, SimSun, Arial Unicode MS, sans-serif";

        public String getRegularPath() {
            return regularPath;
        }

        public void setRegularPath(String regularPath) {
            this.regularPath = regularPath;
        }

        public String getBoldPath() {
            return boldPath;
        }

        public void setBoldPath(String boldPath) {
            this.boldPath = boldPath;
        }

        public String getDefaultFamily() {
            return defaultFamily;
        }

        public void setDefaultFamily(String defaultFamily) {
            this.defaultFamily = defaultFamily;
        }

        public String getPrimaryFamily() {
            if (defaultFamily == null || defaultFamily.trim().isEmpty()) {
                return null;
            }
            String[] families = defaultFamily.split(",");
            return families.length == 0 ? null : families[0].trim();
        }
    }

    public static class Colors {
        private String titleColor = "1F1F1F";
        private String headingColor = "1F1F1F";
        private String accentColor = "7F7F7F";

        public String getTitleColor() {
            return titleColor;
        }

        public void setTitleColor(String titleColor) {
            this.titleColor = titleColor;
        }

        public String getHeadingColor() {
            return headingColor;
        }

        public void setHeadingColor(String headingColor) {
            this.headingColor = headingColor;
        }

        public String getAccentColor() {
            return accentColor;
        }

        public void setAccentColor(String accentColor) {
            this.accentColor = accentColor;
        }
    }

    public static class Watermark {
        private boolean enabled = false;
        private String text;
        private String color = "D9D9D9";
        private int fontSize = 18;

        public boolean isEnabled() {
            return enabled;
        }

        public void setEnabled(boolean enabled) {
            this.enabled = enabled;
        }

        public String getText() {
            return text;
        }

        public void setText(String text) {
            this.text = text;
        }

        public String getColor() {
            return color;
        }

        public void setColor(String color) {
            this.color = color;
        }

        public int getFontSize() {
            return fontSize;
        }

        public void setFontSize(int fontSize) {
            this.fontSize = fontSize;
        }
    }

    public static class Cache {
        private boolean enabled = false;
        private WordRenderCacheType type = WordRenderCacheType.MEMORY;
        private String redisKeyPrefix = "word-render:";

        public boolean isEnabled() {
            return enabled;
        }

        public void setEnabled(boolean enabled) {
            this.enabled = enabled;
        }

        public WordRenderCacheType getType() {
            return type;
        }

        public void setType(WordRenderCacheType type) {
            this.type = type;
        }

        public String getRedisKeyPrefix() {
            return redisKeyPrefix;
        }

        public void setRedisKeyPrefix(String redisKeyPrefix) {
            this.redisKeyPrefix = redisKeyPrefix;
        }
    }
}
