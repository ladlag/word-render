package com.wordrender.cache;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.wordrender.style.WordRenderStyleDefinition;
import org.springframework.data.redis.core.StringRedisTemplate;

public class WordRenderRedisCacheService implements WordRenderCacheService {

    private final StringRedisTemplate redisTemplate;
    private final ObjectMapper objectMapper;
    private final String prefix;

    public WordRenderRedisCacheService(StringRedisTemplate redisTemplate, ObjectMapper objectMapper, String prefix) {
        this.redisTemplate = redisTemplate;
        this.objectMapper = objectMapper;
        this.prefix = prefix;
    }

    @Override
    public WordRenderStyleDefinition getStyle(String key) {
        try {
            String json = redisTemplate.opsForValue().get(prefix + key);
            if (json == null) {
                return null;
            }
            return objectMapper.readValue(json, WordRenderStyleDefinitionMixin.class).toDefinition();
        } catch (Exception ex) {
            return null;
        }
    }

    @Override
    public void putStyle(String key, WordRenderStyleDefinition definition) {
        try {
            redisTemplate.opsForValue().set(prefix + key, objectMapper.writeValueAsString(WordRenderStyleDefinitionMixin.from(definition)));
        } catch (Exception ex) {
            // degrade gracefully when redis is unavailable
        }
    }

    private static class WordRenderStyleDefinitionMixin {
        public String fontFamily;
        public int bodyFontSize;
        public String titleColor;
        public String headingColor;
        public String accentColor;
        public int coverTitleFontSize;
        public int headingOneFontSize;
        public int headingTwoFontSize;
        public int headingThreeFontSize;
        public int headingFourFontSize;
        public int headingFiveFontSize;
        public int headingSixFontSize;
        public int leftMarginTwips;
        public int rightMarginTwips;
        public int topMarginTwips;
        public int bottomMarginTwips;

        static WordRenderStyleDefinitionMixin from(WordRenderStyleDefinition definition) {
            WordRenderStyleDefinitionMixin mixin = new WordRenderStyleDefinitionMixin();
            mixin.fontFamily = definition.getFontFamily();
            mixin.bodyFontSize = definition.getBodyFontSize();
            mixin.titleColor = definition.getTitleColor();
            mixin.headingColor = definition.getHeadingColor();
            mixin.accentColor = definition.getAccentColor();
            mixin.coverTitleFontSize = definition.getCoverTitleFontSize();
            mixin.headingOneFontSize = definition.getHeadingOneFontSize();
            mixin.headingTwoFontSize = definition.getHeadingTwoFontSize();
            mixin.headingThreeFontSize = definition.getHeadingThreeFontSize();
            mixin.headingFourFontSize = definition.getHeadingFourFontSize();
            mixin.headingFiveFontSize = definition.getHeadingFiveFontSize();
            mixin.headingSixFontSize = definition.getHeadingSixFontSize();
            mixin.leftMarginTwips = definition.getLeftMarginTwips();
            mixin.rightMarginTwips = definition.getRightMarginTwips();
            mixin.topMarginTwips = definition.getTopMarginTwips();
            mixin.bottomMarginTwips = definition.getBottomMarginTwips();
            return mixin;
        }

        WordRenderStyleDefinition toDefinition() {
            return new WordRenderStyleDefinition(fontFamily, bodyFontSize, titleColor, headingColor, accentColor,
                coverTitleFontSize, headingOneFontSize, headingTwoFontSize, headingThreeFontSize,
                headingFourFontSize, headingFiveFontSize, headingSixFontSize,
                leftMarginTwips, rightMarginTwips, topMarginTwips, bottomMarginTwips);
        }
    }
}
