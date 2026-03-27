package com.wordrender.support;

import com.wordrender.autoconfigure.WordRenderProperties;
import com.wordrender.core.WordRenderException;
import java.awt.Font;
import java.io.IOException;
import java.io.InputStream;
import org.springframework.core.io.Resource;
import org.springframework.core.io.ResourceLoader;
import org.springframework.util.StringUtils;

public class WordRenderFontResolver {

    private final String defaultFontFamily;
    private final String regularFontFamily;
    private final String boldFontFamily;

    public WordRenderFontResolver(WordRenderProperties properties, ResourceLoader resourceLoader) {
        this.regularFontFamily = loadFontFamily(resourceLoader, properties.getFonts().getRegularPath(), "regular-path");
        this.boldFontFamily = loadFontFamily(resourceLoader, properties.getFonts().getBoldPath(), "bold-path");
        this.defaultFontFamily = resolveDefaultFontFamily(properties);
    }

    public String getDefaultFontFamily() {
        return defaultFontFamily;
    }

    public String getRegularFontFamily() {
        return regularFontFamily;
    }

    public String getBoldFontFamily() {
        return boldFontFamily;
    }

    private String resolveDefaultFontFamily(WordRenderProperties properties) {
        if (StringUtils.hasText(regularFontFamily)) {
            return regularFontFamily;
        }
        if (properties.getFonts() != null && StringUtils.hasText(properties.getFonts().getPrimaryFamily())) {
            return properties.getFonts().getPrimaryFamily().trim();
        }
        if (StringUtils.hasText(properties.getDefaultFontFamily())) {
            return properties.getDefaultFontFamily().trim();
        }
        return "Noto Sans SC";
    }

    private String loadFontFamily(ResourceLoader resourceLoader, String path, String propertyName) {
        if (!StringUtils.hasText(path)) {
            return null;
        }
        Resource resource = resourceLoader.getResource(path);
        if (!resource.exists()) {
            throw new WordRenderException("Configured font resource not found for " + propertyName + ": " + path);
        }
        try (InputStream inputStream = resource.getInputStream()) {
            Font font = Font.createFont(Font.TRUETYPE_FONT, inputStream);
            String family = font.getFamily();
            if (!StringUtils.hasText(family)) {
                throw new WordRenderException("Unable to resolve font family from " + propertyName + ": " + path);
            }
            return family;
        } catch (Exception ex) {
            try (InputStream inputStream = resource.getInputStream()) {
                Font font = Font.createFont(Font.TYPE1_FONT, inputStream);
                String family = font.getFamily();
                if (!StringUtils.hasText(family)) {
                    throw new WordRenderException("Unable to resolve font family from " + propertyName + ": " + path);
                }
                return family;
            } catch (IOException ioEx) {
                throw new WordRenderException("Failed to read font resource for " + propertyName + ": " + path, ioEx);
            } catch (Exception fontEx) {
                throw new WordRenderException("Failed to parse font resource for " + propertyName + ": " + path, fontEx);
            }
        }
    }
}
