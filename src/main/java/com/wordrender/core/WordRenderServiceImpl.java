package com.wordrender.core;

import com.wordrender.autoconfigure.WordRenderProperties;
import com.wordrender.cache.WordRenderCacheService;
import com.wordrender.model.WordRenderContentType;
import com.wordrender.model.WordRenderOptions;
import com.wordrender.model.WordRenderTemplateBinding;
import com.wordrender.model.WordRenderTemplateMode;
import com.wordrender.render.WordRenderAppendBodyTarget;
import com.wordrender.render.WordRenderContentRenderer;
import com.wordrender.render.WordRenderMarkdownRenderer;
import com.wordrender.render.WordRenderPlainTextRenderer;
import com.wordrender.style.WordRenderStyleDefinition;
import com.wordrender.style.WordRenderStyleRegistry;
import java.io.IOException;
import java.io.OutputStream;
import java.io.ByteArrayOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.Map;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.util.StringUtils;

public class WordRenderServiceImpl implements WordRenderService {

    private final WordRenderProperties properties;
    private final WordRenderStyleRegistry styleRegistry;
    private final WordRenderDocumentComposer documentComposer;
    private final WordRenderCacheService cacheService;
    private final WordRenderContentRenderer markdownRenderer;
    private final WordRenderContentRenderer plainTextRenderer;
    private final WordRenderTemplatePlaceholderProcessor templatePlaceholderProcessor;

    public WordRenderServiceImpl(WordRenderProperties properties, WordRenderStyleRegistry styleRegistry,
                                 WordRenderDocumentComposer documentComposer, WordRenderCacheService cacheService) {
        this.properties = properties;
        this.styleRegistry = styleRegistry;
        this.documentComposer = documentComposer;
        this.cacheService = cacheService;
        this.markdownRenderer = new WordRenderMarkdownRenderer();
        this.plainTextRenderer = new WordRenderPlainTextRenderer();
        this.templatePlaceholderProcessor = new WordRenderTemplatePlaceholderProcessor(markdownRenderer, plainTextRenderer);
    }

    @Override
    public byte[] renderDocx(String content, WordRenderOptions options) {
        try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            renderDocxToStream(content, options, outputStream);
            return outputStream.toByteArray();
        } catch (IOException ex) {
            throw new WordRenderException("Failed to render docx report", ex);
        }
    }

    @Override
    public void renderDocxToStream(String content, WordRenderOptions options, OutputStream outputStream) {
        if (outputStream == null) {
            throw new WordRenderException("OutputStream must not be null");
        }
        WordRenderOptions resolvedOptions = resolveOptions(options);
        validate(content, resolvedOptions);
        WordRenderStyleDefinition styleDefinition = resolveStyle(resolvedOptions);
        try (XWPFDocument document = documentComposer.createDocument(resolvedOptions, styleDefinition)) {
            if (resolvedOptions.getTemplateMode() == WordRenderTemplateMode.PLACEHOLDER) {
                templatePlaceholderProcessor.replace(document, resolvedOptions.getTemplateBindings(), styleDefinition);
            } else {
                documentComposer.prepareTemplateForDynamicContent(document, resolvedOptions);
                resolveRenderer(resolvedOptions.getContentType()).render(
                    new WordRenderAppendBodyTarget(document), content, styleDefinition, 1);
            }
            document.write(outputStream);
        } catch (IOException ex) {
            throw new WordRenderException("Failed to render docx report", ex);
        }
    }

    @Override
    public void renderDocxToFile(String content, WordRenderOptions options, Path target) {
        if (target == null) {
            throw new WordRenderException("Target file path must not be null");
        }
        Path tempFile = null;
        try {
            if (target.getParent() != null) {
                Files.createDirectories(target.getParent());
            }
            Path parentDir = target.getParent() == null ? target.toAbsolutePath().getParent() : target.getParent();
            String targetName = target.getFileName() == null ? "word-render" : target.getFileName().toString();
            tempFile = Files.createTempFile(parentDir, targetName + ".", ".tmp");
            try (OutputStream outputStream = Files.newOutputStream(tempFile)) {
                renderDocxToStream(content, options, outputStream);
            }
            moveIntoPlace(tempFile, target);
            tempFile = null;
        } catch (IOException ex) {
            throw new WordRenderException("Failed to write docx file to " + target, ex);
        } finally {
            if (tempFile != null) {
                try {
                    Files.deleteIfExists(tempFile);
                } catch (IOException ignored) {
                    // Best-effort cleanup for failed temp-file writes.
                }
            }
        }
    }

    private void validate(String content, WordRenderOptions options) {
        if (options == null) {
            throw new WordRenderException("WordRenderOptions must not be null");
        }
        if (options.getTemplateMode() == WordRenderTemplateMode.PLACEHOLDER) {
            validateColorOptions(options);
            validateTemplateBindings(options.getTemplateBindings());
            if (!StringUtils.hasText(options.getTemplateResource())) {
                throw new WordRenderException("Template resource must be provided for placeholder mode");
            }
            return;
        }
        validateColorOptions(options);
        if (!StringUtils.hasText(content)) {
            throw new WordRenderException("Report content must not be empty");
        }
        if (content.length() > 1_000_000) {
            throw new WordRenderException("Report content is too large");
        }
    }

    private void moveIntoPlace(Path tempFile, Path target) throws IOException {
        try {
            Files.move(tempFile, target, StandardCopyOption.REPLACE_EXISTING, StandardCopyOption.ATOMIC_MOVE);
        } catch (IOException ex) {
            Files.move(tempFile, target, StandardCopyOption.REPLACE_EXISTING);
        }
    }

    private WordRenderOptions resolveOptions(WordRenderOptions options) {
        WordRenderOptions resolved = new WordRenderOptions();
        resolved.setContentType(options.getContentType() == null ? WordRenderContentType.MARKDOWN : options.getContentType());
        resolved.setReportStyle(options.getReportStyle() == null ? properties.getDefaultStyle() : options.getReportStyle());
        resolved.setTitle(options.getTitle());
        resolved.setSubTitle(options.getSubTitle());
        resolved.setAuthor(options.getAuthor());
        resolved.setOutputName(options.getOutputName());
        resolved.setTemplateResource(options.getTemplateResource());
        resolved.setTemplateMode(options.getTemplateMode() == null ? WordRenderTemplateMode.APPEND : options.getTemplateMode());
        resolved.setAppendPageBreakAfterTemplate(options.isAppendPageBreakAfterTemplate());
        resolved.setHeaderText(options.getHeaderText());
        resolved.setFooterText(options.getFooterText());
        resolved.setShowPageNumber(options.isShowPageNumber());
        resolved.setHeaderPosition(options.getHeaderPosition());
        resolved.setTitleColor(options.getTitleColor());
        resolved.setHeadingColor(options.getHeadingColor());
        resolved.setAccentColor(options.getAccentColor());
        resolved.setWatermarkText(options.getWatermarkText());
        resolved.setWatermarkColor(options.getWatermarkColor());
        resolved.setWatermarkFontSize(options.getWatermarkFontSize());
        resolved.setTemplateBindings(options.getTemplateBindings());
        resolved.setMetadata(options.getMetadata());
        if (!StringUtils.hasText(resolved.getWatermarkText())
            && properties.getWatermark() != null
            && properties.getWatermark().isEnabled()) {
            resolved.setWatermarkText(properties.getWatermark().getText());
        }
        if (!StringUtils.hasText(resolved.getWatermarkColor()) && properties.getWatermark() != null) {
            resolved.setWatermarkColor(properties.getWatermark().getColor());
        }
        if (resolved.getWatermarkFontSize() == null && properties.getWatermark() != null) {
            resolved.setWatermarkFontSize(properties.getWatermark().getFontSize());
        }
        return resolved;
    }

    private WordRenderStyleDefinition resolveStyle(WordRenderOptions options) {
        String cacheKey = new StringBuilder("style:")
            .append(options.getReportStyle().name())
            .append(':').append(nullSafe(options.getTitleColor()))
            .append(':').append(nullSafe(options.getHeadingColor()))
            .append(':').append(nullSafe(options.getAccentColor()))
            .toString();
        WordRenderStyleDefinition cached = cacheService.getStyle(cacheKey);
        if (cached != null) {
            return cached;
        }
        WordRenderStyleDefinition definition = styleRegistry.applyOverrides(
            styleRegistry.getStyle(options.getReportStyle()),
            options.getTitleColor(),
            options.getHeadingColor(),
            options.getAccentColor());
        cacheService.putStyle(cacheKey, definition);
        return definition;
    }

    private String nullSafe(String value) {
        return value == null ? "" : value.trim().replace("#", "").toUpperCase();
    }

    private WordRenderContentRenderer resolveRenderer(WordRenderContentType contentType) {
        return contentType == WordRenderContentType.PLAIN_TEXT ? plainTextRenderer : markdownRenderer;
    }

    private void validateTemplateBindings(Map<String, WordRenderTemplateBinding> bindings) {
        if (bindings == null || bindings.isEmpty()) {
            throw new WordRenderException("Template bindings must not be empty in placeholder mode");
        }
        for (Map.Entry<String, WordRenderTemplateBinding> entry : bindings.entrySet()) {
            if (!StringUtils.hasText(entry.getKey())) {
                throw new WordRenderException("Template binding key must not be empty");
            }
            if (entry.getValue() == null) {
                throw new WordRenderException("Template binding must not be null for key " + entry.getKey());
            }
            String content = entry.getValue().getContent();
            if (StringUtils.hasText(content) && content.length() > 1_000_000) {
                throw new WordRenderException("Template binding content is too large for key " + entry.getKey());
            }
        }
    }

    private void validateColorOptions(WordRenderOptions options) {
        validateHexColor(options.getTitleColor(), "titleColor");
        validateHexColor(options.getHeadingColor(), "headingColor");
        validateHexColor(options.getAccentColor(), "accentColor");
        validateHexColor(options.getWatermarkColor(), "watermarkColor");
        validateWatermarkFontSize(options.getWatermarkFontSize());
    }

    private void validateHexColor(String color, String fieldName) {
        if (!StringUtils.hasText(color)) {
            return;
        }
        String normalized = color.trim().replace("#", "").toUpperCase();
        if (!normalized.matches("[0-9A-F]{6}")) {
            throw new WordRenderException(fieldName + " must be a 6-digit hex color");
        }
    }

    private void validateWatermarkFontSize(Integer fontSize) {
        if (fontSize == null) {
            return;
        }
        if (fontSize.intValue() < 8 || fontSize.intValue() > 72) {
            throw new WordRenderException("watermarkFontSize must be between 8 and 72");
        }
    }
}
