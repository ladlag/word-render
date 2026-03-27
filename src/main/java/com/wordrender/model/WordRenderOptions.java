package com.wordrender.model;

import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.Map;

public class WordRenderOptions {

    private WordRenderContentType contentType = WordRenderContentType.MARKDOWN;
    private WordRenderReportStyle reportStyle = WordRenderReportStyle.DEFAULT;
    private String title;
    private String subTitle;
    private String author;
    private String outputName;
    private String templateResource;
    private WordRenderTemplateMode templateMode = WordRenderTemplateMode.APPEND;
    private boolean appendPageBreakAfterTemplate = true;
    private String headerText;
    private String footerText;
    private boolean showPageNumber;
    private WordRenderPosition headerPosition = WordRenderPosition.RIGHT;
    private String titleColor;
    private String headingColor;
    private String accentColor;
    private String watermarkText;
    private String watermarkColor;
    private Integer watermarkFontSize;
    private Map<String, WordRenderTemplateBinding> templateBindings =
        new LinkedHashMap<String, WordRenderTemplateBinding>();
    private Map<String, String> metadata = new LinkedHashMap<String, String>();

    public static Builder builder() {
        return new Builder();
    }

    public WordRenderContentType getContentType() {
        return contentType;
    }

    public void setContentType(WordRenderContentType contentType) {
        this.contentType = contentType;
    }

    public WordRenderReportStyle getReportStyle() {
        return reportStyle;
    }

    public void setReportStyle(WordRenderReportStyle reportStyle) {
        this.reportStyle = reportStyle;
    }

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public String getSubTitle() {
        return subTitle;
    }

    public void setSubTitle(String subTitle) {
        this.subTitle = subTitle;
    }

    public String getAuthor() {
        return author;
    }

    public void setAuthor(String author) {
        this.author = author;
    }

    public String getOutputName() {
        return outputName;
    }

    public void setOutputName(String outputName) {
        this.outputName = outputName;
    }

    public String getTemplateResource() {
        return templateResource;
    }

    public void setTemplateResource(String templateResource) {
        this.templateResource = templateResource;
    }

    public boolean isAppendPageBreakAfterTemplate() {
        return appendPageBreakAfterTemplate;
    }

    public void setAppendPageBreakAfterTemplate(boolean appendPageBreakAfterTemplate) {
        this.appendPageBreakAfterTemplate = appendPageBreakAfterTemplate;
    }

    public WordRenderTemplateMode getTemplateMode() {
        return templateMode;
    }

    public void setTemplateMode(WordRenderTemplateMode templateMode) {
        this.templateMode = templateMode;
    }

    public String getHeaderText() {
        return headerText;
    }

    public void setHeaderText(String headerText) {
        this.headerText = headerText;
    }

    public String getFooterText() {
        return footerText;
    }

    public void setFooterText(String footerText) {
        this.footerText = footerText;
    }

    public boolean isShowPageNumber() {
        return showPageNumber;
    }

    public void setShowPageNumber(boolean showPageNumber) {
        this.showPageNumber = showPageNumber;
    }

    public WordRenderPosition getHeaderPosition() {
        return headerPosition;
    }

    public void setHeaderPosition(WordRenderPosition headerPosition) {
        this.headerPosition = headerPosition;
    }

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

    public String getWatermarkText() {
        return watermarkText;
    }

    public void setWatermarkText(String watermarkText) {
        this.watermarkText = watermarkText;
    }

    public String getWatermarkColor() {
        return watermarkColor;
    }

    public void setWatermarkColor(String watermarkColor) {
        this.watermarkColor = watermarkColor;
    }

    public Integer getWatermarkFontSize() {
        return watermarkFontSize;
    }

    public void setWatermarkFontSize(Integer watermarkFontSize) {
        this.watermarkFontSize = watermarkFontSize;
    }

    public Map<String, WordRenderTemplateBinding> getTemplateBindings() {
        return Collections.unmodifiableMap(templateBindings);
    }

    public void setTemplateBindings(Map<String, WordRenderTemplateBinding> templateBindings) {
        this.templateBindings.clear();
        if (templateBindings != null) {
            this.templateBindings.putAll(templateBindings);
        }
    }

    public void putTemplateBinding(String key, WordRenderTemplateBinding binding) {
        if (key != null && binding != null) {
            this.templateBindings.put(key, binding);
        }
    }

    public Map<String, String> getMetadata() {
        return Collections.unmodifiableMap(metadata);
    }

    public void setMetadata(Map<String, String> metadata) {
        this.metadata.clear();
        if (metadata != null) {
            this.metadata.putAll(metadata);
        }
    }

    public void putMetadata(String key, String value) {
        this.metadata.put(key, value);
    }

    public static final class Builder {
        private final WordRenderOptions options = new WordRenderOptions();

        public Builder contentType(WordRenderContentType contentType) {
            options.setContentType(contentType);
            return this;
        }

        public Builder reportStyle(WordRenderReportStyle reportStyle) {
            options.setReportStyle(reportStyle);
            return this;
        }

        public Builder title(String title) {
            options.setTitle(title);
            return this;
        }

        public Builder subTitle(String subTitle) {
            options.setSubTitle(subTitle);
            return this;
        }

        public Builder author(String author) {
            options.setAuthor(author);
            return this;
        }

        public Builder outputName(String outputName) {
            options.setOutputName(outputName);
            return this;
        }

        public Builder templateResource(String templateResource) {
            options.setTemplateResource(templateResource);
            return this;
        }

        public Builder templateMode(WordRenderTemplateMode templateMode) {
            options.setTemplateMode(templateMode);
            return this;
        }

        public Builder appendPageBreakAfterTemplate(boolean appendPageBreakAfterTemplate) {
            options.setAppendPageBreakAfterTemplate(appendPageBreakAfterTemplate);
            return this;
        }

        public Builder headerText(String headerText) {
            options.setHeaderText(headerText);
            return this;
        }

        public Builder footerText(String footerText) {
            options.setFooterText(footerText);
            return this;
        }

        public Builder showPageNumber(boolean showPageNumber) {
            options.setShowPageNumber(showPageNumber);
            return this;
        }

        public Builder headerPosition(WordRenderPosition headerPosition) {
            options.setHeaderPosition(headerPosition);
            return this;
        }

        public Builder templateBindings(Map<String, WordRenderTemplateBinding> templateBindings) {
            options.setTemplateBindings(templateBindings);
            return this;
        }

        public Builder titleColor(String titleColor) {
            options.setTitleColor(titleColor);
            return this;
        }

        public Builder headingColor(String headingColor) {
            options.setHeadingColor(headingColor);
            return this;
        }

        public Builder accentColor(String accentColor) {
            options.setAccentColor(accentColor);
            return this;
        }

        public Builder watermarkText(String watermarkText) {
            options.setWatermarkText(watermarkText);
            return this;
        }

        public Builder watermarkColor(String watermarkColor) {
            options.setWatermarkColor(watermarkColor);
            return this;
        }

        public Builder watermarkFontSize(Integer watermarkFontSize) {
            options.setWatermarkFontSize(watermarkFontSize);
            return this;
        }

        public Builder templateBinding(String key, WordRenderTemplateBinding binding) {
            options.putTemplateBinding(key, binding);
            return this;
        }

        public Builder metadata(Map<String, String> metadata) {
            options.setMetadata(metadata);
            return this;
        }

        public Builder metadata(String key, String value) {
            options.putMetadata(key, value);
            return this;
        }

        public WordRenderOptions build() {
            return options;
        }
    }
}
