package com.wordrender.model;

public class WordRenderTemplateBinding {

    private String content;
    private WordRenderContentType contentType = WordRenderContentType.MARKDOWN;
    private int baseHeadingLevel = 1;

    public static Builder builder() {
        return new Builder();
    }

    public String getContent() {
        return content;
    }

    public void setContent(String content) {
        this.content = content;
    }

    public WordRenderContentType getContentType() {
        return contentType;
    }

    public void setContentType(WordRenderContentType contentType) {
        this.contentType = contentType;
    }

    public int getBaseHeadingLevel() {
        return baseHeadingLevel;
    }

    public void setBaseHeadingLevel(int baseHeadingLevel) {
        this.baseHeadingLevel = baseHeadingLevel;
    }

    public static final class Builder {
        private final WordRenderTemplateBinding binding = new WordRenderTemplateBinding();

        public Builder content(String content) {
            binding.setContent(content);
            return this;
        }

        public Builder contentType(WordRenderContentType contentType) {
            binding.setContentType(contentType);
            return this;
        }

        public Builder baseHeadingLevel(int baseHeadingLevel) {
            binding.setBaseHeadingLevel(baseHeadingLevel);
            return this;
        }

        public WordRenderTemplateBinding build() {
            return binding;
        }
    }
}
