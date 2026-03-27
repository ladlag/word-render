package com.wordrender.render;

import com.wordrender.style.WordRenderStyleDefinition;
public interface WordRenderContentRenderer {

    void render(WordRenderBodyTarget target, String content, WordRenderStyleDefinition styleDefinition, int baseHeadingLevel);
}
