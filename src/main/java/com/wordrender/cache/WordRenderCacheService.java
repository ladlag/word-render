package com.wordrender.cache;

import com.wordrender.style.WordRenderStyleDefinition;

public interface WordRenderCacheService {

    WordRenderStyleDefinition getStyle(String key);

    void putStyle(String key, WordRenderStyleDefinition definition);
}
