package com.wordrender.cache;

import com.wordrender.style.WordRenderStyleDefinition;

public class WordRenderNoOpCacheService implements WordRenderCacheService {

    @Override
    public WordRenderStyleDefinition getStyle(String key) {
        return null;
    }

    @Override
    public void putStyle(String key, WordRenderStyleDefinition definition) {
        // no-op
    }
}
