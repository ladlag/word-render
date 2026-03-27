package com.wordrender.cache;

import com.wordrender.style.WordRenderStyleDefinition;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

public class WordRenderMemoryCacheService implements WordRenderCacheService {

    private final Map<String, WordRenderStyleDefinition> styleCache = new ConcurrentHashMap<String, WordRenderStyleDefinition>();

    @Override
    public WordRenderStyleDefinition getStyle(String key) {
        return styleCache.get(key);
    }

    @Override
    public void putStyle(String key, WordRenderStyleDefinition definition) {
        if (definition != null) {
            styleCache.put(key, definition);
        }
    }
}
