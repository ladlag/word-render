package com.wordrender.autoconfigure;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.wordrender.cache.WordRenderCacheService;
import com.wordrender.cache.WordRenderMemoryCacheService;
import com.wordrender.cache.WordRenderNoOpCacheService;
import com.wordrender.cache.WordRenderRedisCacheService;
import com.wordrender.core.WordRenderDocumentComposer;
import com.wordrender.core.WordRenderService;
import com.wordrender.core.WordRenderServiceImpl;
import com.wordrender.model.WordRenderCacheType;
import com.wordrender.style.WordRenderStyleRegistry;
import com.wordrender.support.WordRenderFontResolver;
import com.wordrender.support.WordRenderWatermarkCustomizer;
import org.springframework.boot.autoconfigure.condition.ConditionalOnClass;
import org.springframework.boot.autoconfigure.condition.ConditionalOnMissingBean;
import org.springframework.boot.autoconfigure.condition.ConditionalOnProperty;
import org.springframework.boot.context.properties.EnableConfigurationProperties;
import org.springframework.beans.factory.ObjectProvider;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.core.io.ResourceLoader;
import org.springframework.data.redis.core.StringRedisTemplate;

@Configuration
@ConditionalOnProperty(prefix = "word-render", name = "enabled", havingValue = "true", matchIfMissing = true)
@EnableConfigurationProperties(WordRenderProperties.class)
public class WordRenderAutoConfiguration {

    @Bean
    @ConditionalOnMissingBean
    public WordRenderFontResolver wordRenderFontResolver(WordRenderProperties properties, ResourceLoader resourceLoader) {
        return new WordRenderFontResolver(properties, resourceLoader);
    }

    @Bean
    @ConditionalOnMissingBean
    public WordRenderStyleRegistry wordRenderStyleRegistry(WordRenderProperties properties,
                                                           WordRenderFontResolver fontResolver) {
        return new WordRenderStyleRegistry(properties, fontResolver);
    }

    @Bean
    @ConditionalOnMissingBean
    public WordRenderWatermarkCustomizer wordRenderWatermarkCustomizer(WordRenderFontResolver fontResolver) {
        return new WordRenderWatermarkCustomizer(fontResolver);
    }

    @Bean
    @ConditionalOnMissingBean
    public WordRenderDocumentComposer wordRenderDocumentComposer(WordRenderProperties properties,
                                                                 WordRenderFontResolver fontResolver,
                                                                 WordRenderWatermarkCustomizer watermarkCustomizer) {
        return new WordRenderDocumentComposer(properties, fontResolver, watermarkCustomizer);
    }

    @Bean
    @ConditionalOnMissingBean
    public WordRenderCacheService wordRenderCacheService(WordRenderProperties properties,
                                                         ObjectProvider<StringRedisTemplate> redisTemplateProvider,
                                                         ObjectProvider<ObjectMapper> objectMapperProvider) {
        if (!properties.getCache().isEnabled()) {
            return new WordRenderNoOpCacheService();
        }
        if (properties.getCache().getType() == WordRenderCacheType.REDIS) {
            StringRedisTemplate redisTemplate = redisTemplateProvider.getIfAvailable();
            ObjectMapper objectMapper = objectMapperProvider.getIfAvailable();
            if (redisTemplate != null && objectMapper != null) {
                return new WordRenderRedisCacheService(redisTemplate, objectMapper, properties.getCache().getRedisKeyPrefix());
            }
            return new WordRenderMemoryCacheService();
        }
        return new WordRenderMemoryCacheService();
    }

    @Bean
    @ConditionalOnMissingBean
    public WordRenderService wordRenderService(WordRenderProperties properties,
                                               WordRenderStyleRegistry styleRegistry,
                                               WordRenderDocumentComposer documentComposer,
                                               WordRenderCacheService cacheService) {
        return new WordRenderServiceImpl(properties, styleRegistry, documentComposer, cacheService);
    }
}
