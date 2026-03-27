package com.wordrender.core;

import com.wordrender.model.WordRenderOptions;
import java.io.OutputStream;
import java.nio.file.Path;

public interface WordRenderService {

    byte[] renderDocx(String content, WordRenderOptions options);

    void renderDocxToStream(String content, WordRenderOptions options, OutputStream outputStream);

    void renderDocxToFile(String content, WordRenderOptions options, Path target);
}
