package com.wordrender.core;

public class WordRenderException extends RuntimeException {

    public WordRenderException(String message) {
        super(message);
    }

    public WordRenderException(String message, Throwable cause) {
        super(message, cause);
    }
}
