package com.wordrender.core;

import com.wordrender.model.WordRenderContentType;
import com.wordrender.model.WordRenderTemplateBinding;
import com.wordrender.render.WordRenderBodyTarget;
import com.wordrender.render.WordRenderContentRenderer;
import com.wordrender.render.WordRenderCursorBodyTarget;
import com.wordrender.style.WordRenderStyleDefinition;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.poi.xwpf.usermodel.IBody;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.springframework.util.StringUtils;

public class WordRenderTemplatePlaceholderProcessor {

    private static final Pattern PLACEHOLDER_PATTERN = Pattern.compile("^\\$\\{([A-Za-z0-9_.-]+)\\}$");

    private final WordRenderContentRenderer markdownRenderer;
    private final WordRenderContentRenderer plainTextRenderer;

    public WordRenderTemplatePlaceholderProcessor(WordRenderContentRenderer markdownRenderer,
                                                  WordRenderContentRenderer plainTextRenderer) {
        this.markdownRenderer = markdownRenderer;
        this.plainTextRenderer = plainTextRenderer;
    }

    public void replace(XWPFDocument document, Map<String, WordRenderTemplateBinding> bindings,
                        WordRenderStyleDefinition styleDefinition) {
        Set<String> matched = new LinkedHashSet<String>();
        processBody(document, bindings, matched, styleDefinition);
        for (XWPFHeader header : document.getHeaderList()) {
            processBody(header, bindings, matched, styleDefinition);
        }
        for (XWPFFooter footer : document.getFooterList()) {
            processBody(footer, bindings, matched, styleDefinition);
        }
        List<String> missing = new ArrayList<String>();
        for (String bindingKey : bindings.keySet()) {
            if (!matched.contains(bindingKey)) {
                missing.add(bindingKey);
            }
        }
        if (!missing.isEmpty()) {
            throw new WordRenderException("Template placeholder(s) not found: " + missing);
        }
    }

    private void processBody(IBody body, Map<String, WordRenderTemplateBinding> bindings, Set<String> matched,
                             WordRenderStyleDefinition styleDefinition) {
        List<IBodyElement> elements = new ArrayList<IBodyElement>(body.getBodyElements());
        for (IBodyElement element : elements) {
            if (element instanceof XWPFParagraph) {
                processParagraph(body, (XWPFParagraph) element, bindings, matched, styleDefinition);
                continue;
            }
            if (element instanceof XWPFTable) {
                processTable((XWPFTable) element, bindings, matched, styleDefinition);
            }
        }
    }

    private void processTable(XWPFTable table, Map<String, WordRenderTemplateBinding> bindings, Set<String> matched,
                              WordRenderStyleDefinition styleDefinition) {
        for (XWPFTableRow row : table.getRows()) {
            for (XWPFTableCell cell : row.getTableCells()) {
                processBody(cell, bindings, matched, styleDefinition);
            }
        }
    }

    private void processParagraph(IBody body, XWPFParagraph paragraph, Map<String, WordRenderTemplateBinding> bindings,
                                  Set<String> matched, WordRenderStyleDefinition styleDefinition) {
        String placeholderKey = extractPlaceholderKey(paragraph.getText());
        if (placeholderKey == null || !bindings.containsKey(placeholderKey)) {
            return;
        }
        WordRenderTemplateBinding binding = bindings.get(placeholderKey);
        matched.add(placeholderKey);
        try (XmlCursor cursor = paragraph.getCTP().newCursor()) {
            WordRenderBodyTarget target = new WordRenderCursorBodyTarget(body, cursor);
            if (StringUtils.hasText(binding.getContent())) {
                resolveRenderer(binding.getContentType()).render(target, binding.getContent(), styleDefinition,
                    binding.getBaseHeadingLevel());
            }
        }
        removeParagraph(body, paragraph);
    }

    private String extractPlaceholderKey(String paragraphText) {
        if (!StringUtils.hasText(paragraphText)) {
            return null;
        }
        Matcher matcher = PLACEHOLDER_PATTERN.matcher(paragraphText.trim());
        return matcher.matches() ? matcher.group(1) : null;
    }

    private WordRenderContentRenderer resolveRenderer(WordRenderContentType contentType) {
        return contentType == WordRenderContentType.PLAIN_TEXT ? plainTextRenderer : markdownRenderer;
    }

    private void removeParagraph(IBody body, XWPFParagraph paragraph) {
        if (body instanceof XWPFDocument) {
            int index = ((XWPFDocument) body).getPosOfParagraph(paragraph);
            if (index >= 0) {
                ((XWPFDocument) body).removeBodyElement(index);
            }
            return;
        }
        if (body instanceof XWPFHeader) {
            ((XWPFHeader) body).removeParagraph(paragraph);
            return;
        }
        if (body instanceof XWPFFooter) {
            ((XWPFFooter) body).removeParagraph(paragraph);
            return;
        }
        if (body instanceof XWPFTableCell) {
            int index = ((XWPFTableCell) body).getParagraphs().indexOf(paragraph);
            if (index >= 0) {
                ((XWPFTableCell) body).removeParagraph(index);
            }
            if (((XWPFTableCell) body).getParagraphs().isEmpty()) {
                ((XWPFTableCell) body).addParagraph();
            }
            return;
        }
        throw new WordRenderException("Unsupported placeholder body type: " + body.getClass().getName());
    }
}
