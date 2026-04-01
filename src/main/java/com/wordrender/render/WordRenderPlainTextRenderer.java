package com.wordrender.render;

import com.wordrender.style.WordRenderStyleDefinition;
import com.wordrender.support.WordRenderPoiSupport;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

public class WordRenderPlainTextRenderer implements WordRenderContentRenderer {

    @Override
    public void render(WordRenderBodyTarget target, String content, WordRenderStyleDefinition styleDefinition,
                       int baseHeadingLevel) {
        String[] blocks = content.split("\\r?\\n\\s*\\r?\\n");
        for (String block : blocks) {
            XWPFParagraph paragraph = target.createParagraph(ParagraphAlignment.BOTH, 120);
            WordRenderPoiSupport.applyFirstLineIndentChars(paragraph, 2);
            String[] lines = block.split("\\r?\\n");
            for (int i = 0; i < lines.length; i++) {
                WordRenderPoiSupport.appendText(paragraph, lines[i], styleDefinition.getFontFamily(),
                    styleDefinition.getBodyFontSize(), false, false, null);
                if (i < lines.length - 1) {
                    paragraph.createRun().addBreak();
                }
            }
        }
    }
}
