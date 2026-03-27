package com.wordrender.render;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;

public interface WordRenderBodyTarget {

    XWPFDocument getDocument();

    XWPFParagraph createParagraph(ParagraphAlignment alignment, int spacingAfter);

    XWPFTable createTable(int rows, int columns);
}
