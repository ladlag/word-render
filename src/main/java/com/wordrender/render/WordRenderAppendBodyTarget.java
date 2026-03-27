package com.wordrender.render;

import com.wordrender.core.WordRenderException;
import com.wordrender.support.WordRenderPoiSupport;
import org.apache.poi.xwpf.usermodel.IBody;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHeaderFooter;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.xmlbeans.XmlCursor;

public class WordRenderAppendBodyTarget implements WordRenderBodyTarget {

    private final IBody body;

    public WordRenderAppendBodyTarget(IBody body) {
        this.body = body;
    }

    @Override
    public XWPFDocument getDocument() {
        return body.getXWPFDocument();
    }

    @Override
    public XWPFParagraph createParagraph(ParagraphAlignment alignment, int spacingAfter) {
        XWPFParagraph paragraph;
        if (body instanceof XWPFDocument) {
            paragraph = ((XWPFDocument) body).createParagraph();
        } else if (body instanceof XWPFHeaderFooter) {
            paragraph = ((XWPFHeaderFooter) body).createParagraph();
        } else if (body instanceof XWPFTableCell) {
            paragraph = ((XWPFTableCell) body).addParagraph();
        } else {
            throw new WordRenderException("Unsupported body type: " + body.getClass().getName());
        }
        paragraph.setAlignment(alignment);
        paragraph.setSpacingAfter(spacingAfter);
        return paragraph;
    }

    @Override
    public XWPFTable createTable(int rows, int columns) {
        XWPFTable table;
        if (body instanceof XWPFDocument) {
            table = ((XWPFDocument) body).createTable(rows, columns);
        } else if (body instanceof XWPFHeaderFooter) {
            table = ((XWPFHeaderFooter) body).createTable(rows, columns);
        } else if (body instanceof XWPFTableCell) {
            try (XmlCursor cursor = ((XWPFTableCell) body).getCTTc().newCursor()) {
                cursor.toEndToken();
                table = body.insertNewTbl(cursor);
            }
        } else {
            throw new WordRenderException("Unsupported body type: " + body.getClass().getName());
        }
        WordRenderPoiSupport.initializeTable(table, rows, columns);
        return table;
    }
}
