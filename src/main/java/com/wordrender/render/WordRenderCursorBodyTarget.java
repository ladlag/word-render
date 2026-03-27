package com.wordrender.render;

import com.wordrender.support.WordRenderPoiSupport;
import org.apache.poi.xwpf.usermodel.IBody;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.xmlbeans.XmlCursor;

public class WordRenderCursorBodyTarget implements WordRenderBodyTarget {

    private final IBody body;
    private XmlCursor cursor;

    public WordRenderCursorBodyTarget(IBody body, XmlCursor cursor) {
        this.body = body;
        this.cursor = cursor;
    }

    @Override
    public XWPFDocument getDocument() {
        return body.getXWPFDocument();
    }

    @Override
    public XWPFParagraph createParagraph(ParagraphAlignment alignment, int spacingAfter) {
        XWPFParagraph paragraph = body.insertNewParagraph(cursor);
        paragraph.setAlignment(alignment);
        paragraph.setSpacingAfter(spacingAfter);
        advanceCursor(paragraph.getCTP().newCursor());
        return paragraph;
    }

    @Override
    public XWPFTable createTable(int rows, int columns) {
        XWPFTable table = body.insertNewTbl(cursor);
        WordRenderPoiSupport.initializeTable(table, rows, columns);
        advanceCursor(table.getCTTbl().newCursor());
        return table;
    }

    private void advanceCursor(XmlCursor nextCursor) {
        if (cursor != null) {
            cursor.dispose();
        }
        nextCursor.toEndToken();
        nextCursor.toNextToken();
        cursor = nextCursor;
    }
}
