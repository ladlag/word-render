package com.wordrender.support;

import java.math.BigInteger;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.TableRowAlign;
import org.apache.poi.xwpf.usermodel.XWPFRelation;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHyperlinkRun;
import org.apache.poi.xwpf.usermodel.XWPFNumbering;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTAbstractNum;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTLvl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblGrid;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblGridCol;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblLayoutType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STNumberFormat;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblLayoutType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;

public final class WordRenderPoiSupport {

    private WordRenderPoiSupport() {
    }

    public static XWPFParagraph createParagraph(XWPFDocument document, ParagraphAlignment alignment, int spacingAfter) {
        XWPFParagraph paragraph = document.createParagraph();
        paragraph.setAlignment(alignment);
        paragraph.setSpacingAfter(spacingAfter);
        return paragraph;
    }

    public static XWPFRun createRun(XWPFParagraph paragraph, String fontFamily, int fontSize, String color, boolean bold) {
        XWPFRun run = paragraph.createRun();
        applyFontFamily(run, fontFamily);
        run.setFontSize(fontSize);
        if (color != null) {
            run.setColor(color);
        }
        run.setBold(bold);
        return run;
    }

    public static void appendText(XWPFParagraph paragraph, String text, String fontFamily, int fontSize,
                                  boolean bold, boolean italic, String color) {
        XWPFRun run = paragraph.createRun();
        applyFontFamily(run, fontFamily);
        run.setFontSize(fontSize);
        run.setBold(bold);
        run.setItalic(italic);
        if (color != null) {
            run.setColor(color);
        }
        run.setText(text);
    }

    public static void appendCode(XWPFParagraph paragraph, String text, String fontFamily, int fontSize) {
        XWPFRun run = paragraph.createRun();
        applyFontFamily(run, fontFamily);
        run.setFontSize(fontSize);
        run.setText(text);
        run.setColor("C00000");
    }

    public static void appendHyperlink(XWPFParagraph paragraph, String url, String label, String fontFamily, int fontSize) {
        String id = paragraph.getDocument().getPackagePart()
            .addExternalRelationship(url, XWPFRelation.HYPERLINK.getRelation()).getId();
        CTP ctp = paragraph.getCTP();
        org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHyperlink hyperlink = ctp.addNewHyperlink();
        hyperlink.setId(id);
        CTR ctr = hyperlink.addNewR();
        CTText ctText = ctr.addNewT();
        ctText.setStringValue(label);
        XWPFHyperlinkRun run = new XWPFHyperlinkRun(hyperlink, ctr, paragraph);
        run.setUnderline(UnderlinePatterns.SINGLE);
        run.setColor("0563C1");
        applyFontFamily(run, fontFamily);
        run.setFontSize(fontSize);
    }

    public static void applyFontFamily(XWPFRun run, String fontFamily) {
        if (run == null || fontFamily == null || fontFamily.trim().isEmpty()) {
            return;
        }
        String normalized = fontFamily.trim();
        run.setFontFamily(normalized);
        CTRPr runProperties = run.getCTR().isSetRPr() ? run.getCTR().getRPr() : run.getCTR().addNewRPr();
        CTFonts fonts = runProperties.sizeOfRFontsArray() > 0 ? runProperties.getRFontsArray(0) : runProperties.addNewRFonts();
        fonts.setAscii(normalized);
        fonts.setHAnsi(normalized);
        fonts.setEastAsia(normalized);
        fonts.setCs(normalized);
    }

    public static void copyParagraphFormatting(XWPFParagraph source, XWPFParagraph target) {
        if (source == null || target == null) {
            return;
        }
        if (source.getCTP().isSetPPr()) {
            target.getCTP().setPPr((CTPPr) source.getCTP().getPPr().copy());
        }
        XWPFRun sourceRun = source.getRuns().isEmpty() ? null : source.getRuns().get(0);
        if (sourceRun != null) {
            XWPFRun targetRun = target.createRun();
            if (sourceRun.getCTR().isSetRPr()) {
                targetRun.getCTR().setRPr((CTRPr) sourceRun.getCTR().getRPr().copy());
            }
            target.removeRun(0);
        }
    }

    public static String resolveParagraphFontFamily(XWPFParagraph paragraph) {
        if (paragraph == null || paragraph.getRuns().isEmpty()) {
            return null;
        }
        for (XWPFRun run : paragraph.getRuns()) {
            if (run.getCTR().isSetRPr()) {
                CTRPr runProperties = run.getCTR().getRPr();
                if (runProperties.sizeOfRFontsArray() > 0) {
                    CTFonts fonts = runProperties.getRFontsArray(0);
                    if (fonts.isSetEastAsia() && fonts.getEastAsia() != null && !fonts.getEastAsia().trim().isEmpty()) {
                        return fonts.getEastAsia().trim();
                    }
                    if (fonts.isSetAscii() && fonts.getAscii() != null && !fonts.getAscii().trim().isEmpty()) {
                        return fonts.getAscii().trim();
                    }
                    if (fonts.isSetHAnsi() && fonts.getHAnsi() != null && !fonts.getHAnsi().trim().isEmpty()) {
                        return fonts.getHAnsi().trim();
                    }
                }
            }
            String runFont = run.getFontFamily();
            if (runFont != null && !runFont.trim().isEmpty()) {
                return runFont.trim();
            }
        }
        return null;
    }

    public static BigInteger ensureNumbering(XWPFDocument document, boolean ordered) {
        XWPFNumbering numbering = document.createNumbering();
        CTAbstractNum abstractNum = CTAbstractNum.Factory.newInstance();
        abstractNum.setAbstractNumId(BigInteger.valueOf(ordered ? 1 : 2));
        for (int level = 0; level < 3; level++) {
            CTLvl lvl = abstractNum.addNewLvl();
            lvl.setIlvl(BigInteger.valueOf(level));
            lvl.addNewNumFmt().setVal(ordered ? STNumberFormat.DECIMAL : STNumberFormat.BULLET);
            lvl.addNewLvlText().setVal(ordered ? "%" + (level + 1) + "." : "\u2022");
            lvl.addNewStart().setVal(BigInteger.ONE);
            lvl.addNewLvlJc().setVal(STJc.LEFT);
            lvl.addNewPPr().addNewInd().setLeft(BigInteger.valueOf(720L * (level + 1)));
        }
        BigInteger abstractNumId = numbering.addAbstractNum(new org.apache.poi.xwpf.usermodel.XWPFAbstractNum(abstractNum));
        return numbering.addNum(abstractNumId);
    }

    public static void applyList(XWPFParagraph paragraph, BigInteger numId, int level) {
        paragraph.setNumID(numId);
        if (!paragraph.getCTP().isSetPPr()) {
            paragraph.getCTP().addNewPPr();
        }
        if (!paragraph.getCTP().getPPr().isSetNumPr()) {
            paragraph.getCTP().getPPr().addNewNumPr();
        }
        paragraph.getCTP().getPPr().getNumPr().addNewIlvl().setVal(BigInteger.valueOf(level));
    }

    public static void initializeTable(XWPFTable table, int rows, int columns) {
        if (rows <= 0 || columns <= 0) {
            return;
        }
        int[] defaultWidths = createEvenColumnWidths(columns);
        configureTableLayout(table, defaultWidths);
        if (table.getNumberOfRows() == 0) {
            table.createRow();
        }
        ensureColumns(table.getRow(0), columns);
        while (table.getNumberOfRows() < rows) {
            XWPFTableRow row = table.createRow();
            ensureColumns(row, columns);
        }
        for (int rowIndex = 0; rowIndex < table.getNumberOfRows(); rowIndex++) {
            ensureColumns(table.getRow(rowIndex), columns);
            for (int colIndex = 0; colIndex < columns; colIndex++) {
                XWPFTableCell cell = table.getRow(rowIndex).getCell(colIndex);
                setCellWidth(cell, defaultWidths[colIndex]);
                while (cell.getParagraphs().size() > 1) {
                    cell.removeParagraph(cell.getParagraphs().size() - 1);
                }
                if (!cell.getParagraphs().isEmpty()) {
                    XWPFParagraph paragraph = cell.getParagraphs().get(0);
                    for (int runIndex = paragraph.getRuns().size() - 1; runIndex >= 0; runIndex--) {
                        paragraph.removeRun(runIndex);
                    }
                }
            }
        }
    }

    public static void applyTableColumnWidths(XWPFTable table, int[] columnWidths) {
        if (table == null || columnWidths == null || columnWidths.length == 0) {
            return;
        }
        configureTableLayout(table, columnWidths);
        for (XWPFTableRow row : table.getRows()) {
            ensureColumns(row, columnWidths.length);
            for (int index = 0; index < columnWidths.length; index++) {
                setCellWidth(row.getCell(index), columnWidths[index]);
            }
        }
    }

    public static void applyHeaderCellStyle(XWPFTableCell cell, String fillColor) {
        CTTcPr cellProperties = cell.getCTTc().isSetTcPr() ? cell.getCTTc().getTcPr() : cell.getCTTc().addNewTcPr();
        CTShd shading = cellProperties.isSetShd() ? cellProperties.getShd() : cellProperties.addNewShd();
        shading.setVal(STShd.CLEAR);
        shading.setColor("auto");
        shading.setFill(fillColor);
    }

    private static void ensureColumns(XWPFTableRow row, int columns) {
        while (row.getTableCells().size() < columns) {
            row.addNewTableCell();
        }
    }

    private static void configureTableLayout(XWPFTable table, int[] columnWidths) {
        table.setWidth("100%");
        table.setTableAlignment(TableRowAlign.CENTER);
        table.setCellMargins(80, 100, 80, 100);

        CTTblPr tableProperties = table.getCTTbl().getTblPr() == null
            ? table.getCTTbl().addNewTblPr()
            : table.getCTTbl().getTblPr();
        CTTblLayoutType layoutType = tableProperties.isSetTblLayout()
            ? tableProperties.getTblLayout()
            : tableProperties.addNewTblLayout();
        layoutType.setType(STTblLayoutType.FIXED);

        CTTblWidth width = tableProperties.isSetTblW() ? tableProperties.getTblW() : tableProperties.addNewTblW();
        width.setType(STTblWidth.PCT);
        width.setW(BigInteger.valueOf(5000));

        CTTblGrid grid = table.getCTTbl().getTblGrid() == null
            ? table.getCTTbl().addNewTblGrid()
            : table.getCTTbl().getTblGrid();
        while (grid.sizeOfGridColArray() > 0) {
            grid.removeGridCol(0);
        }
        for (int columnWidth : columnWidths) {
            CTTblGridCol gridCol = grid.addNewGridCol();
            gridCol.setW(BigInteger.valueOf(columnWidth));
        }
    }

    public static int[] createAdaptiveColumnWidths(int[] contentWeights) {
        if (contentWeights == null || contentWeights.length == 0) {
            return new int[0];
        }
        int totalWidth = 9000;
        int minColumnWidth = 1200;
        int[] normalizedWeights = new int[contentWeights.length];
        int totalWeight = 0;
        for (int index = 0; index < contentWeights.length; index++) {
            normalizedWeights[index] = Math.max(1, contentWeights[index]);
            totalWeight += normalizedWeights[index];
        }
        int[] widths = new int[contentWeights.length];
        int allocated = 0;
        for (int index = 0; index < contentWeights.length; index++) {
            int remainingColumns = contentWeights.length - index;
            int remainingWidth = totalWidth - allocated;
            int proportionalWidth = index == contentWeights.length - 1
                ? remainingWidth
                : Math.max(minColumnWidth, (int) Math.round((double) totalWidth * normalizedWeights[index] / totalWeight));
            int maxWidthForColumn = remainingWidth - ((remainingColumns - 1) * minColumnWidth);
            widths[index] = Math.min(Math.max(minColumnWidth, proportionalWidth), maxWidthForColumn);
            allocated += widths[index];
        }
        widths[widths.length - 1] += totalWidth - allocated;
        return widths;
    }

    private static int[] createEvenColumnWidths(int columns) {
        if (columns <= 0) {
            return new int[0];
        }
        int bodyWidth = 9000;
        int[] widths = new int[columns];
        int defaultWidth = Math.max(1200, bodyWidth / columns);
        int allocated = 0;
        for (int index = 0; index < columns; index++) {
            widths[index] = index == columns - 1 ? bodyWidth - allocated : defaultWidth;
            allocated += widths[index];
        }
        return widths;
    }

    private static void setCellWidth(XWPFTableCell cell, int widthTwips) {
        CTTcPr cellProperties = cell.getCTTc().isSetTcPr() ? cell.getCTTc().getTcPr() : cell.getCTTc().addNewTcPr();
        CTTblWidth width = cellProperties.isSetTcW() ? cellProperties.getTcW() : cellProperties.addNewTcW();
        width.setType(STTblWidth.DXA);
        width.setW(BigInteger.valueOf(widthTwips));
    }
}
