package com.wordrender.render;

import com.vladsch.flexmark.ast.BlockQuote;
import com.vladsch.flexmark.ast.BulletList;
import com.vladsch.flexmark.ast.BulletListItem;
import com.vladsch.flexmark.ast.Code;
import com.vladsch.flexmark.ast.Emphasis;
import com.vladsch.flexmark.ast.FencedCodeBlock;
import com.vladsch.flexmark.ast.HardLineBreak;
import com.vladsch.flexmark.ast.Heading;
import com.vladsch.flexmark.ast.Link;
import com.vladsch.flexmark.ast.OrderedList;
import com.vladsch.flexmark.ast.OrderedListItem;
import com.vladsch.flexmark.ast.Paragraph;
import com.vladsch.flexmark.ast.SoftLineBreak;
import com.vladsch.flexmark.ast.StrongEmphasis;
import com.vladsch.flexmark.ast.Text;
import com.vladsch.flexmark.ast.ThematicBreak;
import com.vladsch.flexmark.ext.tables.TableBlock;
import com.vladsch.flexmark.ext.tables.TableBody;
import com.vladsch.flexmark.ext.tables.TableCell;
import com.vladsch.flexmark.ext.tables.TableHead;
import com.vladsch.flexmark.ext.tables.TableRow;
import com.vladsch.flexmark.ext.tables.TableSeparator;
import com.vladsch.flexmark.parser.Parser;
import com.vladsch.flexmark.util.ast.Node;
import com.vladsch.flexmark.util.data.MutableDataSet;
import com.wordrender.style.WordRenderStyleDefinition;
import com.wordrender.support.WordRenderPoiSupport;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STShd;

public class WordRenderMarkdownRenderer implements WordRenderContentRenderer {

    private final Parser parser;

    public WordRenderMarkdownRenderer() {
        MutableDataSet options = new MutableDataSet();
        options.set(Parser.EXTENSIONS, Arrays.asList(
            com.vladsch.flexmark.ext.tables.TablesExtension.create()
        ));
        this.parser = Parser.builder(options).build();
    }

    @Override
    public void render(WordRenderBodyTarget target, String content, WordRenderStyleDefinition styleDefinition,
                       int baseHeadingLevel) {
        Node root = parser.parse(content == null ? "" : content);
        RenderContext context = new RenderContext(target, styleDefinition, baseHeadingLevel);
        for (Node node = root.getFirstChild(); node != null; node = node.getNext()) {
            renderBlock(context, node, 0);
        }
    }

    private void renderBlock(RenderContext context, Node node, int listLevel) {
        if (node instanceof Heading) {
            renderHeading(context, (Heading) node);
            return;
        }
        if (node instanceof Paragraph) {
            renderParagraph(context, (Paragraph) node, null);
            return;
        }
        if (node instanceof BulletList) {
            renderList(context, node, false, listLevel);
            return;
        }
        if (node instanceof OrderedList) {
            renderList(context, node, true, listLevel);
            return;
        }
        if (node instanceof BlockQuote) {
            renderBlockQuote(context, (BlockQuote) node);
            return;
        }
        if (node instanceof FencedCodeBlock) {
            renderCodeBlock(context, (FencedCodeBlock) node);
            return;
        }
        if (node instanceof TableBlock) {
            renderTable(context, (TableBlock) node);
            return;
        }
        if (node instanceof ThematicBreak) {
            XWPFParagraph paragraph = context.target.createParagraph(ParagraphAlignment.CENTER, 120);
            XWPFRun run = WordRenderPoiSupport.createRun(paragraph, context.styleDefinition.getFontFamily(),
                context.styleDefinition.getBodyFontSize(), context.styleDefinition.getAccentColor(), false);
            run.setText("--------------------------------------------------");
            return;
        }
        for (Node child = node.getFirstChild(); child != null; child = child.getNext()) {
            renderBlock(context, child, listLevel);
        }
    }

    private void renderHeading(RenderContext context, Heading heading) {
        XWPFParagraph paragraph = context.target.createParagraph(ParagraphAlignment.LEFT, 140);
        int level = Math.min(heading.getLevel() + context.baseHeadingLevel - 1, 6);
        int fontSize = resolveHeadingFontSize(context, level);
        paragraph.setStyle(context.styleDefinition.resolveHeadingStyleId(level));
        appendInlineChildren(paragraph, heading, context.styleDefinition, fontSize, true, false);
    }

    private int resolveHeadingFontSize(RenderContext context, int level) {
        if (level <= 1) {
            return context.styleDefinition.getHeadingOneFontSize();
        }
        if (level == 2) {
            return context.styleDefinition.getHeadingTwoFontSize();
        }
        if (level == 3) {
            return context.styleDefinition.getHeadingThreeFontSize();
        }
        if (level == 4) {
            return context.styleDefinition.getHeadingFourFontSize();
        }
        if (level == 5) {
            return context.styleDefinition.getHeadingFiveFontSize();
        }
        return context.styleDefinition.getHeadingSixFontSize();
    }

    private void renderParagraph(RenderContext context, Paragraph paragraphNode, BigInteger numId) {
        XWPFParagraph paragraph = context.target.createParagraph(ParagraphAlignment.BOTH, 120);
        if (numId != null) {
            paragraph.setNumID(numId);
        } else {
            WordRenderPoiSupport.applyFirstLineIndentChars(paragraph, 2);
        }
        appendInlineChildren(paragraph, paragraphNode, context.styleDefinition,
            context.styleDefinition.getBodyFontSize(), false, false);
    }

    private void renderList(RenderContext context, Node listNode, boolean ordered, int listLevel) {
        BigInteger numId = ordered ? context.orderedNumId : context.bulletNumId;
        for (Node child = listNode.getFirstChild(); child != null; child = child.getNext()) {
            if ((ordered && child instanceof OrderedListItem) || (!ordered && child instanceof BulletListItem)) {
                renderListItem(context, child, numId, listLevel);
            }
        }
    }

    private void renderListItem(RenderContext context, Node item, BigInteger numId, int listLevel) {
        for (Node child = item.getFirstChild(); child != null; child = child.getNext()) {
            if (child instanceof Paragraph) {
                XWPFParagraph paragraph = context.target.createParagraph(ParagraphAlignment.BOTH, 80);
                WordRenderPoiSupport.applyList(paragraph, numId, Math.min(listLevel, 2));
                appendInlineChildren(paragraph, child, context.styleDefinition,
                    context.styleDefinition.getBodyFontSize(), false, false);
            } else if (child instanceof BulletList || child instanceof OrderedList) {
                renderList(context, child, child instanceof OrderedList, listLevel + 1);
            } else {
                renderBlock(context, child, listLevel);
            }
        }
    }

    private void renderBlockQuote(RenderContext context, BlockQuote quote) {
        XWPFParagraph paragraph = context.target.createParagraph(ParagraphAlignment.LEFT, 100);
        paragraph.setIndentationLeft(360);
        paragraph.setBorderLeft(org.apache.poi.xwpf.usermodel.Borders.SINGLE);
        appendInlineChildren(paragraph, quote, context.styleDefinition,
            context.styleDefinition.getBodyFontSize(), false, false);
    }

    private void renderCodeBlock(RenderContext context, FencedCodeBlock codeBlock) {
        XWPFParagraph paragraph = context.target.createParagraph(ParagraphAlignment.LEFT, 120);
        if (!paragraph.getCTP().isSetPPr()) {
            paragraph.getCTP().addNewPPr();
        }
        CTShd shd = paragraph.getCTP().getPPr().isSetShd()
            ? paragraph.getCTP().getPPr().getShd()
            : paragraph.getCTP().getPPr().addNewShd();
        shd.setVal(STShd.CLEAR);
        shd.setColor("auto");
        shd.setFill("F4F4F4");
        XWPFRun run = paragraph.createRun();
        WordRenderPoiSupport.applyFontFamily(run, "Courier New");
        run.setFontSize(context.styleDefinition.getBodyFontSize());
        run.setText(codeBlock.getContentChars().normalizeEOL().toString());
    }

    private void renderTable(RenderContext context, TableBlock tableBlock) {
        List<TableRenderRow> rows = collectTableRows(tableBlock);
        int columns = rows.isEmpty() ? 0 : rows.get(0).cells.size();
        if (columns == 0) {
            return;
        }
        XWPFTable table = context.target.createTable(rows.size(), columns);
        WordRenderPoiSupport.applyTableColumnWidths(table, buildColumnWidths(rows, columns));
        for (int rowIndex = 0; rowIndex < rows.size(); rowIndex++) {
            TableRenderRow renderRow = rows.get(rowIndex);
            for (int cellIndex = 0; cellIndex < renderRow.cells.size(); cellIndex++) {
                XWPFTableCell cell = table.getRow(rowIndex).getCell(cellIndex);
                XWPFParagraph paragraph = cell.getParagraphs().get(0);
                clearParagraph(paragraph);
                paragraph.setAlignment(ParagraphAlignment.LEFT);
                appendInlineChildren(paragraph, renderRow.cells.get(cellIndex), context.styleDefinition,
                    context.styleDefinition.getBodyFontSize(), renderRow.header, false);
                if (renderRow.header) {
                    WordRenderPoiSupport.applyHeaderCellStyle(cell, "EAEAEA");
                }
            }
        }
    }

    private List<TableRenderRow> collectTableRows(TableBlock tableBlock) {
        List<TableRenderRow> rows = new ArrayList<>();
        for (Node section = tableBlock.getFirstChild(); section != null; section = section.getNext()) {
            if (section instanceof TableSeparator) {
                continue;
            }
            for (Node rowNode = section.getFirstChild(); rowNode != null; rowNode = rowNode.getNext()) {
                if (!(rowNode instanceof TableRow)) {
                    continue;
                }
                List<TableCell> cells = new ArrayList<>();
                for (Node cellNode = rowNode.getFirstChild(); cellNode != null; cellNode = cellNode.getNext()) {
                    if (cellNode instanceof TableCell) {
                        cells.add((TableCell) cellNode);
                    }
                }
                if (!cells.isEmpty()) {
                    rows.add(new TableRenderRow(section instanceof TableHead, cells));
                }
            }
        }
        return rows;
    }

    private int[] buildColumnWidths(List<TableRenderRow> rows, int columns) {
        int[] contentWeights = new int[columns];
        for (int index = 0; index < columns; index++) {
            contentWeights[index] = 8;
        }
        for (TableRenderRow row : rows) {
            for (int cellIndex = 0; cellIndex < row.cells.size(); cellIndex++) {
                contentWeights[cellIndex] = Math.max(contentWeights[cellIndex],
                    estimateCellWeight(extractPlainText(row.cells.get(cellIndex))));
            }
        }
        return WordRenderPoiSupport.createAdaptiveColumnWidths(contentWeights);
    }

    private int estimateCellWeight(String text) {
        if (text == null || text.trim().isEmpty()) {
            return 8;
        }
        int visualLength = 0;
        for (char character : text.trim().toCharArray()) {
            visualLength += character <= 127 ? 1 : 2;
        }
        return Math.max(8, Math.min(visualLength + 4, 32));
    }

    private String extractPlainText(Node node) {
        if (node == null) {
            return "";
        }
        if (node instanceof Text) {
            return ((Text) node).getChars().toString();
        }
        if (node instanceof Code) {
            return ((Code) node).getText().toString();
        }
        if (node instanceof SoftLineBreak || node instanceof HardLineBreak) {
            return " ";
        }
        if (node instanceof Link) {
            Link link = (Link) node;
            return link.getText().isEmpty() ? link.getUrl().toString() : link.getText().toString();
        }
        StringBuilder builder = new StringBuilder();
        for (Node child = node.getFirstChild(); child != null; child = child.getNext()) {
            builder.append(extractPlainText(child));
        }
        return builder.toString().replaceAll("\\s+", " ").trim();
    }

    private int countCells(TableRow row) {
        int count = 0;
        for (Node cellNode = row.getFirstChild(); cellNode != null; cellNode = cellNode.getNext()) {
            if (cellNode instanceof TableCell) {
                count++;
            }
        }
        return count;
    }

    private void clearParagraph(XWPFParagraph paragraph) {
        for (int i = paragraph.getRuns().size() - 1; i >= 0; i--) {
            paragraph.removeRun(i);
        }
    }

    private void appendInlineChildren(XWPFParagraph paragraph, Node parent, WordRenderStyleDefinition styleDefinition,
                                      int fontSize, boolean inheritedBold, boolean inheritedItalic) {
        for (Node child = parent.getFirstChild(); child != null; child = child.getNext()) {
            appendInline(paragraph, child, styleDefinition, fontSize, inheritedBold, inheritedItalic);
        }
    }

    private void appendInline(XWPFParagraph paragraph, Node node, WordRenderStyleDefinition styleDefinition,
                              int fontSize, boolean bold, boolean italic) {
        if (node instanceof Text) {
            WordRenderPoiSupport.appendText(paragraph, ((Text) node).getChars().toString(),
                styleDefinition.getFontFamily(), fontSize, bold, italic, null);
            return;
        }
        if (node instanceof SoftLineBreak || node instanceof HardLineBreak) {
            paragraph.createRun().addBreak();
            return;
        }
        if (node instanceof StrongEmphasis) {
            appendInlineChildren(paragraph, node, styleDefinition, fontSize, true, italic);
            return;
        }
        if (node instanceof Emphasis) {
            appendInlineChildren(paragraph, node, styleDefinition, fontSize, bold, true);
            return;
        }
        if (node instanceof Code) {
            WordRenderPoiSupport.appendCode(paragraph, ((Code) node).getText().toString(), "Courier New", fontSize);
            return;
        }
        if (node instanceof Link) {
            Link link = (Link) node;
            String label = link.getText().isEmpty() ? link.getUrl().toString() : link.getText().toString();
            WordRenderPoiSupport.appendHyperlink(paragraph, link.getUrl().toString(), label,
                styleDefinition.getFontFamily(), fontSize);
            return;
        }
        appendInlineChildren(paragraph, node, styleDefinition, fontSize, bold, italic);
    }

    private static class RenderContext {
        private final WordRenderBodyTarget target;
        private final WordRenderStyleDefinition styleDefinition;
        private final BigInteger bulletNumId;
        private final BigInteger orderedNumId;
        private final int baseHeadingLevel;

        private RenderContext(WordRenderBodyTarget target, WordRenderStyleDefinition styleDefinition,
                              int baseHeadingLevel) {
            this.target = target;
            this.styleDefinition = styleDefinition;
            this.baseHeadingLevel = Math.max(1, baseHeadingLevel);
            this.bulletNumId = WordRenderPoiSupport.ensureNumbering(target.getDocument(), false);
            this.orderedNumId = WordRenderPoiSupport.ensureNumbering(target.getDocument(), true);
        }
    }

    private static class TableRenderRow {
        private final boolean header;
        private final List<TableCell> cells;

        private TableRenderRow(boolean header, List<TableCell> cells) {
            this.header = header;
            this.cells = cells;
        }
    }
}
