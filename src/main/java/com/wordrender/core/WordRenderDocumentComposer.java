package com.wordrender.core;

import com.wordrender.autoconfigure.WordRenderProperties;
import com.wordrender.model.WordRenderOptions;
import com.wordrender.model.WordRenderPageSize;
import com.wordrender.model.WordRenderPosition;
import com.wordrender.style.WordRenderStyleDefinition;
import com.wordrender.support.WordRenderFontResolver;
import com.wordrender.support.WordRenderPoiSupport;
import com.wordrender.support.WordRenderWatermarkCustomizer;
import java.math.BigInteger;
import java.io.IOException;
import java.io.InputStream;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFStyle;
import org.apache.poi.xwpf.usermodel.XWPFStyles;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSimpleField;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTStyle;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTString;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDecimalNumber;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHpsMeasure;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTOnOff;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STHdrFtr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STStyleType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTabJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTabTlc;
import org.springframework.core.io.DefaultResourceLoader;
import org.springframework.core.io.Resource;

public class WordRenderDocumentComposer {

    private final WordRenderProperties properties;
    private final WordRenderFontResolver fontResolver;
    private final WordRenderWatermarkCustomizer watermarkCustomizer;
    private final DefaultResourceLoader resourceLoader = new DefaultResourceLoader();

    public WordRenderDocumentComposer(WordRenderProperties properties,
                                      WordRenderFontResolver fontResolver,
                                      WordRenderWatermarkCustomizer watermarkCustomizer) {
        this.properties = properties;
        this.fontResolver = fontResolver;
        this.watermarkCustomizer = watermarkCustomizer;
    }

    public XWPFDocument createDocument(WordRenderOptions options, WordRenderStyleDefinition styleDefinition) {
        XWPFDocument document = createBaseDocument(options);
        configurePage(document, styleDefinition);
        configureStyles(document, styleDefinition);
        configureMetadata(document, options);
        addWatermark(document, options);
        addHeader(document, options, styleDefinition);
        addFooter(document, options, styleDefinition);
        addCoverIfNeeded(document, options, styleDefinition);
        return document;
    }

    public boolean hasTemplate(WordRenderOptions options) {
        return hasText(options.getTemplateResource());
    }

    public void prepareTemplateForDynamicContent(XWPFDocument document, WordRenderOptions options) {
        if (!hasTemplate(options) || !options.isAppendPageBreakAfterTemplate()) {
            return;
        }
        XWPFParagraph pageBreak = document.createParagraph();
        pageBreak.createRun().addBreak(BreakType.PAGE);
    }

    private XWPFDocument createBaseDocument(WordRenderOptions options) {
        if (!hasTemplate(options)) {
            return new XWPFDocument();
        }
        Resource resource = resourceLoader.getResource(options.getTemplateResource());
        if (!resource.exists()) {
            throw new WordRenderException("Template resource not found: " + options.getTemplateResource());
        }
        try (InputStream inputStream = resource.getInputStream()) {
            return new XWPFDocument(inputStream);
        } catch (IOException ex) {
            throw new WordRenderException("Failed to load template resource: " + options.getTemplateResource(), ex);
        }
    }

    private void configurePage(XWPFDocument document, WordRenderStyleDefinition styleDefinition) {
        CTSectPr sectPr = document.getDocument().getBody().isSetSectPr()
            ? document.getDocument().getBody().getSectPr()
            : document.getDocument().getBody().addNewSectPr();
        CTPageMar pageMar = sectPr.isSetPgMar() ? sectPr.getPgMar() : sectPr.addNewPgMar();
        pageMar.setLeft(BigInteger.valueOf(styleDefinition.getLeftMarginTwips()));
        pageMar.setRight(BigInteger.valueOf(styleDefinition.getRightMarginTwips()));
        pageMar.setTop(BigInteger.valueOf(styleDefinition.getTopMarginTwips()));
        pageMar.setBottom(BigInteger.valueOf(styleDefinition.getBottomMarginTwips()));
        CTPageSz pageSz = sectPr.isSetPgSz() ? sectPr.getPgSz() : sectPr.addNewPgSz();
        if (properties.getPageSize() == WordRenderPageSize.LETTER) {
            pageSz.setW(BigInteger.valueOf(12240));
            pageSz.setH(BigInteger.valueOf(15840));
        } else {
            pageSz.setW(BigInteger.valueOf(11906));
            pageSz.setH(BigInteger.valueOf(16838));
        }
    }

    private void configureMetadata(XWPFDocument document, WordRenderOptions options) {
        POIXMLProperties.CoreProperties coreProperties = document.getProperties().getCoreProperties();
        if (options.getTitle() != null) {
            coreProperties.setTitle(options.getTitle());
        }
        if (options.getAuthor() != null) {
            coreProperties.setCreator(options.getAuthor());
        }
        coreProperties.setDescription("Generated by WordRender");
    }

    private void configureStyles(XWPFDocument document, WordRenderStyleDefinition styleDefinition) {
        XWPFStyles styles = document.getStyles();
        if (styles == null) {
            styles = document.createStyles();
        }
        ensureNormalStyle(styles, styleDefinition);
        ensureParagraphStyle(styles, "Heading1", "Heading 1", styleDefinition.getHeadingOneFontSize(),
            styleDefinition.getHeadingColor(), 0);
        ensureParagraphStyle(styles, "Heading2", "Heading 2", styleDefinition.getHeadingTwoFontSize(),
            styleDefinition.getHeadingColor(), 1);
        ensureParagraphStyle(styles, "Heading3", "Heading 3", styleDefinition.getHeadingThreeFontSize(),
            styleDefinition.getHeadingColor(), 2);
        ensureParagraphStyle(styles, "Heading4", "Heading 4", styleDefinition.getHeadingFourFontSize(),
            styleDefinition.getHeadingColor(), 3);
        ensureParagraphStyle(styles, "Heading5", "Heading 5", styleDefinition.getHeadingFiveFontSize(),
            styleDefinition.getHeadingColor(), 4);
        ensureParagraphStyle(styles, "Heading6", "Heading 6", styleDefinition.getHeadingSixFontSize(),
            styleDefinition.getHeadingColor(), 5);
    }

    private void ensureNormalStyle(XWPFStyles styles, WordRenderStyleDefinition styleDefinition) {
        if (styles.styleExist("Normal")) {
            return;
        }
        CTStyle normal = CTStyle.Factory.newInstance();
        normal.setStyleId("Normal");
        normal.setType(STStyleType.PARAGRAPH);
        normal.setDefault(true);
        normal.addNewName().setVal("Normal");
        normal.addNewQFormat();

        CTRPr runProperties = normal.addNewRPr();
        CTFonts fonts = runProperties.addNewRFonts();
        fonts.setAscii(fontResolver.getDefaultFontFamily());
        fonts.setHAnsi(fontResolver.getDefaultFontFamily());
        fonts.setCs(fontResolver.getDefaultFontFamily());
        fonts.setEastAsia(fontResolver.getDefaultFontFamily());

        CTHpsMeasure size = runProperties.addNewSz();
        size.setVal(BigInteger.valueOf(styleDefinition.getBodyFontSize() * 2L));
        CTHpsMeasure sizeCs = runProperties.addNewSzCs();
        sizeCs.setVal(BigInteger.valueOf(styleDefinition.getBodyFontSize() * 2L));

        styles.addStyle(new XWPFStyle(normal));
    }

    private void ensureParagraphStyle(XWPFStyles styles, String styleId, String name, int fontSize, String color,
                                      int outlineLevel) {
        if (styles.styleExist(styleId)) {
            return;
        }
        CTStyle ctStyle = CTStyle.Factory.newInstance();
        ctStyle.setStyleId(styleId);
        ctStyle.setType(STStyleType.PARAGRAPH);
        ctStyle.setCustomStyle(false);

        CTString styleName = ctStyle.addNewName();
        styleName.setVal(name);

        CTString basedOn = ctStyle.addNewBasedOn();
        basedOn.setVal("Normal");

        CTString next = ctStyle.addNewNext();
        next.setVal("Normal");

        CTDecimalNumber priority = ctStyle.addNewUiPriority();
        priority.setVal(BigInteger.valueOf(9 + outlineLevel));

        ctStyle.addNewQFormat();
        ctStyle.addNewUnhideWhenUsed();
        ctStyle.addNewAutoRedefine();

        CTDecimalNumber outline = ctStyle.addNewPPr().addNewOutlineLvl();
        outline.setVal(BigInteger.valueOf(outlineLevel));

        CTRPr runProperties = ctStyle.addNewRPr();
        CTFonts fonts = runProperties.addNewRFonts();
        fonts.setAscii(fontResolver.getDefaultFontFamily());
        fonts.setHAnsi(fontResolver.getDefaultFontFamily());
        fonts.setCs(fontResolver.getDefaultFontFamily());
        fonts.setEastAsia(fontResolver.getDefaultFontFamily());

        CTHpsMeasure size = runProperties.addNewSz();
        size.setVal(BigInteger.valueOf(fontSize * 2L));
        CTHpsMeasure sizeCs = runProperties.addNewSzCs();
        sizeCs.setVal(BigInteger.valueOf(fontSize * 2L));

        runProperties.addNewB();
        if (color != null) {
            runProperties.addNewColor().setVal(color);
        }

        styles.addStyle(new XWPFStyle(ctStyle));
    }

    private void addWatermark(XWPFDocument document, WordRenderOptions options) {
        if (!hasText(options.getWatermarkText())) {
            return;
        }
        XWPFHeaderFooterPolicy policy = new XWPFHeaderFooterPolicy(document);
        Map<STHdrFtr.Enum, List<String>> existingHeaderTexts = captureHeaderTexts(policy);
        policy.createWatermark(options.getWatermarkText());
        restoreHeaderTexts(policy, existingHeaderTexts);
        watermarkCustomizer.softenWatermarkHeaders(policy, options);
    }

    private void addHeader(XWPFDocument document, WordRenderOptions options, WordRenderStyleDefinition styleDefinition) {
        if (!hasText(options.getHeaderText())) {
            return;
        }
        XWPFHeaderFooterPolicy policy = new XWPFHeaderFooterPolicy(document);
        if (hasText(options.getWatermarkText())) {
            appendHeaderText(ensureHeader(policy, STHdrFtr.DEFAULT), options, styleDefinition);
            appendHeaderText(ensureHeader(policy, STHdrFtr.FIRST), options, styleDefinition);
            appendHeaderText(ensureHeader(policy, STHdrFtr.EVEN), options, styleDefinition);
            return;
        }
        appendHeaderText(ensureHeader(policy, STHdrFtr.DEFAULT), options, styleDefinition);
    }

    private void addFooter(XWPFDocument document, WordRenderOptions options, WordRenderStyleDefinition styleDefinition) {
        if (!hasText(options.getFooterText()) && !options.isShowPageNumber()) {
            return;
        }
        XWPFHeaderFooterPolicy policy = new XWPFHeaderFooterPolicy(document);
        XWPFFooter footer = ensureFooter(policy);
        XWPFParagraph paragraph = footer.createParagraph();
        paragraph.setSpacingBefore(0);
        paragraph.setSpacingAfter(0);
        paragraph.setAlignment(ParagraphAlignment.LEFT);

        if (hasText(options.getFooterText())) {
            XWPFRun run = paragraph.createRun();
            run.setFontFamily(styleDefinition.getFontFamily());
            run.setFontSize(resolveAuxiliaryFontSize(styleDefinition));
            run.setColor(resolveAuxiliaryColor());
            run.setText(options.getFooterText());
        }
        if (options.isShowPageNumber()) {
            addRightTabStop(paragraph, resolveRightTabPosition(styleDefinition));
            paragraph.createRun().addTab();

            appendPageField(paragraph, styleDefinition, " PAGE ");

            XWPFRun slash = paragraph.createRun();
            slash.setFontFamily(styleDefinition.getFontFamily());
            slash.setFontSize(resolveAuxiliaryFontSize(styleDefinition));
            slash.setColor(resolveAuxiliaryColor());
            slash.setText("/");

            appendPageField(paragraph, styleDefinition, " NUMPAGES ");
        }
    }

    private XWPFHeader ensureHeader(XWPFHeaderFooterPolicy policy, STHdrFtr.Enum type) {
        XWPFHeader header = policy.getHeader(type);
        return header != null ? header : policy.createHeader(type);
    }

    private XWPFFooter ensureFooter(XWPFHeaderFooterPolicy policy) {
        XWPFFooter footer = policy.getDefaultFooter();
        return footer != null ? footer : policy.createFooter(STHdrFtr.DEFAULT);
    }

    private void appendHeaderText(XWPFHeader header, WordRenderOptions options, WordRenderStyleDefinition styleDefinition) {
        XWPFParagraph paragraph = header.createParagraph();
        paragraph.setAlignment(toAlignment(options.getHeaderPosition()));
        XWPFRun run = paragraph.createRun();
        run.setFontFamily(styleDefinition.getFontFamily());
        run.setFontSize(resolveAuxiliaryFontSize(styleDefinition));
        run.setColor(resolveAuxiliaryColor());
        run.setText(options.getHeaderText());
    }

    private Map<STHdrFtr.Enum, List<String>> captureHeaderTexts(XWPFHeaderFooterPolicy policy) {
        Map<STHdrFtr.Enum, List<String>> snapshot = new HashMap<STHdrFtr.Enum, List<String>>();
        snapshot.put(STHdrFtr.DEFAULT, extractHeaderTexts(policy.getDefaultHeader()));
        snapshot.put(STHdrFtr.FIRST, extractHeaderTexts(policy.getFirstPageHeader()));
        snapshot.put(STHdrFtr.EVEN, extractHeaderTexts(policy.getEvenPageHeader()));
        return snapshot;
    }

    private List<String> extractHeaderTexts(XWPFHeader header) {
        List<String> texts = new ArrayList<String>();
        if (header == null) {
            return texts;
        }
        for (XWPFParagraph paragraph : header.getParagraphs()) {
            if (hasText(paragraph.getText())) {
                texts.add(paragraph.getText());
            }
        }
        return texts;
    }

    private void restoreHeaderTexts(XWPFHeaderFooterPolicy policy, Map<STHdrFtr.Enum, List<String>> snapshot) {
        appendRestoredHeaderTexts(policy.getDefaultHeader(), snapshot.get(STHdrFtr.DEFAULT));
        appendRestoredHeaderTexts(policy.getFirstPageHeader(), snapshot.get(STHdrFtr.FIRST));
        appendRestoredHeaderTexts(policy.getEvenPageHeader(), snapshot.get(STHdrFtr.EVEN));
    }

    private void appendRestoredHeaderTexts(XWPFHeader header, List<String> texts) {
        if (header == null || texts == null || texts.isEmpty()) {
            return;
        }
        for (String text : texts) {
            XWPFParagraph paragraph = header.createParagraph();
            paragraph.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun run = paragraph.createRun();
            run.setText(text);
            run.setColor(resolveAuxiliaryColor());
        }
    }

    private void addCoverIfNeeded(XWPFDocument document, WordRenderOptions options, WordRenderStyleDefinition styleDefinition) {
        if (hasTemplate(options) || !shouldRenderCover(options)) {
            return;
        }
        XWPFParagraph titleParagraph = WordRenderPoiSupport.createParagraph(document, ParagraphAlignment.CENTER, 280);
        titleParagraph.setSpacingBefore(1200);
        XWPFRun titleRun = WordRenderPoiSupport.createRun(titleParagraph, styleDefinition.getFontFamily(),
            styleDefinition.getCoverTitleFontSize(), styleDefinition.getTitleColor(), true);
        titleRun.setText(options.getTitle());

        if (options.getSubTitle() != null && !options.getSubTitle().trim().isEmpty()) {
            XWPFParagraph subTitleParagraph = WordRenderPoiSupport.createParagraph(document, ParagraphAlignment.CENTER, 220);
            XWPFRun subTitleRun = WordRenderPoiSupport.createRun(subTitleParagraph, styleDefinition.getFontFamily(),
                styleDefinition.getHeadingTwoFontSize(), styleDefinition.getAccentColor(), false);
            subTitleRun.setText(options.getSubTitle());
        }

        XWPFParagraph infoParagraph = WordRenderPoiSupport.createParagraph(document, ParagraphAlignment.CENTER, 120);
        XWPFRun infoRun = WordRenderPoiSupport.createRun(infoParagraph, styleDefinition.getFontFamily(),
            styleDefinition.getBodyFontSize(), null, false);
        StringBuilder builder = new StringBuilder();
        if (options.getAuthor() != null) {
            builder.append("Author: ").append(options.getAuthor());
        }
        if (builder.length() > 0) {
            builder.append(" | ");
        }
        builder.append("Date: ").append(LocalDate.now().format(DateTimeFormatter.ISO_DATE));
        infoRun.setText(builder.toString());

        for (Map.Entry<String, String> entry : options.getMetadata().entrySet()) {
            XWPFParagraph metadataParagraph = WordRenderPoiSupport.createParagraph(document, ParagraphAlignment.CENTER, 40);
            XWPFRun metadataRun = WordRenderPoiSupport.createRun(metadataParagraph, styleDefinition.getFontFamily(),
                styleDefinition.getBodyFontSize(), null, false);
            metadataRun.setText(entry.getKey() + ": " + entry.getValue());
        }

        XWPFParagraph separatorParagraph = WordRenderPoiSupport.createParagraph(document, ParagraphAlignment.CENTER, 160);
        XWPFRun separatorRun = WordRenderPoiSupport.createRun(separatorParagraph, styleDefinition.getFontFamily(),
            styleDefinition.getHeadingThreeFontSize(), styleDefinition.getAccentColor(), false);
        separatorRun.setText("Generated by WordRender");

        XWPFParagraph pageBreak = document.createParagraph();
        pageBreak.createRun().addBreak(BreakType.PAGE);
    }

    private boolean shouldRenderCover(WordRenderOptions options) {
        return hasText(options.getTitle())
            || hasText(options.getSubTitle())
            || hasText(options.getAuthor())
            || !options.getMetadata().isEmpty();
    }

    private boolean hasText(String value) {
        return value != null && !value.trim().isEmpty();
    }

    private ParagraphAlignment toAlignment(WordRenderPosition position) {
        if (position == WordRenderPosition.LEFT) {
            return ParagraphAlignment.LEFT;
        }
        if (position == WordRenderPosition.CENTER) {
            return ParagraphAlignment.CENTER;
        }
        return ParagraphAlignment.RIGHT;
    }

    private int resolveAuxiliaryFontSize(WordRenderStyleDefinition styleDefinition) {
        return Math.max(8, styleDefinition.getBodyFontSize() - 2);
    }

    private String resolveAuxiliaryColor() {
        return "7F7F7F";
    }

    private void addRightTabStop(XWPFParagraph paragraph, BigInteger position) {
        if (!paragraph.getCTP().isSetPPr()) {
            paragraph.getCTP().addNewPPr();
        }
        if (!paragraph.getCTP().getPPr().isSetTabs()) {
            paragraph.getCTP().getPPr().addNewTabs();
        }
        org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTabStop tabStop =
            paragraph.getCTP().getPPr().getTabs().addNewTab();
        tabStop.setVal(STTabJc.RIGHT);
        tabStop.setLeader(STTabTlc.NONE);
        tabStop.setPos(position);
    }

    private BigInteger resolveRightTabPosition(WordRenderStyleDefinition styleDefinition) {
        BigInteger pageWidth = properties.getPageSize() == WordRenderPageSize.LETTER
            ? BigInteger.valueOf(12240)
            : BigInteger.valueOf(11906);
        BigInteger leftMargin = BigInteger.valueOf(styleDefinition.getLeftMarginTwips());
        BigInteger rightMargin = BigInteger.valueOf(styleDefinition.getRightMarginTwips());
        return pageWidth.subtract(leftMargin).subtract(rightMargin);
    }

    private void appendPageField(XWPFParagraph paragraph, WordRenderStyleDefinition styleDefinition, String instruction) {
        XWPFRun placeholder = paragraph.createRun();
        placeholder.setFontFamily(styleDefinition.getFontFamily());
        placeholder.setFontSize(resolveAuxiliaryFontSize(styleDefinition));
        placeholder.setColor(resolveAuxiliaryColor());

        CTSimpleField field = paragraph.getCTP().addNewFldSimple();
        field.setInstr(instruction);
    }
}
