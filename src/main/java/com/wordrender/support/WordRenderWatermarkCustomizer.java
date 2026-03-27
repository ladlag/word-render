package com.wordrender.support;

import com.wordrender.model.WordRenderOptions;
import javax.xml.namespace.QName;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;

/**
 * Applies a low-contrast watermark style after POI generates the default VML watermark.
 * POI exposes creation but not appearance tuning, so this class isolates the minimal
 * post-processing needed for readable business documents.
 */
public class WordRenderWatermarkCustomizer {

    private static final String VML_NAMESPACE = "urn:schemas-microsoft-com:vml";
    private static final QName STYLE_ATTRIBUTE = new QName("", "style");
    private static final QName FILLCOLOR_ATTRIBUTE = new QName("", "fillcolor");
    private static final QName STROKED_ATTRIBUTE = new QName("", "stroked");
    private static final QName OPACITY_ATTRIBUTE = new QName("", "opacity");
    private static final String WATERMARK_OPACITY = "0.10";
    private static final String WATERMARK_TEXT_STYLE_TEMPLATE = "font-family:&quot;%s&quot;;font-size:%dpt";
    private static final String WATERMARK_SHAPE_STYLE_TEMPLATE =
        "position:absolute;" +
        "margin-left:0;" +
        "margin-top:0;" +
        "width:%dpt;" +
        "height:%dpt;" +
        "z-index:-251654144;" +
        "mso-wrap-edited:f;" +
        "mso-position-horizontal:center;" +
        "mso-position-horizontal-relative:page;" +
        "mso-position-vertical:center;" +
        "mso-position-vertical-relative:page;" +
        "rotation:315";

    private final WordRenderFontResolver fontResolver;

    public WordRenderWatermarkCustomizer(WordRenderFontResolver fontResolver) {
        this.fontResolver = fontResolver;
    }

    public void softenWatermarkHeaders(XWPFHeaderFooterPolicy policy, WordRenderOptions options) {
        softenWatermarkHeader(policy.getDefaultHeader(), options);
        softenWatermarkHeader(policy.getFirstPageHeader(), options);
        softenWatermarkHeader(policy.getEvenPageHeader(), options);
    }

    private void softenWatermarkHeader(XWPFHeader header, WordRenderOptions options) {
        if (header == null) {
            return;
        }
        XmlObject[] shapes = header._getHdrFtr().selectPath(
            "declare namespace v='urn:schemas-microsoft-com:vml' .//v:shape"
        );
        for (XmlObject shape : shapes) {
            softenWatermarkShape(shape, options);
        }
    }

    private void softenWatermarkShape(XmlObject shape, WordRenderOptions options) {
        int fontSize = options.getWatermarkFontSize() == null ? 18 : options.getWatermarkFontSize().intValue();
        int shapeWidth = Math.max(220, fontSize * 15);
        int shapeHeight = Math.max(72, fontSize * 5);
        String shapeStyle = String.format(WATERMARK_SHAPE_STYLE_TEMPLATE, shapeWidth, shapeHeight);
        String fillColor = normalizeColor(options.getWatermarkColor());
        try (XmlCursor shapeCursor = shape.newCursor()) {
            shapeCursor.setAttributeText(STYLE_ATTRIBUTE, shapeStyle);
            shapeCursor.setAttributeText(FILLCOLOR_ATTRIBUTE, fillColor);
            shapeCursor.setAttributeText(STROKED_ATTRIBUTE, "f");
        }

        XmlObject[] textPaths = shape.selectPath(
            "declare namespace v='urn:schemas-microsoft-com:vml' .//v:textpath"
        );
        String textPathStyle = String.format(WATERMARK_TEXT_STYLE_TEMPLATE, fontResolver.getDefaultFontFamily(), fontSize);
        for (XmlObject textPath : textPaths) {
            try (XmlCursor textPathCursor = textPath.newCursor()) {
                textPathCursor.setAttributeText(STYLE_ATTRIBUTE, textPathStyle);
            }
        }

        XmlObject[] fills = shape.selectPath(
            "declare namespace v='urn:schemas-microsoft-com:vml' .//v:fill"
        );
        if (fills.length == 0) {
            try (XmlCursor shapeCursor = shape.newCursor()) {
                shapeCursor.toEndToken();
                shapeCursor.beginElement("fill", VML_NAMESPACE);
                shapeCursor.insertAttributeWithValue(OPACITY_ATTRIBUTE, WATERMARK_OPACITY);
            }
            return;
        }

        for (XmlObject fill : fills) {
            try (XmlCursor fillCursor = fill.newCursor()) {
                fillCursor.setAttributeText(OPACITY_ATTRIBUTE, WATERMARK_OPACITY);
            }
        }
    }

    private String normalizeColor(String color) {
        if (color == null || color.trim().isEmpty()) {
            return "#D9D9D9";
        }
        String normalized = color.trim().replace("#", "").toUpperCase();
        return "#" + normalized;
    }
}
