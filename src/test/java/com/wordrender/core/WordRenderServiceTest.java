package com.wordrender.core;

import com.wordrender.autoconfigure.WordRenderAutoConfiguration;
import com.wordrender.model.WordRenderContentType;
import com.wordrender.model.WordRenderOptions;
import com.wordrender.model.WordRenderPosition;
import com.wordrender.model.WordRenderReportStyle;
import com.wordrender.model.WordRenderTemplateBinding;
import com.wordrender.model.WordRenderTemplateMode;
import com.wordrender.autoconfigure.WordRenderProperties;

import java.awt.Font;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.runner.ApplicationContextRunner;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

class WordRenderServiceTest {

    private static final String[] OUTPUT_TEMPLATE_ARTIFACTS = new String[] {
        "fixed-template.docx",
        "placeholder-single.docx",
        "placeholder-multiple.docx",
        "placeholder-table-cell.docx",
        "placeholder-missing.docx",
        "complex-template-placeholder-base.docx"
    };

    private final ApplicationContextRunner contextRunner = new ApplicationContextRunner()
            .withUserConfiguration(WordRenderAutoConfiguration.class);

    @Test
    void shouldRenderMarkdownDocx() {
        contextRunner.run(context -> {
            WordRenderService service = context.getBean(WordRenderService.class);
            byte[] bytes = service.renderDocx(
                    "# 标题\n\n这是 **粗体** 文本。\n\n- 项目一\n- 项目二",
                    WordRenderOptions.builder()
                            .contentType(WordRenderContentType.MARKDOWN)
                            .reportStyle(WordRenderReportStyle.FORMAL)
                            .title("Markdown 报告")
                            .headerText("Markdown 页眉")
                            .footerText("WordRender Footer")
                            .showPageNumber(true)
                            .headerPosition(WordRenderPosition.RIGHT)
                            .author("tester")
                            .build()
            );

            assertThat(bytes).isNotEmpty();
            XWPFDocument document = new XWPFDocument(new ByteArrayInputStream(bytes));
            assertThat(document.getParagraphs()).isNotEmpty();
            assertThat(document.getParagraphs().get(0).getText()).contains("Markdown 报告");
            assertThat(document.getHeaderList()).hasSize(1);
            assertThat(document.getHeaderList().get(0).getText()).contains("Markdown 页眉");
            assertThat(document.getFooterList()).hasSize(1);
            assertThat(document.getFooterList().get(0).getText()).contains("WordRender Footer");
            document.close();
        });
    }

    @Test
    void shouldRenderPlainTextToFile() {
        contextRunner.run(context -> {
            WordRenderService service = context.getBean(WordRenderService.class);
            Path tempDir = outputDir();
            Files.createDirectories(tempDir);
            Path tempFile = tempDir.resolve("plain-text-report.docx");
            service.renderDocxToFile(
                    "第一段内容\n第二行\n\n第二段内容",
                    WordRenderOptions.builder()
                            .contentType(WordRenderContentType.PLAIN_TEXT)
                            .title("普通文本报告")
                            .build(),
                    tempFile
            );
            assertThat(Files.exists(tempFile)).isTrue();
            assertThat(Files.size(tempFile)).isGreaterThan(0L);
        });
    }

    @Test
    void shouldRenderMarkdownToStream() {
        contextRunner.run(context -> {
            WordRenderService service = context.getBean(WordRenderService.class);
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

            service.renderDocxToStream(
                "# 流式输出报告\n\n## 结论\n\n- 适合生产导出\n- 避免额外 byte[] 副本",
                WordRenderOptions.builder()
                    .contentType(WordRenderContentType.MARKDOWN)
                    .reportStyle(WordRenderReportStyle.FORMAL)
                    .title("流式输出报告")
                    .footerText("stream")
                    .showPageNumber(true)
                    .build(),
                outputStream
            );

            byte[] bytes = outputStream.toByteArray();
            assertThat(bytes).isNotEmpty();
            XWPFDocument document = new XWPFDocument(new ByteArrayInputStream(bytes));
            assertThat(document.getParagraphs().stream().anyMatch(p -> p.getText().contains("流式输出报告"))).isTrue();
            document.close();
        });
    }

    @Test
    void shouldRenderMarkdownToProjectTempDirectory() {
        contextRunner.run(context -> {
            WordRenderService service = context.getBean(WordRenderService.class);
            Path tempDir = outputDir();
            Files.createDirectories(tempDir);
            Path outputFile = tempDir.resolve("markdown-report.docx");

            service.renderDocxToFile(
                    "# 项目临时目录报告\n\n## 结论\n\n- 测试生成文件\n- 校验项目临时目录输出",
                    WordRenderOptions.builder()
                            .contentType(WordRenderContentType.MARKDOWN)
                            .reportStyle(WordRenderReportStyle.FORMAL)
                            .title("项目临时目录报告")
                            .headerText("Markdown 页眉")
                            .footerText("WordRender Footer")
                            .showPageNumber(true)
                            .headerPosition(WordRenderPosition.RIGHT)
                            .author("tester")
                            .build(),
                    outputFile
            );

            assertThat(Files.exists(outputFile)).isTrue();
            assertThat(Files.size(outputFile)).isGreaterThan(0L);

            XWPFDocument document = new XWPFDocument(Files.newInputStream(outputFile));
            assertThat(document.getParagraphs()).isNotEmpty();
            assertThat(document.getParagraphs().get(0).getText()).contains("项目临时目录报告");
            document.close();
        });
    }

    @Test
    void shouldGenerateComplexMarkdownReportToFile() {
        contextRunner.run(context -> {
            WordRenderService service = context.getBean(WordRenderService.class);
            Path tempDir = outputDir();
            Files.createDirectories(tempDir);
            Path outputFile = tempDir.resolve("complex-markdown-report.docx");

            service.renderDocxToFile(
                complexMarkdownContent(),
                WordRenderOptions.builder()
                    .contentType(WordRenderContentType.MARKDOWN)
                    .reportStyle(WordRenderReportStyle.FORMAL)
                    .title("复杂样式 Markdown 报告")
                    .subTitle("覆盖常见结构与复杂层级")
                    .headerText("复杂 Markdown 样例")
                    .footerText("WordRender Complex Sample")
                    .showPageNumber(true)
                    .headerPosition(WordRenderPosition.RIGHT)
                    .author("test-suite")
                    .metadata("sample", "complex-markdown")
                    .build(),
                outputFile
            );

            assertThat(Files.exists(outputFile)).isTrue();
            assertThat(Files.size(outputFile)).isGreaterThan(0L);

            XWPFDocument document = new XWPFDocument(Files.newInputStream(outputFile));
            assertThat(document.getParagraphs().stream().anyMatch(p -> p.getText().contains("复杂样式 Markdown 报告"))).isTrue();
            assertThat(document.getParagraphs().stream().anyMatch(p -> p.getText().contains("一级概览"))).isTrue();
            assertThat(document.getParagraphs().stream().anyMatch(p -> p.getText().contains("五级细节"))).isTrue();
            assertThat(document.getParagraphs().stream()
                .filter(p -> p.getText().contains("五级细节"))
                .findFirst()
                .map(p -> p.getStyle())
                .orElse(null)).isEqualTo("Heading5");
            assertThat(document.getParagraphs().stream().anyMatch(p -> p.getText().contains("需要特别关注北方区域回款周期"))).isTrue();
            assertThat(document.getParagraphs().stream().anyMatch(p -> p.getText().contains("System.out.println"))).isTrue();
            assertThat(document.getTables()).isNotEmpty();
            assertThat(document.getTables().stream()
                .anyMatch(table -> table.getText().contains("区域")
                    && table.getText().contains("182")
                    && table.getText().contains("18.7%")
                    && table.getText().contains("华东"))).isTrue();
            document.close();
        });
    }

    @Test
    void shouldBindProperties() {
        String fontPath = existingFontPath();
        String fontFamily = loadFontFamily(fontPath);
        new ApplicationContextRunner()
                .withUserConfiguration(WordRenderAutoConfiguration.class)
                .withPropertyValues(
                        "word-render.default-style=FORMAL",
                        "word-render.default-font-family=Microsoft YaHei",
                        "word-render.default-font-size=12",
                        "word-render.fonts.regular-path=file:" + fontPath,
                        "word-render.fonts.bold-path=file:" + fontPath,
                        "word-render.fonts.default-family=Fake Font, Microsoft YaHei, SimSun, sans-serif",
                        "word-render.colors.title-color=202020",
                        "word-render.colors.heading-color=202020",
                        "word-render.colors.accent-color=666666",
                        "word-render.watermark.enabled=true",
                        "word-render.watermark.text=内部资料",
                        "word-render.watermark.color=BFBFBF",
                        "word-render.watermark.font-size=22",
                        "word-render.cache.enabled=true",
                        "word-render.cache.type=MEMORY",
                        "word-render.cache.redis-key-prefix=word-render:"
                )
                .run(context -> {
                    WordRenderProperties properties = context.getBean(WordRenderProperties.class);
                    WordRenderService service = context.getBean(WordRenderService.class);
                    byte[] bytes = service.renderDocx(
                            "simple text",
                            WordRenderOptions.builder()
                                    .contentType(WordRenderContentType.PLAIN_TEXT)
                                    .title("Props")
                                    .build()
                    );
                    assertThat(properties.resolveDefaultFontFamily()).isEqualTo("Fake Font");
                    assertThat(properties.getFonts().getRegularPath()).isEqualTo("file:" + fontPath);
                    assertThat(properties.getColors().getHeadingColor()).isEqualTo("202020");
                    assertThat(properties.getWatermark().getText()).isEqualTo("内部资料");
                    assertThat(properties.getWatermark().getColor()).isEqualTo("BFBFBF");
                    assertThat(properties.getWatermark().getFontSize()).isEqualTo(22);
                    assertThat(readZipEntry(bytes, "word/styles.xml")).contains(fontFamily);
                    assertThat(bytes).isNotEmpty();
                });
    }

    @Test
    void shouldApplyCustomHeadingColorsAndWatermark() {
        contextRunner.run(context -> {
            WordRenderService service = context.getBean(WordRenderService.class);
            Path tempDir = outputDir();
            Files.createDirectories(tempDir);
            Path outputFile = tempDir.resolve("watermark-report.docx");

            service.renderDocxToFile(
                "# 一级标题\n\n## 二级标题\n\n正文内容",
                WordRenderOptions.builder()
                    .contentType(WordRenderContentType.MARKDOWN)
                    .title("颜色与水印验证")
                    .titleColor("#111111")
                    .headingColor("#222222")
                    .accentColor("#777777")
                    .watermarkText("内部资料")
                    .watermarkColor("#BFBFBF")
                    .watermarkFontSize(24)
                    .build(),
                outputFile
            );

            assertThat(Files.exists(outputFile)).isTrue();
            assertThat(Files.size(outputFile)).isGreaterThan(0L);

            byte[] bytes = Files.readAllBytes(outputFile);

            try (XWPFDocument document = new XWPFDocument(new ByteArrayInputStream(bytes))) {
                assertThat(document.getParagraphs().stream()
                    .filter(p -> p.getText().contains("一级标题"))
                    .findFirst()
                    .map(p -> p.getStyle())
                    .orElse(null)).isEqualTo("Heading1");
            }

            assertThat(readZipEntry(bytes, "word/styles.xml")).contains("w:val=\"222222\"");
            String headerXml = readZipEntry(bytes, "word/header1.xml");
            assertThat(headerXml).contains("内部资料");
            assertThat(headerXml).contains("v:shape");
            assertThat(headerXml).contains("fillcolor=\"#BFBFBF\"");
            assertThat(headerXml).contains("font-size:24pt");
            assertThat(headerXml).contains("rotation:315");
            assertThat(headerXml).contains("opacity=\"0.10\"");
        });
    }

    @Test
    void shouldRejectInvalidColorValue() {
        contextRunner.run(context -> {
            WordRenderService service = context.getBean(WordRenderService.class);

            assertThatThrownBy(() -> service.renderDocx(
                "# 标题",
                WordRenderOptions.builder()
                    .contentType(WordRenderContentType.MARKDOWN)
                    .headingColor("blue")
                    .build()
            )).isInstanceOf(WordRenderException.class)
                .hasMessageContaining("headingColor");
        });
    }

    @Test
    void shouldRejectInvalidWatermarkColorFromProperties() {
        new ApplicationContextRunner()
            .withUserConfiguration(WordRenderAutoConfiguration.class)
            .withPropertyValues(
                "word-render.watermark.enabled=true",
                "word-render.watermark.text=内部资料",
                "word-render.watermark.color=invalid"
            )
            .run(context -> {
                WordRenderService service = context.getBean(WordRenderService.class);

                assertThatThrownBy(() -> service.renderDocx(
                    "simple text",
                    WordRenderOptions.builder()
                        .contentType(WordRenderContentType.PLAIN_TEXT)
                        .title("Props watermark")
                        .build()
                )).isInstanceOf(WordRenderException.class)
                    .hasMessageContaining("watermarkColor");
            });
    }

    @Test
    void shouldRejectInvalidWatermarkFontSizeFromProperties() {
        new ApplicationContextRunner()
            .withUserConfiguration(WordRenderAutoConfiguration.class)
            .withPropertyValues(
                "word-render.watermark.enabled=true",
                "word-render.watermark.text=内部资料",
                "word-render.watermark.font-size=100"
            )
            .run(context -> {
                WordRenderService service = context.getBean(WordRenderService.class);

                assertThatThrownBy(() -> service.renderDocx(
                    "simple text",
                    WordRenderOptions.builder()
                        .contentType(WordRenderContentType.PLAIN_TEXT)
                        .title("Props watermark")
                        .build()
                )).isInstanceOf(WordRenderException.class)
                    .hasMessageContaining("watermarkFontSize");
            });
    }

    @Test
    void shouldKeepHeaderTextWhenWatermarkIsEnabled() {
        contextRunner.run(context -> {
            WordRenderService service = context.getBean(WordRenderService.class);
            byte[] bytes = service.renderDocx(
                "# 标题\n\n正文",
                WordRenderOptions.builder()
                    .contentType(WordRenderContentType.MARKDOWN)
                    .headerText("正式页眉")
                    .headerPosition(WordRenderPosition.RIGHT)
                    .watermarkText("内部资料")
                    .build()
            );

            String defaultHeader = readZipEntry(bytes, "word/header1.xml");
            assertThat(defaultHeader).contains("内部资料");
            assertThat(defaultHeader).contains("正式页眉");
        });
    }

    @Test
    void shouldNotLeavePartialTargetFileWhenFileRenderFails() throws Exception {
        contextRunner.run(context -> {
            WordRenderService service = context.getBean(WordRenderService.class);
            Path targetFile = outputDir().resolve("failed-report.docx");
            Files.deleteIfExists(targetFile);

            assertThatThrownBy(() -> service.renderDocxToFile(
                "# 标题",
                WordRenderOptions.builder()
                    .contentType(WordRenderContentType.MARKDOWN)
                    .watermarkText("内部资料")
                    .watermarkColor("bad")
                    .build(),
                targetFile
            )).isInstanceOf(WordRenderException.class)
                .hasMessageContaining("watermarkColor");

            assertThat(Files.exists(targetFile)).isFalse();
        });
    }

    @Test
    void shouldNotCreateHeaderOrFooterWhenNotProvided() {
        contextRunner.run(context -> {
            WordRenderService service = context.getBean(WordRenderService.class);
            byte[] bytes = service.renderDocx(
                    "simple text body only",
                    WordRenderOptions.builder()
                            .contentType(WordRenderContentType.PLAIN_TEXT)
                            .build()
            );

            XWPFDocument document = new XWPFDocument(new ByteArrayInputStream(bytes));
            assertThat(document.getHeaderList()).isEmpty();
            assertThat(document.getFooterList()).isEmpty();
            document.close();
        });
    }

    @Test
    void shouldAppendDynamicContentAfterTemplateDocument() {
        contextRunner.run(context -> {
            WordRenderService service = context.getBean(WordRenderService.class);
            Path templatePath = createTemplate("fixed-template.docx", "固定模板首页", "固定模板第二段");

            byte[] bytes = service.renderDocx(
                "# 动态报告内容\n\n这是模板后的个性化部分。",
                WordRenderOptions.builder()
                    .contentType(WordRenderContentType.MARKDOWN)
                    .templateResource(templatePath.toUri().toString())
                    .appendPageBreakAfterTemplate(true)
                    .headerText("模板报告")
                    .showPageNumber(true)
                    .build()
            );

            XWPFDocument document = new XWPFDocument(new ByteArrayInputStream(bytes));
            assertThat(document.getParagraphs().stream().anyMatch(p -> p.getText().contains("固定模板首页"))).isTrue();
            assertThat(document.getParagraphs().stream().anyMatch(p -> p.getText().contains("动态报告内容"))).isTrue();
            document.close();
        });
    }

    @Test
    void shouldReplaceSinglePlaceholderInsideTemplate() {
        contextRunner.run(context -> {
            WordRenderService service = context.getBean(WordRenderService.class);
            Path templatePath = createTemplate("placeholder-single.docx", "固定封面", "${wr_content}", "固定附录");

            byte[] bytes = service.renderDocx(
                "",
                WordRenderOptions.builder()
                    .templateResource(templatePath.toUri().toString())
                    .templateMode(WordRenderTemplateMode.PLACEHOLDER)
                    .templateBinding("wr_content", WordRenderTemplateBinding.builder()
                        .content("# 动态正文\n\n- 第一项\n- 第二项")
                        .contentType(WordRenderContentType.MARKDOWN)
                        .build())
                    .build()
            );

            XWPFDocument document = new XWPFDocument(new ByteArrayInputStream(bytes));
            assertThat(document.getParagraphs().stream().anyMatch(p -> p.getText().contains("固定封面"))).isTrue();
            assertThat(document.getParagraphs().stream().anyMatch(p -> p.getText().contains("动态正文"))).isTrue();
            assertThat(document.getParagraphs().stream().anyMatch(p -> p.getText().contains("${wr_content}"))).isFalse();
            assertThat(document.getParagraphs().stream().anyMatch(p -> p.getText().contains("固定附录"))).isTrue();
            document.close();
        });
    }

    @Test
    void shouldReplaceInlinePlaceholderSplitAcrossRuns() {
        contextRunner.run(context -> {
            WordRenderService service = context.getBean(WordRenderService.class);
            Path fixtureDir = fixtureDir();
            Path templatePath = fixtureDir.resolve("placeholder-inline-runs.docx");

            try (XWPFDocument templateDocument = new XWPFDocument();
                 OutputStream outputStream = Files.newOutputStream(templatePath)) {
                XWPFParagraph paragraph = templateDocument.createParagraph();
                paragraph.createRun().setText("授信申请人：");
                paragraph.createRun().setText("${");
                paragraph.createRun().setText("wr");
                paragraph.createRun().setText("_name}");
                templateDocument.write(outputStream);
            }

            byte[] bytes = service.renderDocx(
                "",
                WordRenderOptions.builder()
                    .templateResource(templatePath.toUri().toString())
                    .templateMode(WordRenderTemplateMode.PLACEHOLDER)
                    .templateBinding("wr_name", WordRenderTemplateBinding.builder()
                        .content("北京某科技有限公司")
                        .contentType(WordRenderContentType.PLAIN_TEXT)
                        .build())
                    .build()
            );

            try (XWPFDocument document = new XWPFDocument(new ByteArrayInputStream(bytes))) {
                assertThat(document.getParagraphs().stream()
                    .anyMatch(p -> p.getText().contains("授信申请人：北京某科技有限公司"))).isTrue();
            }
        });
    }

    @Test
    void shouldReplaceMultiplePlaceholdersAndApplyBaseHeadingLevel() {
        contextRunner.run(context -> {
            WordRenderService service = context.getBean(WordRenderService.class);
            Path fixtureDir = fixtureDir();
            Path templatePath = fixtureDir.resolve("placeholder-multiple.docx");

            try (XWPFDocument templateDocument = new XWPFDocument();
                 OutputStream outputStream = Files.newOutputStream(templatePath)) {
                templateDocument.createParagraph().createRun().setText("一级章节");
                templateDocument.createParagraph().createRun().setText("二级章节");
                templateDocument.createParagraph().createRun().setText("${wr_summary}");
                templateDocument.createParagraph().createRun().setText("${wr_content}");
                templateDocument.createParagraph().createRun().setText("${wr_appendix}");
                templateDocument.write(outputStream);
            }

            byte[] bytes = service.renderDocx(
                "",
                WordRenderOptions.builder()
                    .templateResource(templatePath.toUri().toString())
                    .templateMode(WordRenderTemplateMode.PLACEHOLDER)
                    .templateBinding("wr_summary", WordRenderTemplateBinding.builder()
                        .content("这是摘要内容")
                        .contentType(WordRenderContentType.PLAIN_TEXT)
                        .build())
                    .templateBinding("wr_content", WordRenderTemplateBinding.builder()
                        .content("# 动态标题\n\n正文段落")
                        .contentType(WordRenderContentType.MARKDOWN)
                        .baseHeadingLevel(3)
                        .build())
                    .templateBinding("wr_appendix", WordRenderTemplateBinding.builder()
                        .content("附录内容")
                        .contentType(WordRenderContentType.PLAIN_TEXT)
                        .build())
                    .build()
            );

            XWPFDocument document = new XWPFDocument(new ByteArrayInputStream(bytes));
            assertThat(document.getParagraphs().stream().anyMatch(p -> p.getText().contains("这是摘要内容"))).isTrue();
            assertThat(document.getParagraphs().stream().anyMatch(p -> p.getText().contains("附录内容"))).isTrue();
            assertThat(document.getParagraphs().stream()
                .filter(p -> p.getText().contains("动态标题"))
                .findFirst()
                .map(p -> p.getStyle())
                .orElse(null)).isEqualTo("Heading3");
            document.close();
        });
    }

    @Test
    void shouldReplacePlaceholderInsideTableCell() {
        contextRunner.run(context -> {
            WordRenderService service = context.getBean(WordRenderService.class);
            Path fixtureDir = fixtureDir();
            Path templatePath = fixtureDir.resolve("placeholder-table-cell.docx");

            try (XWPFDocument templateDocument = new XWPFDocument();
                 OutputStream outputStream = Files.newOutputStream(templatePath)) {
                templateDocument.createParagraph().createRun().setText("固定模板首页");
                templateDocument.createTable(1, 1).getRow(0).getCell(0).setText("${wr_content}");
                templateDocument.write(outputStream);
            }

            byte[] bytes = service.renderDocx(
                "",
                WordRenderOptions.builder()
                    .templateResource(templatePath.toUri().toString())
                    .templateMode(WordRenderTemplateMode.PLACEHOLDER)
                    .templateBinding("wr_content", WordRenderTemplateBinding.builder()
                        .content("表格中的动态内容")
                        .contentType(WordRenderContentType.PLAIN_TEXT)
                        .build())
                    .build()
            );

            XWPFDocument document = new XWPFDocument(new ByteArrayInputStream(bytes));
            assertThat(document.getTables()).isNotEmpty();
            assertThat(document.getTables().get(0).getRow(0).getCell(0).getText()).contains("表格中的动态内容");
            document.close();
        });
    }

    @Test
    void shouldGenerateComplexPlaceholderTemplateReportToFile() {
        contextRunner.run(context -> {
            WordRenderService service = context.getBean(WordRenderService.class);
            Path fixtureDir = fixtureDir();
            Path templatePath = fixtureDir.resolve("complex-template-placeholder-base.docx");
            Path outputFile = outputDir().resolve("complex-template-placeholder-report.docx");

            try (XWPFDocument templateDocument = new XWPFDocument();
                 OutputStream outputStream = Files.newOutputStream(templatePath)) {
                templateDocument.createParagraph().createRun().setText("固定封面页");
                templateDocument.createParagraph().createRun().setText("一级章节：经营总览");
                templateDocument.createParagraph().createRun().setText("二级章节：摘要区");
                templateDocument.createParagraph().createRun().setText("${wr_summary}");
                templateDocument.createParagraph().createRun().setText("二级章节：正文区");
                templateDocument.createParagraph().createRun().setText("${wr_content}");
                templateDocument.createParagraph().createRun().setText("固定附录页");
                templateDocument.createParagraph().createRun().setText("${wr_appendix}");
                templateDocument.write(outputStream);
            }

            byte[] bytes = service.renderDocx(
                "",
                WordRenderOptions.builder()
                    .templateResource(templatePath.toUri().toString())
                    .templateMode(WordRenderTemplateMode.PLACEHOLDER)
                    .headerText("复杂模板样例")
                    .footerText("WordRender Template Sample")
                    .showPageNumber(true)
                    .templateBinding("wr_summary", WordRenderTemplateBinding.builder()
                        .content("这是复杂模板样例的摘要部分，用于验证前置固定页和摘要占位共存。")
                        .contentType(WordRenderContentType.PLAIN_TEXT)
                        .build())
                    .templateBinding("wr_content", WordRenderTemplateBinding.builder()
                        .content(complexTemplateMarkdownContent())
                        .contentType(WordRenderContentType.MARKDOWN)
                        .baseHeadingLevel(3)
                        .build())
                    .templateBinding("wr_appendix", WordRenderTemplateBinding.builder()
                        .content("附录说明：该文档用于自动化测试与人工验收。")
                        .contentType(WordRenderContentType.PLAIN_TEXT)
                        .build())
                    .build()
            );

            Files.write(outputFile, bytes);
            assertThat(Files.exists(outputFile)).isTrue();
            assertThat(Files.size(outputFile)).isGreaterThan(0L);

            XWPFDocument document = new XWPFDocument(Files.newInputStream(outputFile));
            assertThat(document.getParagraphs().stream().anyMatch(p -> p.getText().contains("固定封面页"))).isTrue();
            assertThat(document.getParagraphs().stream().anyMatch(p -> p.getText().contains("固定附录页"))).isTrue();
            assertThat(document.getParagraphs().stream().anyMatch(p -> p.getText().contains("复杂模板样例的摘要部分"))).isTrue();
            assertThat(document.getParagraphs().stream().anyMatch(p -> p.getText().contains("三级正文标题"))).isTrue();
            assertThat(document.getParagraphs().stream()
                .filter(p -> p.getText().contains("三级正文标题"))
                .findFirst()
                .map(p -> p.getStyle())
                .orElse(null)).isEqualTo("Heading3");
            assertThat(document.getParagraphs().stream().anyMatch(p -> p.getText().contains("四级子标题"))).isTrue();
            assertThat(document.getParagraphs().stream()
                .filter(p -> p.getText().contains("四级子标题"))
                .findFirst()
                .map(p -> p.getStyle())
                .orElse(null)).isEqualTo("Heading4");
            assertThat(document.getTables()).isNotEmpty();
            assertThat(document.getTables().stream()
                .anyMatch(table -> table.getText().contains("华东") && table.getText().contains("高")
                    && table.getText().contains("建议优先跟进"))).isTrue();
            document.close();
        });
    }

    @Test
    void shouldPreserveTemplatePageBreakBeforePlaceholderContent() {
        contextRunner.run(context -> {
            WordRenderService service = context.getBean(WordRenderService.class);
            Path fixtureDir = fixtureDir();
            Path templatePath = fixtureDir.resolve("placeholder-page-break.docx");

            try (XWPFDocument templateDocument = new XWPFDocument();
                 OutputStream outputStream = Files.newOutputStream(templatePath)) {
                templateDocument.createParagraph().createRun().setText("固定首页");
                XWPFParagraph contentParagraph = templateDocument.createParagraph();
                contentParagraph.createRun().addBreak(BreakType.PAGE);
                contentParagraph.createRun().setText("${wr_content}");
                templateDocument.write(outputStream);
            }

            byte[] bytes = service.renderDocx(
                "",
                WordRenderOptions.builder()
                    .templateResource(templatePath.toUri().toString())
                    .templateMode(WordRenderTemplateMode.PLACEHOLDER)
                    .templateBinding("wr_content", WordRenderTemplateBinding.builder()
                        .content("# 动态内容标题\n\n正文内容")
                        .contentType(WordRenderContentType.MARKDOWN)
                        .build())
                    .build()
            );

            String documentXml = readZipEntry(bytes, "word/document.xml");
            assertThat(documentXml).contains("w:type=\"page\"");
            try (XWPFDocument document = new XWPFDocument(new ByteArrayInputStream(bytes))) {
                assertThat(document.getParagraphs().stream().anyMatch(p -> p.getText().contains("固定首页"))).isTrue();
                assertThat(document.getParagraphs().stream().anyMatch(p -> p.getText().contains("动态内容标题"))).isTrue();
            }
        });
    }

    @Test
    void shouldFailWhenPlaceholderBindingDoesNotExistInTemplate() {
        contextRunner.run(context -> {
            WordRenderService service = context.getBean(WordRenderService.class);
            Path templatePath = createTemplate("placeholder-missing.docx", "固定封面", "${wr_content}");

            assertThatThrownBy(() -> service.renderDocx(
                "",
                WordRenderOptions.builder()
                    .templateResource(templatePath.toUri().toString())
                    .templateMode(WordRenderTemplateMode.PLACEHOLDER)
                    .templateBinding("wr_unknown", WordRenderTemplateBinding.builder()
                        .content("不会命中的内容")
                        .contentType(WordRenderContentType.PLAIN_TEXT)
                        .build())
                    .build()
            )).isInstanceOf(WordRenderException.class)
                .hasMessageContaining("wr_unknown");
        });
    }

    private Path createTemplate(String fileName, String... paragraphs) throws Exception {
        Path fixtureDir = fixtureDir();
        Files.createDirectories(fixtureDir);
        Path templatePath = fixtureDir.resolve(fileName);
        try (XWPFDocument templateDocument = new XWPFDocument();
             OutputStream outputStream = Files.newOutputStream(templatePath)) {
            for (String paragraphText : paragraphs) {
                templateDocument.createParagraph().createRun().setText(paragraphText);
            }
            templateDocument.write(outputStream);
        }
        return templatePath;
    }

    private Path outputDir() throws Exception {
        Path outputDir = Paths.get("target", "word-render-test-output");
        Files.createDirectories(outputDir);
        for (String artifact : OUTPUT_TEMPLATE_ARTIFACTS) {
            Files.deleteIfExists(outputDir.resolve(artifact));
        }
        return outputDir;
    }

    private Path fixtureDir() throws Exception {
        Path fixtureDir = Paths.get("target", "word-render-test-fixtures");
        Files.createDirectories(fixtureDir);
        return fixtureDir;
    }

    private String readZipEntry(byte[] bytes, String entryName) throws Exception {
        try (ZipInputStream zipInputStream = new ZipInputStream(new ByteArrayInputStream(bytes))) {
            ZipEntry entry;
            while ((entry = zipInputStream.getNextEntry()) != null) {
                if (entryName.equals(entry.getName())) {
                    ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
                    byte[] buffer = new byte[1024];
                    int length;
                    while ((length = zipInputStream.read(buffer)) > 0) {
                        outputStream.write(buffer, 0, length);
                    }
                    return new String(outputStream.toByteArray(), StandardCharsets.UTF_8);
                }
            }
        }
        return "";
    }

    private String existingFontPath() {
        String[] candidates = new String[] {
            "/System/Library/Fonts/Geneva.ttf",
            "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
            "/System/Library/Fonts/SFNSMono.ttf",
            "/Library/Fonts/Arial Unicode.ttf"
        };
        for (String candidate : candidates) {
            if (Files.exists(Paths.get(candidate))) {
                return candidate;
            }
        }
        throw new IllegalStateException("No test font file found on this machine");
    }

    private String loadFontFamily(String fontPath) {
        try (InputStream inputStream = Files.newInputStream(Paths.get(fontPath))) {
            return Font.createFont(Font.TRUETYPE_FONT, inputStream).getFamily();
        } catch (Exception ex) {
            throw new IllegalStateException("Failed to load test font family from " + fontPath, ex);
        }
    }

    private String complexMarkdownContent() {
        return "# 一级概览\n\n"
            + "这是一个包含 **粗体**、*斜体* 和 `行内代码` 的复杂 Markdown 样例，用于验证 WordRender 的常见排版能力。\n\n"
            + "## 二级章节\n\n"
            + "1. 第一项包含中英文 mixed content\n"
            + "2. 第二项包含多级列表\n"
            + "   - 子项 A\n"
            + "   - 子项 B\n\n"
            + "### 三级分析\n\n"
            + "> 需要特别关注北方区域回款周期和重点项目的签约节奏。\n\n"
            + "#### 四级动作\n\n"
            + "- 建议一：补充销售支持资源\n"
            + "- 建议二：提升商机转化效率\n\n"
            + "##### 五级细节\n\n"
            + "---\n\n"
            + "```java\n"
            + "System.out.println(\"word-render\");\n"
            + "```\n\n"
            + "| 区域 | 线索数 | 转化率 | 风险等级 |\n"
            + "| --- | --- | --- | --- |\n"
            + "| 华东 | 182 | 18.7% | 低 |\n"
            + "| 华北 | 95 | 11.2% | 中 |\n"
            + "| 华南 | 136 | 16.4% | 低 |\n";
    }

    private String complexTemplateMarkdownContent() {
        return "# 三级正文标题\n\n"
            + "这是模板正文区的动态 Markdown，包含 **强调**、列表和表格。\n\n"
            + "## 四级子标题\n\n"
            + "- 动作项一\n"
            + "- 动作项二\n\n"
            + "| 区域 | 风险等级 | 建议 |\n"
            + "| --- | --- | --- |\n"
            + "| 华东 | 高 | 建议优先跟进 |\n"
            + "| 华南 | 中 | 建议保持观察 |\n";
    }
}
