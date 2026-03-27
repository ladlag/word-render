# WordRender

`WordRender` 是一个基于 `JDK 8 + Spring Boot 2.7.x` 的零模板编码 DOCX 报告生成 SDK。业务方只需要传入文本内容和少量选项，不需要编写 Word 模板代码，也不需要理解 Apache POI。

## Features

- 支持 Markdown 长文本生成 `docx`
- 支持普通文本生成 `docx`
- 支持直接写入 `OutputStream`，更适合生产导出场景
- 内置 `DEFAULT`、`FORMAL`、`SIMPLE` 三套报告风格
- 支持通过 `application.yaml` 调整默认风格、字体、缓存策略
- 支持统一正式风格的标题颜色，并可通过配置或单次调用覆盖颜色
- 支持按需传入页眉、页脚和 `1/N` 页码
- 支持文字水印
- 支持可选模板资源，先保留固定前置页，再追加动态内容
- 支持 `docx` 模板占位填充，可在模板正文、中间页、附录页插入动态块内容
- 支持模板占位内容的标题基准级别控制，例如动态正文从三级标题开始
- Redis 缓存为可选增强项，默认不开启
- 不依赖 HTML 转 Word 主链路

## 维护边界

组件整体遵循“优先 Apache POI/XWPF 官方 API”的原则：

- 正文、标题、页眉、页脚、分页、模板加载等主流程优先使用官方 API
- 只有官方 API 无法覆盖时，才下沉到 XMLBeans / 底层 OOXML 能力

当前仍然保留少量底层处理的场景：

- 水印样式微调：POI 只提供 `createWatermark(...)`，但不开放颜色、透明度、字号和旋转角度
- 模板占位块替换：需要在原占位段落处插入多段正文、列表、表格时，需要 cursor 定位
- Word 版式细节：如页码域、标题样式定义、表格列宽/底纹等，POI 高层 API 不完整

这些下沉逻辑已经集中隔离在少数组件里，避免散落在主流程中：

- [WordRenderWatermarkCustomizer.java](/Users/ben/repo/codex/AITaskOS/word-render/src/main/java/com/wordrender/support/WordRenderWatermarkCustomizer.java)
- [WordRenderTemplatePlaceholderProcessor.java](/Users/ben/repo/codex/AITaskOS/word-render/src/main/java/com/wordrender/core/WordRenderTemplatePlaceholderProcessor.java)
- [WordRenderPoiSupport.java](/Users/ben/repo/codex/AITaskOS/word-render/src/main/java/com/wordrender/support/WordRenderPoiSupport.java)

## Maven

```xml
<dependency>
    <groupId>com.wordrender</groupId>
    <artifactId>word-render</artifactId>
    <version>0.1.0-SNAPSHOT</version>
</dependency>
```

## application.yaml

```yaml
word-render:
  enabled: true
  default-style: FORMAL
  default-font-family: Microsoft YaHei
  default-font-size: 11
  page-size: A4
  fonts:
    regular-path: classpath:/fonts/HarmonyOS_SansSC_Regular.ttf
    bold-path: classpath:/fonts/HarmonyOS_SansSC_Bold.ttf
    default-family: Microsoft YaHei, PingFang SC, Noto Sans SC, SimSun, Arial Unicode MS, sans-serif
  colors:
    title-color: "1F1F1F"
    heading-color: "1F1F1F"
    accent-color: "7F7F7F"
  watermark:
    enabled: true
    text: "内部资料"
    color: "BFBFBF"
    font-size: 22
  cache:
    enabled: true
    type: MEMORY
    redis-key-prefix: "word-render:"
```

如果使用 Redis 共享样式缓存：

```yaml
spring:
  redis:
    host: 127.0.0.1
    port: 6379

word-render:
  cache:
    enabled: true
    type: REDIS
```

如果你不准备分发自定义字体文件，建议直接使用更常见的中文字体：

```yaml
word-render:
  default-font-family: Microsoft YaHei
  fonts:
    default-family: Microsoft YaHei, PingFang SC, Noto Sans SC, SimSun, Arial Unicode MS, sans-serif
```

说明：

- 未配置 `regular-path` / `bold-path` 时，会直接使用 `default-font-family` 或 `default-family` 的首选字体名
- 配置了 `regular-path` / `bold-path` 时，组件会实际读取字体文件并校验可用性，优先使用解析出来的字体家族名
- `word-render.colors.*` 用于统一标题体系配色，默认已收敛为正式风格的深色标题
- `word-render.watermark.*` 用于开启全局文字水印，单次调用也可以通过 `WordRenderOptions.watermarkText(...)` 覆盖
- `word-render.watermark.color` 和 `word-render.watermark.font-size` 可控制水印颜色与字号
- 当前版本不会把字体文件嵌入 `docx`，最终显示效果仍取决于打开文档的环境是否安装对应字体

## 最简 Java 调用

生产上如果文档较大或并发较高，优先推荐 `renderDocxToStream(...)` 或 `renderDocxToFile(...)`。  
`renderDocx(...) -> byte[]` 仍然保留，但更适合下载接口、小文档或必须拿到内存字节数组的场景。

```java
WordRenderService service = applicationContext.getBean(WordRenderService.class);

byte[] docx = service.renderDocx(
    "# 经营分析报告\n\n这是 Dify workflow 输出的 Markdown 报告。",
    WordRenderOptions.builder()
        .contentType(WordRenderContentType.MARKDOWN)
        .reportStyle(WordRenderReportStyle.FORMAL)
        .title("经营分析报告")
        .headerText("经营分析")
        .headerPosition(WordRenderPosition.RIGHT)
        .footerText("仅供内部使用")
        .showPageNumber(true)
        .headingColor("#1F1F1F")
        .watermarkText("内部资料")
        .author("Dify Workflow")
        .metadata("reportNo", "WR-20260320")
        .build()
);
```

## 详细使用示例

### 1. Markdown 直接生成 byte[]

适用于 Dify workflow 直接输出 Markdown，然后你要上传 OSS、返回下载流或存储二进制。  
如果是几十页报告、并发接近 20，优先改用下面的流式输出方式。

```java
String markdown = ""
    + "# 经营分析报告\n\n"
    + "## 核心结论\n\n"
    + "- 本周新增客户 **28 家**\n"
    + "- 商机转化率提升到 `18.7%`\n\n"
    + "> 本报告由 Dify 自动生成。\n\n"
    + "| 指标 | 本周 | 上周 | 环比 |\n"
    + "| --- | --- | --- | --- |\n"
    + "| 线索数 | 320 | 275 | +16.4% |\n"
    + "| 成交数 | 28 | 21 | +33.3% |\n";

byte[] bytes = wordRenderService.renderDocx(
    markdown,
    WordRenderOptions.builder()
        .contentType(WordRenderContentType.MARKDOWN)
        .reportStyle(WordRenderReportStyle.FORMAL)
        .title("经营分析报告")
        .subTitle("Dify Markdown 直出示例")
        .headerText("经营分析")
        .headerPosition(WordRenderPosition.RIGHT)
        .footerText("仅供内部使用")
        .showPageNumber(true)
        .watermarkText("内部资料")
        .watermarkColor("#BFBFBF")
        .watermarkFontSize(22)
        .author("dify-workflow")
        .metadata("reportNo", "WR-20260327-01")
        .build()
);
```

### 2. 普通文本直接生成 byte[]

```java
String plainText = ""
    + "第一段：这是报告背景。\n"
    + "第二行：说明本次分析范围。\n\n"
    + "第二段：这是关键结论。\n"
    + "第三段：这是后续建议。";

byte[] bytes = wordRenderService.renderDocx(
    plainText,
    WordRenderOptions.builder()
        .contentType(WordRenderContentType.PLAIN_TEXT)
        .reportStyle(WordRenderReportStyle.DEFAULT)
        .title("普通文本报告")
        .footerText("运营中心")
        .showPageNumber(true)
        .author("report-job")
        .build()
);
```

### 3. Markdown 直接落盘到文件

```java
Path output = Paths.get("/tmp/word-render/weekly-report.docx");

wordRenderService.renderDocxToFile(
    "# 周报总览\n\n## 核心结论\n\n- 本周新增客户 28 家\n- 转化率提升到 **18.7%**",
    WordRenderOptions.builder()
        .contentType(WordRenderContentType.MARKDOWN)
        .reportStyle(WordRenderReportStyle.FORMAL)
        .title("周报总览")
        .subTitle("自动生成文件示例")
        .headerText("销售周报")
        .headerPosition(WordRenderPosition.RIGHT)
        .footerText("销售中心")
        .showPageNumber(true)
        .watermarkText("内部资料")
        .watermarkColor("#BFBFBF")
        .watermarkFontSize(22)
        .author("dify-workflow")
        .metadata("department", "销售中心")
        .metadata("reportNo", "WR-20260327-02")
        .build(),
    output
);
```

### 3.1 生产推荐：直接写 OutputStream

适用于：

- Web 下载接口
- 上传对象存储
- 20 左右并发导出
- 减少内存中额外 `byte[]` 副本

```java
try (OutputStream outputStream = response.getOutputStream()) {
    wordRenderService.renderDocxToStream(
        markdown,
        WordRenderOptions.builder()
            .contentType(WordRenderContentType.MARKDOWN)
            .reportStyle(WordRenderReportStyle.FORMAL)
            .title("经营分析报告")
            .footerText("内部资料")
            .showPageNumber(true)
            .watermarkText("内部资料")
            .watermarkColor("#BFBFBF")
            .watermarkFontSize(22)
            .build(),
        outputStream
    );
}
```

### 4. 模板末尾追加模式

适用于前几页固定、动态内容统一追加到模板后面的场景。

```java
byte[] bytes = wordRenderService.renderDocx(
    "# 个性化分析\n\n## 动态结论\n\n- 这是每次生成都不同的正文",
    WordRenderOptions.builder()
        .contentType(WordRenderContentType.MARKDOWN)
        .templateResource("classpath:/templates/fixed-report-template.docx")
        .appendPageBreakAfterTemplate(true)
        .headerText("经营分析")
        .headerPosition(WordRenderPosition.RIGHT)
        .footerText("仅供内部使用")
        .showPageNumber(true)
        .watermarkText("内部资料")
        .watermarkColor("#BFBFBF")
        .watermarkFontSize(22)
        .build()
);
```

### 5. 模板占位填充模式

适用于前几页固定、中间几页动态、后几页固定附录或声明页。

占位符不是只支持 `${wr_summary}`、`${wr_content}`、`${wr_appendix}` 这 3 个名字。
这 3 个只是示例命名，实际支持任意占位名，只要模板里和 `templateBinding(...)` 里保持一致即可。

模板中的占位段落示例：

```text
${wr_summary}
${wr_content}
${wr_appendix}
```

调用示例：

```java
byte[] bytes = wordRenderService.renderDocx(
    "",
    WordRenderOptions.builder()
        .templateResource("classpath:/templates/fixed-sections-template.docx")
        .templateMode(WordRenderTemplateMode.PLACEHOLDER)
        .templateBinding("wr_summary", WordRenderTemplateBinding.builder()
            .content("这是自动生成的摘要，适合填充固定模板中的摘要区域。")
            .contentType(WordRenderContentType.PLAIN_TEXT)
            .build())
        .templateBinding("wr_content", WordRenderTemplateBinding.builder()
            .content("# 动态正文标题\n\n## 子章节\n\n- 列表项一\n- 列表项二\n\n| 指标 | 数值 |\n| --- | --- |\n| 转化率 | 18.7% |")
            .contentType(WordRenderContentType.MARKDOWN)
            .baseHeadingLevel(3)
            .build())
        .templateBinding("wr_appendix", WordRenderTemplateBinding.builder()
            .content("附录说明：以上内容仅用于内部分析。")
            .contentType(WordRenderContentType.PLAIN_TEXT)
            .build())
        .footerText("内部资料")
        .showPageNumber(true)
        .build()
);
```

### 5.1 多动态内容块模板示例

适用于这种结构：

- 固定 1-N 页
- 动态内容块 A
- 固定 M+5 页
- 动态内容块 B
- 固定尾页（可选）

模板示例：

```text
固定封面页
固定目录页
${wr_section_a}
固定中间说明页
${wr_section_b}
固定附录页
```

调用示例：

```java
byte[] bytes = wordRenderService.renderDocx(
    "",
    WordRenderOptions.builder()
        .templateResource("classpath:/templates/multi-dynamic-sections.docx")
        .templateMode(WordRenderTemplateMode.PLACEHOLDER)
        .templateBinding("wr_section_a", WordRenderTemplateBinding.builder()
            .content("# 第一部分标题\n\n## 第一部分小节\n\n- 动态内容 A1\n- 动态内容 A2")
            .contentType(WordRenderContentType.MARKDOWN)
            .baseHeadingLevel(3)
            .build())
        .templateBinding("wr_section_b", WordRenderTemplateBinding.builder()
            .content("第二部分是普通文本内容，用于插入固定说明页之后的动态区域。")
            .contentType(WordRenderContentType.PLAIN_TEXT)
            .build())
        .footerText("内部资料")
        .showPageNumber(true)
        .build()
);
```

这个模式下：

- 模板里的固定页内容会原样保留
- `${wr_section_a}` 会替换成第一段动态内容
- `${wr_section_b}` 会替换成第二段动态内容
- 固定尾页如果存在，也会继续保留

### 6. 标题偏移示例

当模板里已经有一级、二级标题时，可以让动态正文的最高标题从三级开始：

```java
WordRenderTemplateBinding contentBinding = WordRenderTemplateBinding.builder()
    .content("# 动态一级标题\n\n## 动态二级标题")
    .contentType(WordRenderContentType.MARKDOWN)
    .baseHeadingLevel(3)
    .build();
```

效果：

- Markdown `#` -> Word 三级标题
- Markdown `##` -> Word 四级标题语义

### 7. 页眉、页脚、页码示例

```java
WordRenderOptions options = WordRenderOptions.builder()
    .contentType(WordRenderContentType.MARKDOWN)
    .title("正式报告")
    .headerText("经营分析")
    .headerPosition(WordRenderPosition.RIGHT)
    .footerText("仅供内部使用")
    .showPageNumber(true)
    .build();
```

当前默认布局：

- 页眉：按 `headerPosition` 放置
- 页脚文字：左下
- 页码：右下，显示为 `1/N`

### 8. 标题颜色与正式风格示例

```java
WordRenderOptions options = WordRenderOptions.builder()
    .contentType(WordRenderContentType.MARKDOWN)
    .title("颜色定制报告")
    .titleColor("#1F1F1F")
    .headingColor("#1F1F1F")
    .accentColor("#7F7F7F")
    .build();
```

注意：

- `titleColor` 用于封面主标题
- `headingColor` 用于正文标题体系
- `accentColor` 用于副标题和辅助元素
- 颜色值必须是 6 位十六进制

### 9. 文字水印示例

```java
WordRenderOptions options = WordRenderOptions.builder()
    .contentType(WordRenderContentType.MARKDOWN)
    .title("带水印报告")
    .watermarkText("内部资料")
    .watermarkColor("#BFBFBF")
    .watermarkFontSize(22)
    .build();
```

当前水印风格：

- 浅灰色
- 低透明度
- 斜向放置
- 尺寸较小，尽量不遮挡正文

## Spring Boot 业务服务示例

```java
@Service
public class ReportApplicationService {

    private final WordRenderService wordRenderService;

    public ReportApplicationService(WordRenderService wordRenderService) {
        this.wordRenderService = wordRenderService;
    }

    public byte[] renderDifyMarkdown(String markdown) {
        return wordRenderService.renderDocx(markdown, WordRenderOptions.builder()
            .contentType(WordRenderContentType.MARKDOWN)
            .reportStyle(WordRenderReportStyle.FORMAL)
            .title("Dify 报告")
            .subTitle("自动生成的业务分析")
            .headerText("Dify Workflow")
            .headerPosition(WordRenderPosition.RIGHT)
            .footerText("内部资料")
            .showPageNumber(true)
            .watermarkText("内部资料")
            .author("workflow-engine")
            .build());
    }
}
```

## 直接生成文件示例

```java
Path output = Paths.get("/tmp/word-render/weekly-report.docx");

wordRenderService.renderDocxToFile(
    "# 周报总览\n\n## 核心结论\n\n- 本周新增客户 28 家\n- 转化率提升到 **18.7%**",
    WordRenderOptions.builder()
        .contentType(WordRenderContentType.MARKDOWN)
        .reportStyle(WordRenderReportStyle.FORMAL)
        .title("周报总览")
        .subTitle("自动生成文件示例")
        .headerText("销售周报")
        .headerPosition(WordRenderPosition.RIGHT)
        .footerText("销售中心")
        .showPageNumber(true)
        .author("dify-workflow")
        .metadata("department", "销售中心")
        .metadata("reportNo", "WR-20260320-01")
        .build(),
    output
);
```

普通文本落盘示例：

```java
Path output = Paths.get("/tmp/word-render/plain-report.docx");

wordRenderService.renderDocxToFile(
    "这是一个普通文本报告。\n\n第一段用于描述背景。\n第二段用于描述结论。",
    WordRenderOptions.builder()
        .contentType(WordRenderContentType.PLAIN_TEXT)
        .reportStyle(WordRenderReportStyle.DEFAULT)
        .title("普通文本报告")
        .author("report-job")
        .build(),
    output
);
```

## 模板追加示例

如果你的报告前几页是固定内容，可以准备一个现成的 `docx` 模板，然后把动态正文追加到模板后面：

```java
byte[] docx = wordRenderService.renderDocx(
    "# 个性化分析\n\n这里开始是每次调用都不同的正文内容。",
    WordRenderOptions.builder()
        .contentType(WordRenderContentType.MARKDOWN)
        .templateResource("classpath:/templates/fixed-report-template.docx")
        .appendPageBreakAfterTemplate(true)
        .headerText("经营分析")
        .headerPosition(WordRenderPosition.RIGHT)
        .footerText("仅供内部使用")
        .showPageNumber(true)
        .build()
);
```

说明：

- `templateResource` 支持 `classpath:` 和 `file:` 形式
- 传了模板后，会先加载模板文档，再把动态内容追加到模板后面
- 当前模板模式不会再额外生成默认封面，避免打乱模板里的固定前置页
- 不传 `templateResource` 时，仍然走现在的直接生成链路

## 模板占位填充示例

如果你的报告是“固定封面 + 中间动态内容 + 固定附录”，更适合用占位填充模式：

模板中放置占位段落：

```text
${wr_summary}
${wr_content}
${wr_appendix}
```

调用示例：

```java
byte[] docx = wordRenderService.renderDocx(
    "",
    WordRenderOptions.builder()
        .templateResource("classpath:/templates/fixed-sections-template.docx")
        .templateMode(WordRenderTemplateMode.PLACEHOLDER)
        .templateBinding("wr_summary", WordRenderTemplateBinding.builder()
            .content("这是摘要内容。")
            .contentType(WordRenderContentType.PLAIN_TEXT)
            .build())
        .templateBinding("wr_content", WordRenderTemplateBinding.builder()
            .content("# 动态正文标题\n\n- 列表项一\n- 列表项二")
            .contentType(WordRenderContentType.MARKDOWN)
            .baseHeadingLevel(3)
            .build())
        .templateBinding("wr_appendix", WordRenderTemplateBinding.builder()
            .content("附录说明")
            .contentType(WordRenderContentType.PLAIN_TEXT)
            .build())
        .build()
);
```

说明：

- `templateMode=PLACEHOLDER` 时，会在命中的占位段落位置插入渲染后的 Word 块内容
- 每个占位可单独声明 `MARKDOWN` 或 `PLAIN_TEXT`
- `baseHeadingLevel=3` 时，Markdown 中的 `#` 会按 Word 三级标题输出
- 若调用时绑定了模板中不存在的占位符，会抛出明确异常
- 模板中存在但未绑定的占位符会原样保留
- v1 要求占位符独占一个段落，例如 `${wr_content}`，不要拆散到多个 run

## Dify Markdown 示例

```markdown
# 周报总览

## 核心结论

- 本周新增客户 28 家
- 线索转化率提升到 **18.7%**
- 华东区域表现最好

## 风险提醒

> 北方区域签约周期拉长，需要重点关注。

## 数据表

| 指标 | 数值 |
| --- | --- |
| 新增线索 | 182 |
| 成交客户 | 34 |
| 转化率 | 18.7% |
```

## 普通文本示例

```text
这是一个普通文本报告。

第一段内容用于描述背景。
第二段内容用于描述当前结论。

最后一段用于补充建议项。
```

## 输出风格说明

- `DEFAULT`: 通用正式风格，适合大多数内部报告
- `FORMAL`: 更正式的页边距和版式，适合经营分析、汇报材料
- `SIMPLE`: 极简风格，适合系统导出或轻量总结

## 文档附加元素

- `title` 只用于封面和文档元数据，不会自动变成页眉
- `headerText` 传入后才会生成页眉，位置由 `headerPosition` 控制
- `footerText` 传入后才会生成左下角页脚文字
- `showPageNumber=true` 后才会生成右下角页码，格式为 `1/N`
- 页眉、页脚、页码都会以较小字号和弱化灰色显示，避免抢正文

## 场景说明

- 场景 1：Dify workflow 返回 Markdown，直接生成 `byte[]` 上传对象存储
- 场景 2：普通文本使用默认风格生成 `docx`
- 场景 2.1：Markdown 或普通文本直接落盘生成本地 `docx` 文件
- 场景 3：通过 `metadata` 附加报告编号、部门、业务线
- 场景 4：通过模板保留固定前置页，后续正文按 Markdown 或普通文本动态追加
- 场景 5：通过模板占位填充固定封面、中间动态正文和固定附录
- 场景 6：模板中已定义一二级标题，动态正文通过 `baseHeadingLevel` 从三级开始
- 场景 7：通过 `application.yaml` 切换默认风格
- 场景 8：高频调用场景启用 Redis 缓存，业务代码无需改动

## 测试输出样例

执行 `mvn -Dmaven.repo.local=.m2/repository test` 后，会在 `target/word-render-test-output/` 下生成可直接打开验收的样例文档。

- `plain-text-report.docx`: 轻量普通文本样例
- `markdown-report.docx`: 轻量 Markdown 样例
- `complex-markdown-report.docx`: 复杂 Markdown 样例，包含多级标题、引用、代码块、列表和表格
- `complex-template-placeholder-report.docx`: 复杂模板占位样例，包含固定前后页、摘要区、三级起始正文和表格
- `watermark-report.docx`: 水印样例，包含自定义标题颜色和弱化文字水印

这些文件是测试阶段生成的验收样例，不属于业务接口的一部分。
