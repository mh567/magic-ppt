# 原生制作与效果预览

默认制作路径。先核对模板 profile、源文件和中文渲染验收，再用 `native_pipeline.py build` 试制。保留模板真实母版、版式、背景、固定对象和页码。按配置清除样例内容并替换标题等文字。模板缺失或待验收时读[模板绑定](template-binding.md)。

## 页面计划

从[原生示例计划](../assets/native-deck-example.json)开始。页序由数组顺序决定，ID 从 `P001` 连续编号。填写 `kind`、`title`、`speaker_notes`，按需填写 `subtitle`、`footer`、`sources`。正文的 `on_slide` 列出需核对的精确文字，以整段或独立显示行为单位。三行导航模板的章节页在 `on_slide` 填三条导航文字。脚本将来源以空行和「【来源】」标题附在口播后。

内置模板默认从技能目录加载。自定义 `template_profile` 和图片 `path` 相对于计划文件解析。沿用模板字体，遇到缺字时用 `font` 指定动态文字和正文的实际替代字体，并在 `font_substitution` 记录影响。保留固定对象的字体声明。

Agent 决定正文关系、构图和视觉层级。脚本读取 `objects` 并按数组顺序叠放对象。坐标单位为英寸，以整页左上角为原点，正文对象放在 profile 的正文安全区内。

| `type` | 字段 |
| --- | --- |
| `text` | `box: [x,y,w,h]`、`text` |
| `shape` | `box`、`shape`（如 `ROUNDED_RECTANGLE`）、`text`、`fill`、`stroke` |
| `connector` | `points: [x1,y1,x2,y2]`、`arrow: true`、`stroke`，可选 `line_type`（如 `ELBOW`） |
| `table` | `box`、二维 `rows`，可选 `column_widths`（总和为 1）、`fill`、`header`（`fill` 和 `style`） |
| `chart` | `box`、`categories`、`series`（每项含 `name` 和 `values`），可选 `title`、`legend`、`chart_type`（`COLUMN_CLUSTERED`、`BAR_CLUSTERED`、`LINE` 或 `LINE_MARKERS`） |
| `image` | `box`、`path`、可选 `alt`，用于照片或插画 |

颜色用六位 RGB 字符串。线条 `stroke` 支持 `color`、以磅计的 `width` 和 `dash`（如 `DASH`）。文字可用 `style` 指定 `font`、`size`、`bold`、`color`。多段文字可用 `paragraph_styles` 逐段覆盖。可选 `margins: [左,上,右,下]`、`align`（如 `CENTER`）、`anchor`（如 `MIDDLE`）、`wrap` 和以磅计的 `space_after`。

宿主工具可直接制作 PPTX，再用同一份计划验收。保留模板对象的原始标识、版式和资源。此时可省略构建用的 `objects`，用 `required_objects` 声明各类正文原生对象的最低数量，例如 `{"connector":2,"table":1}`。结构检查用计划核对备注、文字及对象类型和数量，连线语义、版面和文字尺寸由渲染复核确认。

## 试制、验收与交付

```bash
python scripts/native_pipeline.py build --plan 工作目录/slide-plan.json --out 工作目录/final.pptx
python scripts/native_pipeline.py accept --plan 工作目录/slide-plan.json --pptx 工作目录/final.pptx --out 工作目录
```

按[视觉设计](visual-design.md)选代表页，覆盖复杂关系或密集信息，使用同一入口试制。用户要求先确认效果时展示试页并等待反馈。试页合格后扩展为整套计划，再构建和渲染最终 PPTX。结构检查通过后，`accept` 用已配置的 LibreOffice 与 Poppler 渲染，生成联系表和待复核的 `qa.json`。

用宿主渲染时，将最终 PPTX 的逐页 PNG 与 `render.json` 放在同一目录，通过 `--renders` 指定。见[现有渲染命令](manifest.md)的文件约定。记录实际 `renderer`、PPTX 的 `pptx_sha256`、页数 `pages`、按文件名索引的 `image_sha256`，确保与当前 PPTX 和渲染页一一对应。PNG 按页序编号。

逐页查看最终渲染，核对内容、数字、关系、中文可读性、遮挡、裁切、固定区和页码，再用联系表核对整套层级。完成后重复执行 `accept`，加上 `--renders 渲染目录 --reviewer 复核人 --comment "本次检查结论"`。改动成品后重新生成渲染，复核受影响的页面，并重新记录结果。

交付 `final.pptx`、`contact-sheet.jpg` 和 `qa.json`。保存页面计划与资产，说明编辑范围、字体替换和实机验证状态。口播已写入 PPTX 备注，可另存 Markdown。仅在实际生成图像提示词时提供 `prompts.md`。
