# 一次绑定真实模板

只在尚未绑定或用户明确更换模板时执行。目标是把真实 `PPT模板.pptx` 的原生母版资源永久放进 skill，此后调用仅需素材。不能用描述重建的 PPT 冒充原模板。

## 检查与选择页面

1. 从用户已授权文件取得原始 PPTX 的真实字节，保留原件，只处理副本。读取文件技能，渲染所有模板页；核对尺寸、字体、母版、背景、Logo、水印、页码签和标题层级。
2. 运行：

```bash
python scripts/template_tools.py inspect --pptx 原始模板.pptx --out 工作目录/template-inspection.json
```

3. 把每类母版映射到实际模板页。`slide_index` 是从 1 开始的页面位置；`shape_id` 来自检查结果，**不是数组位置**。
4. 以 `assets/bind-config-example.json` 的结构写 `bind-config.json`，填写真实值。cover 与 body 必填，section/closing 按真实模板补充；不能假设第 1/2/3/6 页必然分别对应四类。

`role_shapes` 指向原页的顶层原生文字框：title、subtitle、page_number、footer、navigation、navigation_1/2/3、closing_text。替换时保留原段落与首个文字 run 的格式。

`remove_shape_ids` 只列出样例正文需要清除的对象；Logo、弧线、水印、蓝黄标识、标题背景条、分隔线、页码签底形和任何其他固定对象不删除。固定品牌文字也不随意清空。

正文 `content_box` 为归一化 `[x,y,w,h]`，矩形必须在标题/副标题下方、页码签上方。脚本检查越界与角色文字交叠；还要目视检查水印和其他固定装饰。整个矩形会由正文图片覆盖，因此若装饰位于正文区内，调整矩形、分区放图或用透明图，不把“源对象仍在文件里”当成“视觉未遮挡”。本模板已配置 requires_transparent_content 和 transparent_boxes；所选图像工具生成透明正文，脚本检查真实 alpha 与水印区透明度，避免覆盖原水印。

遇到组合对象内的标题：检查后在**模板副本**中拆分并核对外观，或扩展脚本支持组合对象；不要猜 ID。绑定脚本明确拒绝不能直接定位的角色。

遇到一段多色/多字号标题：首 run 保留无法重现所有样式分段，需手动保持 run 映射并修改替换函数；绑定时说明并验证。普通单样式标题保持源 XML 格式。

正文页码可设置 page_number_mode: inherited-field 和 layout_page_number_shape_id，指向原版式真正的 slidenum 字段。保留源字段，按最终页序自动显示页码，不新增重复号码框。封面 combine_title_subtitle 将主副标题换行填入原框；三行导航使用 navigation_1/2/3 与 navigation_count: 3。

## 绑定、渲染与验收

```bash
python scripts/template_tools.py bind --pptx 原始模板.pptx --config 工作目录/bind-config.json
python scripts/template_tools.py render --profile assets/master/profile.json
```

`render` 需要 `soffice` 与 `pdftoppm` 在 PATH；也可使用宿主渲染器生成同路径参考。源 PPT 从不修改。`blank-preview.pptx` 只用于展示无动态文字/样例正文的固定底板。

按 profile 的 visual_status、render_warnings 和实际参考图判断是否需要补验收。模板已绑定时复用原件；字体或渲染异常先在可用环境补检。采用外部宿主渲染器时，把真实参考写入相同路径并记录 renderer、字体替代和验收证据，再根据真实结果清除 render_warnings。不得仅删除警告绕过验收。

逐类目视比对 `references/<kind>.png` 与 `blank/<kind>.png`：只应去掉规定的可替换文字和样例正文；背景、标识、水印、边界与固定装饰必须保留。确认配置中的正文矩形不会遮挡必须可见的固定元素，再标记：

```bash
python scripts/template_tools.py mark-verified --profile assets/master/profile.json \
  --reviewer agent --comment '已逐类核对原页与空白底板，确认删除范围、角色文字与正文安全区'
```

真正运行完视觉检查后才能写该注释，不能用命令代替检查。

## 槽位索引

绑定后可生成模板页的可寻址文字槽清单，供按模板页复用原文排版时使用：

```bash
python scripts/template_tools.py index --pptx assets/master/source.pptx --out assets/master/slots.json
```

每个槽位记录稳定 id（`s{页}_sh{形状}_p{段落}`）、原始文本、`{shape_id, paragraph}` 地址、容量预算（`chars_per_line`、`max_lines`、`budget`）、`autofit` 标记和字号层级 `level`。文件级另有来源名、SHA256、画布尺寸与 `type_scale`。组合形状内的文字一并索引；退化文本框标 `capacity_unknown` 并跳过容量校验。

字号按 OOXML 继承链解析：run 显式值、段落值、版式占位符 `lstStyle`、母版占位符 `lstStyle`、母版 `txStyles`（按标题、正文、其它分桶），都没有才用 18pt 兜底。占位符标题常靠 `txStyles` 取到 44pt，只读 run 会低估到 18pt，容量预算随之偏大。普通文本框不参与继承，没有显式字号时按兜底值记录。

槽位按段落切分，同一段落内的局部着色或多字号作为格式保留，不拆成多个槽。脚本不判断文字语义，`editable` 默认全为 true，装饰序号、导航数字等锁定项由 Agent 在索引后按页补写。索引不替代绑定的视觉验收，也不改变母版固定区。

## 按槽位选页换字

索引之后可写 `plan.json` 组装成品，只换文字、不动排版：

```bash
python scripts/template_tools.py build --plan plan.json --slots assets/master/slots.json \
  --template assets/master/source.pptx --out final.pptx
```

```json
{
  "schema_version": 1,
  "slides": [
    {"source_slide": 3, "edits": [{"slot": "s3_sh46_p0", "text": "资料查全"}]},
    {"source_slide": 1, "edits": [{"slot": "s1_sh2_p0", "text": "第二章"}, {"slot": "s1_sh2_p1", "text": "复用的扉页"}]}
  ]
}
```

`source_slide` 是模板页序，可重复引用同一页，输出按数组顺序排列；未列出的模板页会被丢弃。替换只改段落文字，保留段落属性与首个 run 的字号、颜色、粗体；段落内的软换行与字段一并清除，不留残字。

以下情况直接报错，不产出半成品：槽位索引与模板 SHA256 不符、槽位 id 不在索引中、`source_slide` 越界、可编辑槽位未覆盖、文本仍是占位内容。文本结尾为省略号或超出容量预算只告警不阻断，避免为凑长度而截断。

## 验收

```bash
python scripts/template_tools.py accept --plan plan.json --slots assets/master/slots.json \
  --template assets/master/source.pptx --pptx final.pptx --out work --renders work/renders
```

`accept` 先做结构校验，再核对渲染证据，写出 `work/qa.json` 与 `work/contact-sheet.jpg`。

结构校验比较成品与模板页的固定对象指纹、母版与主题，并逐槽核对计划文字，因此外部改动固定元素、母版或槽位文字都会被拒。渲染证据必须属于同一个 PPTX 版本，页数与图像哈希需与文件夹一致。未提供 `--reviewer` 与 `--comment` 时 `visual_review` 保持 `pending`，两者任一为空即报错。宿主无本地渲染器时传 `--renders` 指向宿主渲染结果，并保留真实 renderer 名称，不用合成图片冒充验收。

保存 `assets/master/source.pptx`、`profile.json`、原页与空白参考，更新 `assets/template-status.json`；修改 skill 后验证并保存到用户指定位置；提交与推送遵循当前用户授权。这样一次绑定才在后续会话有效。更换模板保留旧版本；`--replace` 只在用户确实要求替换时用。

## 如何理解“完全一致”

原生固定对象与母版关系可以直接保留，并在组装后比较指纹；图像模型重画的背景不能保证完全一致。默认严格路径把生成图片限定到正文区域，避免用近似生成底板覆盖真实模板。字体缺失、渲染器差异、正文图片遮挡都会改变可见结果，因此仍需最终渲染核对。
