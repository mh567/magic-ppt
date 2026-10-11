# 计划格式与命令

所有 JSON 为 UTF-8。相对路径按计划文件所在目录解析；脚本绝不依赖调用时工作目录。

## 顶层字段

| 字段 | 说明 |
|---|---|
| schema_version | 固定 1 |
| title | 整套课件名称 |
| mode | locked-template（默认）或 full-page-image（明确选择） |
| template_profile | 可选，真实绑定 profile 的路径；省略用 skill 的 assets/master/profile.json |
| requested_image_model | ChatGPT 路径默认 ChatGPT current built-in image tool；其他路径默认 Current configured host image tool；明确型号时填写真实标识 |
| require_exact_model | 默认 false；后端未暴露不阻碍执行；仅新任务明确要求型号时才设 true |
| image_provider | 默认 chatgpt-built-in，记录 tool 为 image_gen.imagegen 或 image_gen__imagegen；其他 Agent 用 host-tool；合成测试用 test-fixture |
| image_tool | host-tool 必填，指定实际已配置工具标识；record 的 tool 必须与其一致 |
| min_image_width | 默认 1920；用户接受低分辨率时才降低，不用放大伪称高清 |
| max_attempts_per_slide | 默认 3；记录每次真实生成/修正，不自动无限重试 |
| slides | 有序页面数组 |

## 页面字段

每页包含 `id: P001/P002…`、`order: 1/2…`、`kind: cover/body/section/closing`、`title`、`subtitle`（可选）、`on_slide`（精确上屏字符串数组）、`layout`、`speaker_notes`、`sources`（来源字符串数组）。封面 `footer` 可用真正换行，正文页码由 order 生成。

正文图片涵盖 `on_slide` 的全部内容。默认封面/章节/结束页的上屏内容只能通过绑定角色替换；真实章节模板三行 `navigation_1/2/3` 分别对应三个 on_slide 项；必须恰好三项。不存在对应文字角色时，不往原模板额外添加文本框；先调整绑定或把详细内容拆成 body 页。

已绑定模板：cover 的主副标题换行合并到原框；section 为原三行导航，需三个 on_slide 项；其他导航结构与讨论页使用 body；closing 用于结束页。正文页码按 order 显示，不带前导 0。

布局文字必须限定信息顺序、分组和连接关系，不只说“专业、高端、美观”。上屏文本不要藏在 layout 里，避免漏字验收。

`generation` 与 `qa` 由脚本写入，不预填 passed。

## 运行顺序

```bash
python scripts/deck_pipeline.py prepare --plan 工作目录/deck.json --out 工作目录/build
# 按 build/jobs.json 使用宿主工具逐页生成；记录真实输出路径。
python scripts/deck_pipeline.py record --plan 工作目录/deck.json --slide P001 \
  --image 图片路径 --actual-model backend-not-exposed --tool image_gen.imagegen
# 在看过该页并完成全部检查后登记。
python scripts/deck_pipeline.py approve --plan 工作目录/deck.json --slide P001 \
  --reviewer agent --checks text,numbers,diagram,master,crop --comment '说明已核对的实际结果'
# 每页登记与验收后组装。
python scripts/deck_pipeline.py assemble --plan 工作目录/deck.json --out 工作目录/final.pptx
python scripts/deck_pipeline.py verify --plan 工作目录/deck.json \
  --pptx 工作目录/final.pptx --out 工作目录/qa.json
python scripts/deck_pipeline.py render --pptx 工作目录/final.pptx --out 工作目录/render-v1
python scripts/deck_pipeline.py contact-sheet --images 工作目录/render-v1 \
  --out 工作目录/contact-sheet.jpg
# 逐页看最终渲染图；只看拼图不足以核对小字。
python scripts/deck_pipeline.py final-review --plan 工作目录/deck.json \
  --pptx 工作目录/final.pptx --renders 工作目录/render-v1 \
  --reviewer agent --comment '说明最终逐页检查的结果' --out 工作目录/qa.json
```

宿主渲染器可以替代 `render`。若替代，保留真实 renderer、PPTX sha256、页数和每页 PNG sha256，写出同结构 `render.json`，不能手工假称执行过渲染。

## 恢复与失效

运行 prepare 生成新 jobs；`reuse` 仅表示该页图像和当前内容/模板哈希一致、且已逐项登记验收。修改文字、layout、页序、模板或型号严格要求，旧记录失效。`record` 新图后旧 QA 重置。验证拒绝重复图片、未验收页、低分辨率、非 16:9、变更后的图或内容和不满足指定型号的输出。

备注完整性与固定对象指纹属于自动检查；错别字、箭头方向、视觉遮挡属于人工/智能体目视检查，不能仅靠 JSON 状态证明正确。
