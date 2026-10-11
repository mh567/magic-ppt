# Magic PPT

Magic PPT 是一个将原始素材制作成完整 PowerPoint 演示文稿的 Agent Skill。它覆盖内容整理、页面规划、模板复用、图像生成、口播备注、PPTX 组装和最终渲染检查。

当前版本：`0.1.0`。版本以 [VERSION](VERSION) 为准，变化见 [CHANGELOG.md](CHANGELOG.md)。

## 核心能力

- 读取文档、Markdown、PDF、表格和图片等素材，整理页面内容并保留来源。
- 复用已绑定模板的母版、背景及固定元素。
- 生成独立完整的逐页提示词与页面图像，检查文字、数字、布局及裁切。
- 将口播写入每页备注，组装成 PPTX，并检查最终渲染页面。
- 保存计划、图像及验收记录，支持中断恢复和局部修正。

| 制作方式 | 结果与边界 |
|---|---|
| 模板复用，默认 | 固定元素保持原生，正文采用图像；正文图内文字无法逐字编辑 |
| 整页图像 | 每页一张图，备注可编辑；模板外观为近似呈现 |
| 原生可编辑正文 | 使用宿主演示文稿工具构建文字与图形；随附图像流水线脚本不直接支持此模式 |

## 运行环境

ChatGPT/Codex 路径优先使用当前内置图像工具。其他 Agent 可使用已配置的生成与编辑工具，按实际能力适配，并记录真实工具来源。不同宿主须具备文件读写、必要的脚本执行与图像查看能力；纯聊天环境不保证完成全部流程。

脚本依赖 Python 3、`python-pptx`、`Pillow` 和 `lxml`。额外素材抽取依赖 `python-docx`、`openpyxl`、`pypdf`。最终渲染可使用宿主渲染器，或已安装的 LibreOffice 与 Poppler；系统工具需单独安装。

## 安装

### Codex

先确认当前 GitHub 账户具有仓库读取权限，再克隆到技能目录。目标目录已存在时按更新方法操作。

```bash
git clone https://github.com/mh567/magic-ppt.git "${CODEX_HOME:-$HOME/.codex}/skills/magic-ppt"
```

如需使用随附脚本，可在技能目录创建独立 Python 环境：

```bash
cd "${CODEX_HOME:-$HOME/.codex}/skills/magic-ppt"
python3 -m venv .venv
.venv/bin/python -m pip install python-pptx Pillow lxml python-docx openpyxl pypdf
.venv/bin/python scripts/deck_pipeline.py --version
.venv/bin/python scripts/self_test.py
```

Windows 可使用 `python` 创建环境，并将 `.venv/bin/python` 替换为 `.venv\Scripts\python.exe`。

安装后开启新会话，使用 `$magic-ppt` 调用。

### 其他 Agent

将仓库克隆到该宿主文档指定的技能目录，保留整个 `magic-ppt` 文件夹。宿主支持 `SKILL.md` 时按其方式加载；不支持自动发现时，显式提供技能入口。依赖安装和图像工具配置遵循宿主方式。

模板、提示词和验收规则随技能一起提供；不同工具的参数按当前签名适配，不保证每个宿主具备全部能力。模板绑定与视觉状态见 [模板绑定说明](references/template-binding.md)。

## 使用

```text
使用 $magic-ppt，把这些素材制作成沿用已绑定模板的完整 PPTX，附逐页口播备注。
```

```text
使用 $magic-ppt，把这些素材做成原生可编辑的演示文稿。
```

默认交付 `final.pptx`、含口播的 `prompts.md`、最终渲染联系表 `contact-sheet.jpg` 和 `qa.json`。原生模式按实际工具输出，口播写入备注，图像提示词仅在实际生成时提供。

## 更新

在安装目录先检查本地修改，再拉取远程更新：

```bash
cd "${CODEX_HOME:-$HOME/.codex}/skills/magic-ppt"
git status --short
git pull --ff-only
.venv/bin/python scripts/deck_pipeline.py --version
.venv/bin/python scripts/self_test.py
```

存在本地修改时先保存或提交，再更新；不要直接覆盖自定义模板。依赖变化见变更记录。更新后开启新会话加载新版。

通过文件复制安装的版本，可先备份原目录，再从仓库下载完整新版本并替换，保留所需自定义资产并重新检查模板绑定。

## 版本与验证

版本采用语义化版本规则，`VERSION` 与 `SKILL.md` 中的版本保持一致。未发布变化记入 `CHANGELOG.md` 的 Unreleased，发布时归档为对应版本。

自测验证脚本行为与合成夹具，不代表真实出图质量或目标应用验收。每份成品仍需渲染最终 PPTX 并逐页检查；字体替换、工具缺口及未完成检查会在交付时说明。
