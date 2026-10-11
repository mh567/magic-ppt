# Magic PPT

Magic PPT 是一个将原始素材制作成完整 PowerPoint 演示文稿的 Agent Skill。默认交付原生可编辑 PPTX，覆盖内容整理、页面规划、模板复用、口播备注和最终渲染检查；图片可用于效果预览与确认。

当前版本：`0.1.0`。版本以 [VERSION](VERSION) 为准，变化见 [CHANGELOG.md](CHANGELOG.md)。

## 核心能力

- 读取文档、Markdown、PDF、表格和图片等素材，整理页面内容并保留来源。
- 复用已绑定模板的母版、背景及固定元素。
- 制作原生文字、图形与数据对象，检查准确性、可读性与编辑范围。
- 将口播写入每页备注，组装成 PPTX，并检查最终渲染页面。
- 保存计划、图像及验收记录，支持中断恢复和局部修正。

| 制作方式 | 结果与边界 |
|---|---|
| 原生可编辑，默认 | 文字与结构对象可编辑，照片与插画可为图片；预览可使用渲染图片 |
| 模板复用加正文图片 | 用户明确要求时使用，固定区原生，正文图内文字不可逐字编辑 |
| 整页图像 | 用户明确要求时使用，每页一张图，备注可编辑；模板外观近似 |

## 运行环境

ChatGPT/Codex 路径优先使用当前内置图像工具。其他 Agent 可使用已配置的生成与编辑工具，按实际能力适配，并记录真实工具来源。不同宿主须具备文件读写、必要的脚本执行与图像查看能力；纯聊天环境不保证完成全部流程。

脚本依赖 Python 3、`python-pptx`、`Pillow` 和 `lxml`。额外素材抽取依赖 `python-docx`、`openpyxl`、`pypdf`。最终渲染可使用宿主渲染器，或已安装的 LibreOffice 与 Poppler；系统工具需单独安装。

## 安装

### 在线安装

直接对 Agent 说：

```text
请安装 https://github.com/mh567/magic-ppt 这个 skill，并自动检查运行环境，补齐缺失的 Python 依赖。
```

Agent 会将完整技能安装到当前宿主的 skills 目录，检查依赖与可用工具，并说明检查结果。私有仓库需当前账户具有读取权限。

### 离线安装

从 [Releases](https://github.com/mh567/magic-ppt/releases) 下载 `magic-ppt-<版本号>.zip`，解压到当前 Agent 的 skills 目录即可。目录结构应为：

```text
skills/
  magic-ppt/
    SKILL.md
    scripts/
    references/
    assets/
```

ZIP 需要解压，直接放置压缩文件无法加载技能。开启新会话后使用 `$magic-ppt`，Agent 会自动检查环境。离线包包含技能与模板，不包含 Python 运行时和依赖；完全断网环境需预先具备依赖，缺项会明确报告。

## 自动环境检查

安装完成或首次使用时，Agent 运行 `scripts/setup_check.py`，优先复用已有环境。在线且已授权补齐依赖时使用 `--install`，仅在技能目录的 `.venv` 中安装缺失包。

本地脚本检查 Python 依赖与本地渲染工具；Agent 另行确认宿主图像工具、可用渲染器与模板字体。缺少本地 LibreOffice 时可使用宿主渲染器。检查 Python 依赖通过不代表图像生成或成品验收已完成。

## 使用

```text
使用 $magic-ppt，把这些素材制作成沿用已绑定模板的完整 PPTX，附逐页口播备注。
```

```text
使用 $magic-ppt，把这些素材做成原生可编辑的演示文稿。
```

默认交付原生可编辑的 `final.pptx`、最终渲染联系表 `contact-sheet.jpg` 和 `qa.json`。口播写入备注；仅实际生成提示词时提供 `prompts.md`。可要求先展示代表页效果图并确认后继续制作。

## 更新

在线安装后，直接对 Agent 说：

```text
请更新 magic-ppt 到最新版本，保留我的自定义模板，并重新检查环境。
```

Agent 根据安装方式拉取或替换新版，先保存本地改动和自定义资产，核对版本与模板绑定。离线安装时下载新的版本 ZIP，备份旧目录后解压替换。更新后开启新会话。

## 版本与验证

版本采用语义化版本规则，`VERSION` 与 `SKILL.md` 中的版本保持一致。未发布变化记入 `CHANGELOG.md` 的 Unreleased，发布时归档为对应版本。

自测验证脚本行为与合成夹具，不代表真实出图质量或目标应用验收。每份成品仍需渲染最终 PPTX 并逐页检查；字体替换、工具缺口及未完成检查会在交付时说明。

## 发布流程

1. 将本次变化从 Unreleased 归档到新版本，更新 VERSION 和 SKILL.md 的 metadata.version。
2. 提交并推送 main，等待 Validate and release 工作流通过。
3. 在同一已验证提交创建标签 `v<版本号>` 并推送标签，工作流自动构建并发布离线 ZIP、SHA256SUMS.txt 与 release-manifest.json。

本地可执行 `python scripts/release.py --out ../dist` 验证与打包。要求已提交且干净的工作区；标签必须与版本一致。现有 Release 不覆盖，失败修复后重新检查；已发布版本的功能修复使用新版本号。

发布包使用固定 ZIP 元数据，并从 Git 提交读取文件，不包含 Git 元数据、CI 配置或本地环境。项目版本与提交身份写入发布清单。
