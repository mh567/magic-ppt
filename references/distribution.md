# 增量分发

发布包同时产出整包 ZIP 与增量镜像，已安装的技能只拉改动文件。适用于模板、参考图等大文件偶发小改的场景：整包动辄数 MB，而一次改动常常只有几 KB。

## 发布端

```bash
python scripts/skill_distribution.py publish --source . --out dist/updates \
  --version 0.3.0 --summary '新增模板页复用' --source-url https://example.com/magic-ppt
```

`release.py` 在打包后自动调用同一函数，`--update-source` 可覆盖写给安装端的地址。产物：

| 路径 | 内容 |
|---|---|
| `dist/updates/manifest.json` | 全部文件的 sha256 与大小 |
| `dist/updates/updates.json` | 版本索引，`history` 按版本记新增、修改、删除 |
| `dist/updates/files/<path>` | 仅本次增量的文件本体 |

第二次发布时，`build_release` 与上一份 `manifest.json` 比对，`history` 追加一条最小差集，`files/` 只放差集涉及的文件。删掉的文件从 `files/` 移除。因此镜像目录不断累积，但每次新增数据量等于实际改动量。

## 安装端

```bash
python scripts/skill_distribution.py check --root <技能目录> --index <updates.json>
python scripts/skill_distribution.py update --root <技能目录> --index <updates.json>
```

`check` 只报告，`plan_update` 取本地版本之后所有版本的并集，因此跳过若干版本也能一次收敛。`update` 按 `update_source` 下载，逐个核对 manifest 的 sha256 后才落盘；任一文件校验失败即整体放弃，不留半更新状态。`VERSION` 作为普通文件参与增量，不再单独写入，避免与自身哈希不符。

## 边界

- 只支持 HTTP 静态镜像。私有仓库需自带鉴权，当前实现不处理凭据。
- 不处理本地修改冲突。安装端有自定义内容时先备份，`update` 会按远端覆盖同名文件。
- 镜像目录需整体上传，`files/` 与两个 JSON 保持相对路径不变。
- 整包 ZIP 仍是主交付物，增量镜像补充它，不能替代首次安装。
