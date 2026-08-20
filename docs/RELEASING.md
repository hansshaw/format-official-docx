# 发布清单

本文用于维护者把 `format-official-docx` 发布到 GitHub。普通使用者只需阅读根目录 README。

## 首次发布前

- 确认仓库公开内容不含真实公文、敏感信息、字体文件、Office/LibreOffice/Poppler 二进制文件或生成的 QA 文件。
- 确认 [LICENSE](../LICENSE) 中的版权主体符合维护者意图；当前使用 `format-official-docx contributors`。
- 在 GitHub 仓库设置中启用 **Private vulnerability reporting**。
- 建议仓库描述：`Codex skill for deterministic Chinese official-document DOCX formatting, auditing, and LibreOffice visual QA.`
- 建议 Topics：`codex`、`agent-skills`、`docx`、`chinese-documents`、`libreoffice`、`python-docx`。
- 推荐默认分支使用受保护规则，要求 `validate` 工作流通过后再合并。

## 发布验证

从干净工作目录创建环境并执行：

```powershell
py -3.12 -m venv .venv
& .\.venv\Scripts\python.exe -m pip install --upgrade pip
& .\.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
& .\.venv\Scripts\python.exe -m compileall -q scripts tests
& .\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

再让 Codex 的 `$skill-creator` 验证当前目录，确认 `SKILL.md` 前置元数据、名称和描述有效。

在装有目标字体、LibreOffice 和 Poppler 的 Windows 环境生成一份包含四级标题、表格和多页正文的虚构样稿，依次执行：

```powershell
& .\.venv\Scripts\python.exe .\scripts\official_docx.py apply `
  .\release-draft.docx .\release-final.docx `
  --title-line '关于开展发布前格式验证的' `
  --title-line '通知' --json

$qaDir = Join-Path $env:TEMP 'official-docx-release-qa'
& .\.venv\Scripts\python.exe .\scripts\verify_official_docx.py `
  .\release-final.docx --output-dir $qaDir --emit-pdf --json
```

逐页查看全部 PNG，并确认没有字体替换、意外换行、页码异常、截断或重叠。

## 版本发布

1. 将 `CHANGELOG.md` 中计划发布的内容从 `Unreleased` 移到新版本，并填写日期。
2. 确认 README、格式规格、技能指令、实现和测试对同一规则描述一致。
3. 合并发布提交后创建带注释标签，例如 `v0.1.0`。
4. 在 GitHub 创建同名 Release，发布说明可以直接使用 CHANGELOG 对应版本内容。
5. 发布后用仓库 URL 通过 `$skill-installer` 做一次全新安装和真实调用。

OpenAI 当前建议在需要更广泛、可安装分发时把技能包装为插件。本仓库首版保持为独立技能；若未来转换为插件，应单独规划目录迁移和兼容安装说明，不要在补丁版本中悄然改变仓库根结构。

