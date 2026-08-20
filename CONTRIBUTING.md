# 贡献指南

感谢你改进 `format-official-docx`。本项目优先保证格式规则可解释、脚本行为确定、审计结果可复现，并尽量减少生成 Word 文档时不必要的渲染轮次。

## 开始之前

- 普通缺陷、兼容性问题和规则建议可以提交 GitHub Issue。
- 安全问题请按 [SECURITY.md](SECURITY.md) 私下报告，不要公开附带敏感文档或利用细节。
- 功能建议应说明依据的格式规则、预期适用范围，以及是否与现有规则冲突。
- 不要把含真实姓名、单位、联系方式、印章、内部编号或其他敏感信息的 DOCX 上传到公开 Issue。

## 本地环境

需要 Python 3.10 或更高版本：

```powershell
py -3.12 -m venv .venv
& .\.venv\Scripts\python.exe -m pip install --upgrade pip
& .\.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
```

如需验证最终渲染，还需要合法安装目标中文字体、LibreOffice 和 Poppler。普通单元测试不依赖这些外部程序。

## 修改原则

- 保持 `SKILL.md` 聚焦于 Agent 真正需要遵守的工作流；精确参数和边界放在 `references/format-spec.md`。
- 重复、确定性的 DOCX 操作应实现到 `scripts/`，不要要求 Agent 每次临时重写。
- 修改格式规则时，同时更新实现、审计、测试、README、规格参考和 CHANGELOG。
- 格式化命令应保留输入文件，并默认拒绝覆盖输出。
- 新的审计失败必须给出可操作的中文错误信息。
- 不要提交字体、Office 安装包、LibreOffice、Poppler 二进制文件、真实公文或生成的 QA 页面。

## 测试

提交前运行：

```powershell
& .\.venv\Scripts\python.exe -m compileall -q scripts tests
& .\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

影响渲染、分页、页码或字体的修改还应运行真实验收：

```powershell
$qaDir = Join-Path $env:TEMP 'official-docx-contributor-qa'
& .\.venv\Scripts\python.exe .\scripts\verify_official_docx.py `
  .\path\to\formatted.docx --output-dir $qaDir --json
```

逐页查看所有 `page-N.png`。测试夹具必须使用虚构、去标识化内容，并在测试结束后清理。

## Pull Request

一个 PR 尽量只解决一个明确问题。说明：

- 修改了什么规则或行为；
- 为什么现有行为不正确或不足；
- 如何复现和验证；
- 是否改变向后兼容性、CLI、字体检查或渲染结果；
- 是否已更新文档与 CHANGELOG。

提交 PR 即表示你有权按本项目 MIT License 提交相关内容。

