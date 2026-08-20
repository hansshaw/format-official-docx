# format-official-docx

一个面向 Codex 的中文公文 Word 格式技能。它把现有 `.docx` 统一为确定的公文版式，审计 OOXML 结构，并通过 LibreOffice 和 Poppler 一次性渲染逐页 PNG，供最终视觉检查。

本项目适合需要稳定复用以下版式规则的场景，不试图覆盖所有党政机关公文要素，也不会替代人工校对。

> [!IMPORTANT]
> 这是固定且有明确取舍的个人版式配置，不是通用 Word 格式器、GB/T 9704 全量实现、合规证明或任何政府机构、OpenAI、Microsoft、The Document Foundation 的官方认证。技能保持 Codex 默认的隐式调用策略；安装者采用它作为默认规则后，Codex 可以在匹配的 `.docx` 任务中自动使用。当前请求明确指定其他版式或提供权威模板时，应以当前请求或模板为准。

## 格式规则

| 项目 | 规则 |
| --- | --- |
| 页面 | A4 纵向；上 3.7 cm、下 3.5 cm、左 2.8 cm、右 2.6 cm |
| 主标题 | 方正小标宋简体，二号（22 pt），居中；多行时按语义排成梯形或菱形 |
| 正文 | 仿宋_GB2312，三号（16 pt）；首行缩进 2 字符 |
| 段落 | 固定行距 29 pt；段前、段后 0 |
| 页码 | 宋体，四号（14 pt），居中动态字段，形如 `- 1 -` |
| 一级标题 | `一、`；黑体三号；句末无标点 |
| 二级标题 | `（一）`；楷体_GB2312 三号；句末标点可有可无 |
| 三级标题 | `1.`；仿宋_GB2312 三号；句末必须有标点 |
| 四级标题 | `（1）`；仿宋_GB2312 三号；句末必须有标点 |

脚本还会拒绝半角括号、三级标题全角句点、编号前空格、空标题、自动编号叠加等容易造成版式偏差的输入。完整的 OOXML 参数和识别边界见 [格式规格](references/format-spec.md)。

## 工作方式

- `apply`：格式化输入 DOCX，写入新文件，并立即执行结构审计。
- `audit`：只读审计已有 DOCX，不修改文件。
- `verify`：先审计，再用隔离的 LibreOffice 配置转为 PDF，最后用 Poppler 生成逐页 PNG。
- 技能指令要求 Agent 在交付前查看全部页面，确认字体替换、换行、分页、页码、截断和重叠等结构审计无法判断的问题。

`apply` 不允许输入、输出为同一路径，默认也不覆盖已有输出。只有明确需要替换时才使用 `--force`。

Codex 工作流会配合其文档处理能力创建或编辑内容；两个 Python CLI 可以独立运行。CLI 只接受 `.docx`，不直接处理旧版 `.doc`、含宏 `.docm`、ODT 或 Google Docs。

格式化不是局部样式修补。它会统一正文和相关运行的字体与段落设置，重写 `Normal`/正文标题样式、所有节的纸张和页边距，关闭镜像页边距、奇偶页脚和首页不同页脚，取消分节页码重置，并清空后重建每节页脚。原有页脚内容会丢失；全局样式变化也可能间接影响未被脚本逐段处理的内容。请始终保留原稿并检查输出。

## 环境要求

- Python 3.10 或更高版本。
- Python 包：`python-docx`、`lxml`。
- 最终渲染：LibreOffice `soffice` 和 Poppler `pdftoppm`。
- 完全符合目标版式所需字体：方正小标宋简体、仿宋_GB2312、黑体、楷体_GB2312、宋体。

安装 Python 依赖：

```powershell
py -3.12 -m venv .venv
& .\.venv\Scripts\python.exe -m pip install --upgrade pip
& .\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

Windows 下可安装 LibreOffice，并确保 Poppler 的 `pdftoppm.exe` 位于 `PATH`，或在验证时通过 `--soffice`、`--poppler-dir` 指定路径。Linux 和 macOS 也可使用 PATH 中的对应程序，但当前自动发现逻辑和前向测试主要覆盖 Windows。

字体文件不会随本仓库分发。请从合法来源取得并遵守相应许可；缺少目标字体时，Word 或 LibreOffice 可能静默替换字体，此时不能声称文档“完全合规”。

缺失字体在 CLI 中属于警告，不一定让审计失败：命令可能返回退出码 `0` 并同时输出字体 warning。`--skip-font-check` 也只是不检查本机安装状态，绝不代表字体合规。

## 安装到 Codex

按照 [OpenAI 官方 Codex 技能文档](https://developers.openai.com/codex/skills/)，Codex 会从用户级和仓库级 `.agents/skills` 目录发现技能，也可以让 `$skill-installer` 从 GitHub 仓库安装。

用户级安装时，将本仓库放到：

```text
Windows: %USERPROFILE%\.agents\skills\format-official-docx
macOS/Linux: ~/.agents/skills/format-official-docx
```

为兼容部分较早的 Codex 本地配置，也可以放在 `$CODEX_HOME/skills/format-official-docx`（未设置 `CODEX_HOME` 时通常是 `~/.codex/skills/format-official-docx`）。新安装优先采用上面的官方当前路径。

仓库级安装时，将目录放到目标项目的：

```text
.agents/skills/format-official-docx
```

也可以在 Codex 中调用 `$skill-installer`，并把本仓库的 GitHub URL 交给它。安装后可显式调用：

```text
使用 $format-official-docx 创建一份中文公文 Word 文档。
```

Codex 通常会自动发现技能变更；若技能没有出现，重启 Codex。

本仓库保持为可直接安装的独立技能。OpenAI 官方目前建议在面向更广泛用户做统一安装分发，或需要同时捆绑多个技能、连接器和展示资源时，将技能包装为插件；这不是运行本技能的前置条件。

## 命令行使用

以下示例假设当前目录是本仓库根目录，并已安装 Python 依赖。

### 格式化

```powershell
& .\.venv\Scripts\python.exe .\scripts\official_docx.py apply `
  .\draft.docx .\final.docx `
  --title-line '关于开展年度重点工作的' `
  --title-line '通知' `
  --json
```

`--title-line` 可以重复使用，并会替换主标题内容。多行标题应先按语义断行，再形成单调变长、单调变短的梯形，或先变长再变短的菱形。例如：

```text
关于开展年度重点工作的
通知
```

不要用空格、Tab、文本框或不同字号人为造形。脚本检查几何形状，语义完整性仍需人工确认。

### 只读审计

```powershell
& .\.venv\Scripts\python.exe .\scripts\official_docx.py audit `
  .\final.docx --json
```

字体由其他受控环境提供时，可对 `audit` 使用 `--skip-font-check`。该选项只跳过当前机器的字体安装检查，不改变 DOCX 中的字体要求。

### 最终渲染验收

```powershell
$qaDir = Join-Path $env:TEMP 'official-docx-qa'
& .\.venv\Scripts\python.exe .\scripts\verify_official_docx.py `
  .\final.docx `
  --output-dir $qaDir `
  --emit-pdf `
  --json
```

验证器默认以 170 DPI 生成 `page-1.png`、`page-2.png` 等文件。请始终使用专用 QA 目录：运行前，它会清理该目录内已有的 `page-*.png` 和 `official-page-*.png`。生成的 PDF 和 PNG 含完整文档内容，应与原 DOCX 按同一敏感级别保存和清理。

常用参数：

| 命令 | 参数 | 说明 |
| --- | --- | --- |
| `apply` | `--title-index N` | 指定从 0 开始的正文顶层标题段落索引 |
| `apply` | `--title-line TEXT` | 指定一行主标题，可重复 |
| `apply` | `--force` | 允许覆盖已存在的输出文件 |
| `audit` | `--skip-font-check` | 跳过本机字体安装检查 |
| `verify` | `--output-dir DIR` | 指定 QA 输出目录 |
| `verify` | `--soffice FILE` | 显式指定 LibreOffice 可执行文件 |
| `verify` | `--poppler-dir DIR` | 指定包含 `pdftoppm` 的目录 |
| `verify` | `--dpi N` | PNG 分辨率，范围 72 至 600，默认 170 |
| `verify` | `--timeout N` | 单个外部程序的超时秒数，默认 120 |
| `verify` | `--emit-pdf` | 在 QA 目录保留中间 PDF |

退出码约定：`0` 表示结构通过（仍可能包含 warnings），`1` 表示结构审计已完成但不合规，`2` 表示参数、路径、内容或外部程序运行错误。

## 识别范围与边界

- 默认把第一个非空正文顶层段落视为主标题；可以用 `--title-index` 覆盖。
- 多行主标题使用约 20 个全角视觉单位和梯形/菱形启发式检查；脚本不会自动判断语义断行是否合理。
- 四级标题只按正文顶层段落开头的字面编号识别。表格、文本框和页脚中的相似编号不会被当作正文标题；脚本也不能可靠判断“编号后整段文字”在语义上是不是小标题。
- 普通正文顶层段落应用 2 字符首行缩进；主标题、各级层级标题、表格、文本框和页脚不应用该缩进。
- 页眉、脚注、尾注、批注、嵌入对象和部分独立 OOXML 部件不在自动正文范围内，需要单独检查。
- 本项目落实本 README 所列规则，但不声称完整实现 GB/T 9704，例如版记、发文机关标志、红色分隔线、签发人、印章位置和单双页外侧页码并未全部覆盖。
- Windows 字体检查基于系统字体目录中的文件名；仅安装在用户字体目录的字体可能出现误警告。非 Windows 环境依赖 `fc-match`。
- `verify` 的 PASS 只表示结构审计通过且 LibreOffice/Poppler 成功生成页面，不会自动检测错页、重叠、截断或字体替换。仍须人工逐页查看。
- LibreOffice 与 Microsoft Word 的字体替换、分页和字段更新可能不同，最终文件仍应在实际交付环境中抽查。
- 本项目不是文档净化器，不保证清除作者元数据、批注、修订、隐藏文本、自定义 XML、外部关系或嵌入对象。
- 请勿用未隔离的高权限账户处理来源不可信的 DOCX。详见 [安全策略](SECURITY.md)。

## 项目结构

```text
format-official-docx/
|-- SKILL.md                       # Agent 工作流与强制门禁
|-- agents/openai.yaml             # Codex UI 元数据
|-- references/format-spec.md      # 精确版式和 OOXML 规格
|-- scripts/official_docx.py       # 格式化与结构审计
|-- scripts/verify_official_docx.py # LibreOffice + Poppler 合并验收
|-- tests/                         # 跨平台结构与回归测试
`-- .github/                       # CI、Issue 与 PR 模板
```

## 开发与验证

安装开发依赖后运行：

```powershell
& .\.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
& .\.venv\Scripts\python.exe -m compileall -q scripts tests
& .\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

普通 CI 只做确定性的结构与回归测试，不安装 LibreOffice，以避免每次提交都进行耗时的视觉渲染。发布前还应在装有目标字体、LibreOffice 和 Poppler 的 Windows 环境执行一次 `verify_official_docx.py`，逐页检查生成的 PNG。完整流程见 [发布清单](docs/RELEASING.md)。

## 参与贡献

提交问题或代码前请阅读 [贡献指南](CONTRIBUTING.md)。安全问题请不要公开提交，按 [SECURITY.md](SECURITY.md) 中的方式报告。版本变化记录在 [CHANGELOG.md](CHANGELOG.md)。

## 许可证

代码与技能说明按 [MIT License](LICENSE) 开源。字体、LibreOffice、Poppler、Microsoft Word 及其他第三方组件各自适用其许可证，本仓库不分发这些组件；详见 [第三方依赖说明](THIRD_PARTY_NOTICES.md)。
