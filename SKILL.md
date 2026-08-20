---
name: format-official-docx
description: "Apply the user's mandatory Chinese official-document formatting to every Word or DOCX creation, edit, conversion, and delivery: A4 portrait with fixed margins; 方正小标宋简体 main title; four numbered body-heading levels in 黑体、楷体_GB2312 or 仿宋_GB2312 with required punctuation; 16 pt 仿宋_GB2312 body with a 2-character first-line indent; exact 29 pt spacing; and centered 宋体 '- PAGE -' numbering. Use whenever Codex outputs a .docx for this user unless the current request explicitly requires another format or a supplied template controls the layout."
---

# 中文公文 Word 格式

## 核心契约

- 与 documents 技能共同使用，并继续遵守其受控运行时、操作标记和“渲染后逐页检查”门禁。本技能的固定公文规格就是版式权威，不再读取通用设计预设、比较视觉方案或添加装饰性页眉和版面变体。
- 对本技能处理的 DOCX，`verify_official_docx.py` 是最终渲染门禁的技能内实现，等价替代单独调用 documents 的 `render_docx.py`；不得再运行第二个渲染器。它仍会使用隔离的 LibreOffice 配置并生成 `page-N.png`，因此逐页视觉检查要求不变。
- 在每次交付 DOCX 前应用本技能。仅当当前请求明确指定其他版式，或用户提供的模板被明确指定为版式权威时，才跳过或覆盖相应规则。
- 先读取 [format-spec.md](references/format-spec.md)，再确定主标题的语义断行和正文标题层级。
- 保留输入文件，默认写入新的输出文件；不得直接覆盖原稿。
- 不把脚本审计当作视觉验收。审计通过后仍须渲染并检查每一页。

## 标准流程

1. 使用 documents 技能的工作区依赖创建或编辑内容草稿。
2. 将主标题放在一个顶层段落中。短标题使用一行；预计会自动换行的主标题必须按语义显式分行。
3. 让主标题各行形成梯形或菱形。不得用空格、Tab、文本框或字号变化造形。
4. 从本技能目录运行格式化命令，并为多行标题重复传入 `--title-line`。`$skillDir` 指向已安装技能目录，`$workspace` 指向工作区；输出路径始终放在工作区。例如：

~~~powershell
& $workspacePython "$skillDir\scripts\official_docx.py" apply "$workspace\draft.docx" "$workspace\final.docx" --title-line "关于开展年度重点工作的" --title-line "通知"
~~~

默认不替换已经存在的输出文件；只有确认要替换时才追加 `--force`。

5. `apply` 已包含结构审计。对最终文件运行一次合并验收命令，由技能脚本再次确认结构并使用系统 LibreOffice 生成逐页 PNG：

~~~powershell
& $workspacePython "$skillDir\scripts\verify_official_docx.py" "$workspace\final.docx" --output-dir "$workspace\final-qa"
~~~

6. 以 100% 缩放检查 `final-qa` 中的全部页面，重点核对主标题实际形状、四级正文标题的编号/字体/标点及无首行缩进、普通正文首行缩进、29 磅固定行距、分页和页码。
7. 仅在脚本审计与逐页视觉检查均通过后交付最终 DOCX。

## 高效验收规则

- 固定公文版式不做视觉方案探索、设计预设比较或多轮“美化”审查。
- 默认只创建或修改内容一次、运行 `apply` 一次、运行最终合并验收一次、逐页检查一次。
- 只有结构审计失败或页面出现明确的截断、重叠、错误换行、字体替换、分页或页码缺陷时，才修复并重新运行合并验收。
- `verify_official_docx.py` 使用独立的临时 LibreOffice 用户配置，自动发现 PATH、`C:\Program Files\LibreOffice\program\soffice.exe` 和 32 位程序目录，避免进程锁冲突；它已经替代单独的 `render_docx.py` 调用，不要启动 LibreOffice 图形界面或第二个渲染器做常规审计。

## 主标题选择

- 默认把第一个非空正文顶层段落视为标题；包在内容控件中的正文段落也会被识别。
- 当标题不是第一个非空正文顶层段落时，使用从 0 开始的 `--title-index` 明确指定。
- 当标题超过单行可用宽度时，必须提供多个 --title-line。脚本拒绝依赖 Word 自动换行的过长标题。
- 优先保持词语、机关名称、法规名称和修饰关系完整，再调整各行长度。
- 脚本只检查几何形状；语义完整性必须由代理结合逐页渲染人工确认。

## 正文标题层级

仅把正文顶层的独占段落识别为层级标题；主标题、表格和文本框中的相似编号不参与识别。编号前不得放置空格或 Tab，编号后必须包含实际标题文字。

| 层级 | 编号 | 字体字号 | 句末标点 |
| --- | --- | --- | --- |
| 一级 | `一、` | 黑体，三号 | 无 |
| 二级 | `（一）` | 楷体_GB2312，三号 | 可有可无 |
| 三级 | `1.` | 仿宋_GB2312，三号 | 必须有 |
| 四级 | `（1）` | 仿宋_GB2312，三号 | 必须有 |

一级、二级未另行指定字号，因此沿用正文三号（16 磅）。使用全角括号 `（ ）` 和三级编号中的半角句点 `.`；不要写成 `(一)`、`(1)` 或 `1．`。层级标题不设首行缩进，不得同时带 Word 自动编号，并设置与下段同页、段中不分页。

## 修改范围

格式化命令执行以下确定性变更：

- 将所有节统一为 A4 纵向和指定页边距，关闭镜像页边距、奇偶页脚及首页不同页脚，并取消分节页码重置。
- 将普通正文、表格单元格、超链接和正文中的文本框运行统一为仿宋_GB2312 三号；正文顶层的层级标题按上表覆盖字体。
- 将所有主体段落设为 29 磅固定行距，段前和段后为 0。
- 将普通正文顶层段落（含内容控件中的正文，排除主标题、层级标题、表格单元格和文本框）设为首行缩进 2 字符；层级标题清除首行、悬挂缩进和直接自动编号，改用 Normal 样式以阻断样式继承的冲突格式。
- 将标题设为方正小标宋简体二号、居中、与下段同页，并保留显式换行。
- 为每个节生成独立、居中的动态 PAGE 页脚，显示为 - 1 -，并标记字段在打开时更新。

页眉、脚注、尾注、批注及嵌入对象不属于自动正文范围；存在这些内容时，按同一规格单独检查并在渲染中确认。

## 失败处理

- 缺少方正小标宋简体、黑体、楷体_GB2312、仿宋_GB2312 或宋体时，停止“完全合规”的表述；安装字体或向用户披露替换风险。
- 审计失败时，修复后重新运行 apply 和合并验收（含逐页检查）；不要再额外运行独立 `audit` 或第二个渲染器，也不得只修改审计器或忽略告警。
- 若 LibreOffice 缺失或无法启动，先报告具体错误；无法安装时才按 documents 技能的结构验收回退流程交付，并明确说明未完成视觉渲染检查。
