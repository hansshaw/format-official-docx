# 第三方依赖说明

本仓库只包含技能说明和 Python 脚本，不捆绑以下依赖。使用者应从各项目官方渠道获取并遵守对应许可证。

| 依赖 | 用途 | 本仓库中的关系 | 官方地址 |
| --- | --- | --- | --- |
| `python-docx` | 读取和修改 DOCX | Python 运行依赖；MIT | <https://python-docx.readthedocs.io/> |
| `lxml` | OOXML/XML 解析 | Python 运行依赖；BSD-3-Clause | <https://lxml.de/> |
| LibreOffice | 将 DOCX 转为 PDF | 本地可选渲染依赖；Mozilla Public License 2.0 | <https://www.libreoffice.org/> |
| Poppler | 将 PDF 渲染为 PNG | 本地可选渲染依赖；GPL-2.0-or-later | <https://poppler.freedesktop.org/> |

目标中文字体（包括方正小标宋简体、仿宋_GB2312、楷体_GB2312、黑体和宋体）不由本项目提供，也不属于本项目的许可证范围。请从合法来源取得字体并自行承担授权责任。

`python-docx`、`lxml`、LibreOffice、Poppler、Microsoft Word 和 Codex 是其各自权利人的项目或商标。本仓库不表示与这些项目存在隶属、认证或背书关系。
