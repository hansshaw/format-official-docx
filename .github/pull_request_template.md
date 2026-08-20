## 变更说明

说明本 PR 修改的规则、行为或文档，以及修改原因。

## 验证

列出运行的命令和结果。涉及渲染时，说明使用的 Word/LibreOffice、字体和逐页检查结果。

## 检查清单

- [ ] 修改范围单一且与 Issue 或说明一致
- [ ] 已运行 `python -m compileall -q scripts tests`
- [ ] 已运行 `python -m unittest discover -s tests -v`
- [ ] 格式规则变化已同步更新实现、审计、测试、README、规格和 CHANGELOG
- [ ] 影响渲染的变化已用目标字体、LibreOffice 和 Poppler 做逐页检查
- [ ] 未提交字体、二进制依赖、真实公文、敏感信息或生成的 QA 文件
- [ ] 安全相关内容已通过私密渠道报告

