# 费用分类与截图匹配 Skill

只需提供未报明细 `.xls` 和截图文件夹，按内置 WPS 模板分类汇总、匹配凭证并生成结果副本。明细与截图行数按本次数据自动增减，沿用原样式。

## 安装

公开仓库：https://github.com/stonele01/expense-receipt-workflow

在 Codex 中发送：

```text
请使用 skill-installer 安装 https://github.com/stonele01/expense-receipt-workflow 的根目录技能，技能名 expense-receipt-workflow；完整保留 assets、scripts、references 和 agents。安装后告诉我位置。
```

使用 Codex 内置安装脚本时，参数为 `--repo stonele01/expense-receipt-workflow --path . --name expense-receipt-workflow`。安装后下一轮对话即可使用。

其他 Agent 可下载本仓库 ZIP，完整解压后读取根目录 `SKILL.md`；需具备 Python、图片查看和本地文件操作能力。不能只复制 SKILL.md。模板可离线使用，依赖与限制见下文及操作手册。

## 使用

将整个仓库交给支持本地文件、Python 和图片查看的 Agent，要求先阅读 `SKILL.md`。内置模板位于 `assets/layout-template.xls`，不可遗漏。操作步骤与依赖见 `references/runbook.md`。

可直接使用以下提示词，将路径换为本机实际位置：

```text
请读取本技能的 SKILL.md。我的未报表是【完整路径】，截图目录是【完整路径】，输出目录是【完整路径】。
使用技能内置模板，按实际笔数自动增减明细行和图片行，保持原格式。只填有明确匹配依据的截图；仅金额相同、金额冲突或不能确认的留空并输出独立待核对清单。原文件只读。完成后执行独立验证，报告分类金额、已填图及待核对数量，说明是否经过 WPS 实际打开复算。
```

## 验证与边界

已验证 0、1、17、483 笔扩缩行，以及独立明细加外部截图的完整流程。另一 Agent 的 161 笔试跑结果已独立复核：118 笔配图、43 笔待核对，样式、合计、公式及内部索引校验通过。

适用于本技能内置的 BIFF8 XLS 与 WPS 原生图片结构，不是通用 Excel 转换器。尚未完成本机 WPS 界面打开复算。待核对项不等于已完成报销。

仓库仅含技能、代码与清空业务数据的格式模板，不包含个人费用明细、收款截图或当前审核决定。
