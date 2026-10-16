# 脚本使用与边界

## 依赖与格式选择

在 Codex 桌面中先调用 `load_workspace_dependencies`，使用返回的 Python、Node 与库路径。提取和 HTML 预览用 Pillow；匹配、原生 `.xls` 写入仅用 Python 标准库。独立验证需要 `xlrd==2.0.2`、`olefile==0.47`，如当前环境没有，可安装到本次工作目录的隔离依赖目录，不改系统 Python。

`@oai/artifact-tool` 的公开导入/导出接口针对 `.xlsx`，不能直接保持本模板的旧版 `.xls` 及 WPS 图片流。因此本包的写入使用受限的 BIFF8 内容补丁；Artifact Tool 只作只读预览，不用其 `.xlsx` 导出来替代用户要求保留的原格式。

支持两种输入：

- 独立“未报明细.xls”＋截图文件夹。空白排版模板已内置于 `assets/layout-template.xls`，用户不用单独提供。
- 同结构完整工作簿，截图已经嵌入 WPS 的 `ETCellImageData`。

**明细行数、各类别笔数和图片槽位自动适配，不要求与模板一致。** 原列宽固定，新增明细复制原明细行格式，合计行移到新末行并更新求和范围；截图页按原来的起始列、每排列数和图片/说明双行扩缩。办公费、福利费为 5 列，招聘费、交通费为 2 列。零笔类别保留表头、合计 0 和一个空白图片双行；最后一排不足时留空。福利费原辅助序号按原交错规律同步更新。

目前支持同一套 BIFF8 `.xls` 模板结构：11 个工作表，原始明细“未报”的 A:E 为序号、日期、项目、金额、备注。内置模板保留原明细行、合计行和截图行的样式，不得覆盖或删除其格式。新标签不自动兜底分类。新布局、加密文件、未支持的公式需另行适配。受 `.xls` 65536 行上限约束，超出时明确报错。未做本机 WPS 界面复算。

现成方案评估：WPS 官方提供[按图片名批量插图](https://www.wps.cn/learning/room/d/332361)，但仍需业务匹配。开源 [xlutils](https://github.com/python-excel/xlutils) 采用 xlrd/xlwt 复制；[xlrd 文档](https://github.com/python-excel/xlrd) 明确图片、嵌入对象等不会读取，不能据此保证本模板图片流不丢失。[Apache POI](https://poi.apache.org/components/spreadsheet/) 可编辑 XLS，但本环境未验证其对 WPS 原生图片流的保留。此版复用已验证的原生补丁与独立解析器，按微软的 [Cell Table](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/d690007d-d25c-4822-a650-b24c13288bb7)、[Index](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/67c20922-0427-4c2d-96cc-2267d3f09e8c) 规范扩缩行。

## 独立明细与外部截图

```text
python scripts/prepare_inputs.py 未报明细.xls 截图文件夹 RUN
```

脚本默认定位本技能自己的 `assets/layout-template.xls`；从任意工作目录或解压后的其他位置执行都使用同一内置样式。仅当用户主动更换版式时，才可加 `--template 新模板.xls`。如果资源缺失，应恢复完整技能包，不让用户每个月重新提供模板。

生成 `RUN/prepared-source.xls` 并完成提取。后续命令的 `SOURCE.xls` 必须使用这个中间源表。图片仅接受真实 PNG/JPEG，原文件名用于匹配；完全相同的文件按内容哈希识别。中间表保留模板所有工作表与格式，在“未报”放入实际明细，将文件夹图片导入原生图片库，尚不决定配图位置。

输入文件、模板和图片都只读；准备过程及打包均在本次目录完成。输入完全相同时中间文件可重复生成，确认结果绑定该源文件与图片哈希。不要拿已填图结果代替空白模板。

## 执行顺序

下列 `python`、`node` 指已发现的运行时绝对路径；`SOURCE.xls` 是用户原件，`RUN` 是本次独立目录。命令中的路径均由当前任务提供，不使用历史个人目录。

```text
python scripts/extract.py SOURCE.xls RUN
python scripts/match.py RUN --rules assets/category-rules.json
```

外部输入已经由 `prepare_inputs.py` 完成提取，可省略再次运行 extract。

`raw_rows.json` 是未报原始行；`image_corpus.json` 只有图片标识、名称、原文件哈希和路径，不含分类页或坐标。`template.json` 仅用于版式。`evaluation_only_image_ids.json` 不得提供给匹配程序。

查看 `matches.json` 中 ambiguous/unmatched/unique_amount_candidate 等记录。AI 能通过可见证据解决的情况写入本次 `reviewed_decisions.json`，每项写实际证据；仅金额候选留给用户，不能批量改成已确认。共用付款必须列出所有行、凭证金额和依据。

决定文件示例（键值均需替换成当前输入中真实标识）：

```json
{
  "source_sha256": "当前源文件的真实哈希",
  "matches": {
    "未报!31": {
      "image_key": "当前图片库中的真实key",
      "status": "visual_review",
      "reason": "写出实际看到的实付金额、用途等匹配证据"
    }
  },
  "groups": []
}
```

支持的审核状态：`visual_review`（看图确认）、`equivalent_pair`（相同行组内分配）、`amount_conflict`（金额冲突，保持空白）、`missing_receipt`（无凭证，保持空白）、`user_confirmed`、`user_confirmed_image_conflict`（用户确认了图片，但金额差异仍存在）。

```text
python scripts/match.py RUN --rules assets/category-rules.json --decisions RUN/reviewed_decisions.json
python scripts/build_xls.py SOURCE.xls RUN assets/category-rules.json OUTPUT.xls
python scripts/verify_xls.py SOURCE.xls OUTPUT.xls RUN VALIDATION_DEPS
python scripts/review_report.py RUN 待核对清单.html --decisions RUN/reviewed_decisions.json
```

没有审核决定时省略 `--decisions`。原生补丁保留原目录树与非 Workbook 数据流，修复 BoundSheet、ExtSST、Index 和 DBCell 的偏移；这些必须由验证器通过，不能只看文件成功写出。

`verification.json`、`native-build.json`、`placements.json` 和扩缩行时的 `layout-plan.json` 是验收依据。校验任何失败都不能交付为成品。

## 预览

让 `RUN/node_modules` 以符号链接或 Windows Junction 指向已发现的 bundled node_modules。仅创建链接，不修改依赖目录。

```text
node scripts/render.mjs RUN
```

此脚本读取独立解析最终 `.xls` 后的 `native-view.json`，显示实际保存的单元格值和图片，不读计划值来假装渲染通过。原生格式的权威证据仍是格式记录逐项相同；预览不等于 WPS 实际显示/复算。

## 用户确认后补图

用户在独立 HTML 核对候选，导出 `截图确认结果.json`。该页面离线运行，不上传任何图片。确认文件合并已有审核依据、绑定源文件和图片哈希。

将这个文件作为 `match.py --decisions` 的输入，然后从**同一原始 SOURCE.xls**重新生成另一份 OUTPUT.xls 并验证。页面导出不会直接改表；技能必须实际执行补图和验证，不能仅回复“已确认”。尚未确认的行继续留空。确认图片不授权修改金额或提交报销。

如果确认文件来自另一源表、图片哈希变化、同图重复使用且没有获证实的共用分组，停止相关写入，核对具体冲突。

## 回归检查

```text
python scripts/test_match.py
```

测试覆盖金额候选、同名歧义、未知分类、不同金额不可模糊匹配、共用凭证合计、非法重复使用、日期/金额区分、逐行确认和输入乱序。完整真实回放另外依靠独立验证器；单元测试通过不能替代真实表格验收。

完整扩缩行回归使用获授权的成品样表及其提取目录：

```text
python scripts/test_resize.py SOURCE.xls BASELINE RUN_TESTS VALIDATION_DEPS
```

仅在测试目录生成模拟副本，覆盖 0、1、17、483 笔。用户实测输入包不附历史配图决定，Agent 必须独立看图判断。
