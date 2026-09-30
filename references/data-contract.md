# 表格数据与指标合同

仅在输入包含 XLSX、CSV 或 Word 表格，或需要以数据驱动柱图、排名、占比等对象时读取本页。普通口述稿继续使用 `extract`、`sources` 与 `facts`；两条路线可以同处一个 `plan.json`。

## 建立可追溯的单元格账本

```powershell
python -X utf8 "<skill>/scripts/personal_ppt.py" extract-data "<项目>/统计表.xlsx" --id report --out "<项目>/report.data.json"
python -X utf8 "<skill>/scripts/personal_ppt.py" check-data "<项目>/plan.json"
```

`extract-data` 支持 `.xlsx`、UTF-8/GB18030 `.csv` 和 `.docx` 中的表格。XLSX 地址如 `report:Sheet1!B7`；CSV 地址如 `report:r000007c0002`；Word 表格地址如 `report:t001r0007c002`。账本保存源文件 SHA256、单元格原值、公式、缓存值、有效值、数字格式、`blank`/`zero`/`number`/`text`/`formula_uncached` 状态与位置。Excel 中存储为 `0.2`、显示为 `20%` 的单元格，保留原值 `0.2`，数值指标使用显示百分数 `20`。未缓存的公式不能作为数值证据。表格单元格和普通 DOCX 段落可分别提取，不能把单元格账本冒充完整口述稿覆盖记录。

在计划顶层加入相对路径 `data_sources:["report.data.json"]`。每次 `check-plan` 或 `build` 都会从原文件重新提取并逐字段对照账本；表格被修改后，旧账本与旧质检失效。普通 `sources`、`facts` 和旧计划保持有效。

## 显式定义指标

账本只证明源位置及其值，不推断组织、时间、统计口径。计划作者必须为每个指标提供 `id`、`kind`、`value`、`unit`、`period`、`population`、`measure` 和 `value_type`。`value_type` 取 `actual`、`target`、`estimate` 或 `comparison`。这些维度应来自原材料或标记为待核对，不能为了通过校验编造。占比来源还要写分母 ID 或分母描述。

```json
{
  "data_sources": ["demo.data.json"],
  "metrics": [
    {"id":"completed","kind":"source","value":3,"unit":"项","period":"演示季度","population":"演示项目","measure":"已完成项目数","value_type":"actual","source_ref":"demo:r000003c0002","slice":"已完成"},
    {"id":"total","kind":"source","value":6,"unit":"项","period":"演示季度","population":"演示项目","measure":"项目数","value_type":"actual","source_ref":"demo:r000008c0002","slice":"全部"},
    {"id":"completed_share","kind":"derived","value":50,"unit":"%","period":"演示季度","population":"演示项目","measure":"已完成项目占比","value_type":"actual","operation":"ratio_percent","operands":["completed","total"],"denominator":"total","relation":"subset","decimals":0}
  ]
}
```

上例是合成值，仅解释格式；真实计划中的项目数及状态必须以相应源单元格实际内容为准。`source_ref` 不允许指向普通口述段落或不存在的单元格；填写数值必须与源单元格一致。可把指标 ID 放入原有 `slots[].facts`、`chart.fact_ids`、`bars[].fact`、`ranking.rows[].fact`。由此页面文字、图表缓存与原生对象共用同一个值，但标题的解释与统计对象仍需人工审核。

空白、文本“尚未统计”和无缓存公式可声明 `value:null` 及 `missing_reason`，不得当作零值用于派生运算或可见数字槽位。真实的 `0` 必须来自数值为零的源单元格。图片和过渡页的事实来源依旧按各自合同绑定。

Word 表格若列出的是类别或状态而不是数字，可用 `kind:"table_count"` 在完整连续行范围内确定性计数，不必手算后假装它是原表数值。`query` 指定 `source_id`、`format`、`column`、`first_row`、`last_row`，DOCX 另需 `table`，XLSX 另需 `sheet`。筛选方式三选一：`equals:["完整单元格值"]`、`starts_with:["状态前缀"]` 或 `nonblank:true`。例如：

```json
{"id":"completed","kind":"table_count","value":3,"unit":"项","period":"演示季度","population":"演示表列示项目","measure":"已完成项目数","value_type":"actual","count_basis":"状态以‘已完成’开头的项目行","query":{"source_id":"demo","format":"docx","table":1,"column":4,"first_row":2,"last_row":7,"starts_with":["已完成"]}}
```

该实例是合成示意，不对应任何真实材料。行范围及状态词由计划作者按原表确定并复核；脚本检查每个单元格存在、匹配数量与声明值相同。它无法知道原表是否遗漏了应统计的项目，也无法证明“状态前缀”在业务上等于实际完成。

## 确定性运算与比较边界

| `operation` | 输入顺序与条件 | 输出 |
| --- | --- | --- |
| `sum` / `mean` | 至少两项，同单位、时期、总体、指标定义、值属性；各操作数有不同 `slice`；写明 `aggregation_basis` | 原单位 |
| `difference` | 两项，写 `comparison:category|period|target`；相同单位、总体和指标定义，时期/值属性须符合比较类型 | 原单位 |
| `ratio_percent` | `[分子,分母]`，写 `denominator` 与 `relation:subset|rate`；单位、时期、总体和值属性相同，分母非零；subset 必须在 0–100% | `%` |
| `growth_percent` | `[本期,基期]`，同单位、总体、指标定义和值属性；时期不同、基期大于零；写 `comparison_basis` | `%` |
| `percentage_point_change` | `[本期%,基期%]`，同总体、指标定义和值属性；时期不同；写 `comparison_basis` | `百分点` |
| `target_progress_percent` | `[实际,目标]`，同单位、时期、总体、指标定义；目标非零，显式指定目标为 `denominator` | `%` |

派生数值由脚本重新计算，按 `decimals`（0–8，默认 1）四舍五入后与声明值比较；循环、零分母、错单位/时期、目标冒充实际或估计值混算会被拒绝。`comparison_basis` 和 `aggregation_basis` 是需要审阅的口径说明，程序无法单靠字符串证明两期数据真的可比或两个类别真的互斥。来源没有给出比较条件时，应列为材料缺口，不应生成看似精确的结论。

`check-data` 输出每项指标的值、单位、时期、总体、定义、值属性、单元格来源或派生操作数与分母，可作为逐图证据卡的机器部分。它不代替对原表标题、脚注、统计对象和可见图表含义的审阅。最终仍须在 PowerPoint 渲染成片后核对数值、轴、标签、图例及讲述备注。
