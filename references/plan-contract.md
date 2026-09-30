# plan.json v1

路径：`sources`相对于plan文件；源账本中的`path`相对于账本文件；图片相对于plan文件。输出版本保留绝对输入路径和指纹用于追踪。转移项目后应重新构建验证，不沿用旧交付签章。

顶层字段：`schema_version:1`、`title`、`template_sha256`、`sources:[账本路径]`、`facts:[]`、`pages:[]`、`gaps:[]`。可加入 `storyboard` 记录讲述意群与页时长；可选的 `data_sources`/`metrics`、`transitions`、`media_inventory`/`project_matches`/`asset_dispositions` 见下文。模板SHA从doctor或profile读取。页ID稳定且唯一，不等于最终页码。

## 讲述意群与整套节奏

先阅读 [叙事与排版](narrative-planning.md)。推荐把意群与页面分配写入 `storyboard`；字段可选以兼容既有计划，但提供后会在 `check-story`、`check-plan` 和 `build` 中确定性核验：

```json
{
  "storyboard": {
    "duration_seconds": 180,
    "beats": [
      {"id": "B01", "expected_seconds": 60},
      {"id": "B02", "expected_seconds": 120}
    ],
    "page_assignments": [
      {"page_id": "opening", "beat_id": "B01", "seconds": 60},
      {"page_id": "example", "beat_id": "B02", "seconds": 80},
      {"page_id": "close", "beat_id": "B02", "seconds": 40}
    ]
  }
}
```

每个意群和页面都要有分配；同一意群可拆到多页，但各页秒数之和必须等于该意群的预计时长；总时长须等于全部意群时长。同一 `(page_id, beat_id)` 不能重复计时。`page_id` 使用稳定的 `pages[].id`。这些检查阻止遗漏和重复计时，**不自动评价**叙事顺序是否有说服力、图标是否合义；后者记录在整套视觉审阅中。页上可附 `beat_ids`、`rhetorical_role`、`density_intent`、`selection_reason` 和 `transition_to_next`，用于人工判断。

意群切换可再写可选 `transitions`。先运行 `suggest-transitions` 看候选边界，再写明听众问题、桥接句和选择理由；`micro` 仅补讲述备注，`content`/`pause` 按新判断的必要性插原生页。写了决策就必须运行 `apply-transitions`，未应用的决策不能直接构建。停顿页的 PART 标签和幽灵章节号须同步更新。完整字段见 [过渡页合同](transition-planning.md)，不得按固定页数插页。

## 页面和文本

```json
{
  "id": "case-demo",
  "template_slide": 11,
  "purpose": "说明一次示例活动的成果",
  "slots": {
    "6": {"action":"replace", "kind":"inference", "text":"示例活动形成可复用成果", "sources":["manuscript:p016"]},
    "10": {"action":"replace", "kind":"source", "paragraphs":[{"runs":[{"text":"▪ ","style_run":0},{"text":"合成示例照片说明","style_run":1}]}], "sources":["manuscript:p016"]}
  },
  "pictures": {},
  "notes_refs": ["manuscript:p016"],
  "coverage": []
}
```

示例只展示接口，不能直接构建：所选页全部有文字的槽位（页码除外）必须显式绑定。`profile.json`中的`key`是原生ID祖先路径；同名对象不可当唯一标识。

- `action`：`replace`、`keep`、`remove`。remove须有reason；keep仅允许来源明确为同一参考稿的evidence，不可用于未来报告偷带旧事实。
- `kind`：source、derived、inference、proposal、gap、label。label仅用于通用标签，不能规避事实校验。gap允许构建草稿，阻断finalize。
- 简单同格式文本可用`text`，换行生成段落。原槽位包含多个格式时必须使用`paragraphs`和`style_run`；索引是原槽位全部a:r的0起始序号。
- `style_paragraph`可指定原段落样式；不指定按对应段落继承。样式属性不从计划重新指定，保持参考。
- 页码自动更新。页内引用用普通文本`text:"见第{page}页"`和`target_page:"目标页ID"`，构建时替换。
- `notes_refs`按顺序原文写入备注；`notes_extra`补充本页讲述衔接或材料限制。不得隐去来源中不利于结论的限定条件。
- 段落需跨页分配时，用`notes_segments:[{"source":"来源:p020","quote":"原文连续片段"}]`。逐字验证引用，只有所有片段覆盖原段落后才算完整覆盖。
- `coverage`对象包含`source`、`destination:slide|notes|omit`和`reason`。primary每个段落必须在notes_refs或coverage出现。slide覆盖还要求该段落绑定到上屏槽位。

## 事实与运算

```json
[
 {"id":"total","kind":"source","value":80,"unit":"次","period":"示例季度","source":"manuscript:p007","quote":"共记录80次服务"},
 {"id":"focus","kind":"source","value":20,"unit":"次","period":"示例季度","source":"manuscript:p009","quote":"重点项目20次"},
 {"id":"share","kind":"derived","value":25,"unit":"%","period":"示例季度","operation":"ratio_percent","operands":["focus","total"],"decimals":0}
]
```

支持sum、mean、difference、ratio_percent，拒绝循环、零分母、单位/时期不一致。槽位用`facts:["share"]`绑定；derived槽位必须有facts。可用`visible_qualifiers`约束该槽位同时出现的学校/年份/指标名。计算核验不替代来源语义审阅。

若有 XLSX、CSV 或 DOCX 表格，先用 `extract-data` 建逐单元格账本，并在顶层写 `data_sources` 与 `metrics`。指标 ID 会合入 `facts` 命名空间，供文字、柱形、排名和原生图表共用。指标必须写总体、指标定义、时期、单位和值属性；未缓存公式和缺失值不能当零。比较、增长、百分点变化、目标进度需要显式分母和可比口径，详见 [数据合同](data-contract.md)。

## 图片、图表、排名

非装饰照片全部需要pictures绑定：

```json
{"34":{"action":"replace","path":"photos/event.jpg","sha256":"实际SHA256","role":"evidence","generated":false,"sources":["manuscript:p016"],"crop":{"l":0,"t":0,"r":0,"b":0}}}
```

不指定crop时保留原裁切；指定时使用OOXML百分比单位（100000=100%），图像宽高比不同必须检查主体。`keep`须能证明对应对象仍适用于本期；`remove`须给出reason并产生材料缺口。模板中的背景、品牌占位元素和装饰边条会保留，换用个人模板时还须确认这些元素是否可公开使用。

参考7页的环形图需要chart：`{"action":"replace","name":"数量","categories":[...],"values":[...],"fact_ids":[...]}`。每个值绑定事实，最多4类；更多类别需归并或拆页，不自动创造新配色。keep同样需要来源。构建修复空外链并写入真实Data工作簿，缓存、公式和单元格一并核对。

需要超出原环形图/柱形拓扑的证据图时，可在**无现有原生图表的参考页**加入 `data_chart`，支持 `trend`、`comparison`、`ranking`、`composition`、`target_progress`。其 `series[].fact_ids` 直接指向已核验指标；输出是可编辑的 PowerPoint 原生图表及内嵌 Data 工作簿。必须指定图表区域、显式清除该区域全部旧对象、在可见文字槽位写来源、并给出解释。完整页面语法、适用条件及审核项见 [证据图表](data-charts.md)。旧的 `chart` 合同保持有效。

Word 内嵌大量项目图片时，先用 `media-inventory` 建每次图片出现的账本，再审阅项目目录与匹配候选。顶层加入 `media_inventory`、`project_matches`、`asset_dispositions`；照片 `pictures` 可附 `asset_id`、`project_id` 和对歧义图片的 `confirmation_note`。系统验证图像字节、DOCX锚点、项目确认及全部图片去向，输出 manifest 保留每张 PPT 照片的溯源链。单、双、三图原生布局与裁切建议见 [富图 Word 合同](media-assets.md)；识别结果必须看成片后核对。

参考5页支持`ranking:{"rows":[{"label":"类别名","fact":"数值事实ID","sources":["来源:段落"]}]}`，1–17项按降序；按同一最大值比例计算柱宽，删除空余行。颜色沿用原行的排名色阶，分组阈值变化时须明确调整内容分组并复核说明。表旁总数和说明仍由slots填写，不能保持演示统计。

月度柱图、参考9页双时期比较、侧边形状柱图使用`bars`数组：每条指定`shape`、`value_slot`、`fact`、`axis:width|height`、`maximum`（绘图区上限）、`extent_pt`（上限对应长度）。数值槽位还需绑定同一个facts ID。宽度沿左端缩放，高度沿底端缩放，保持原形状和色彩；0值绘制0.01pt细线。所有轴标签、旧时期数据和结论同步更新，不能只改数值文本而保留旧柱长。

参考4、5、7、9、10页的数值标签与柱形已建立确定性映射。改变数值后缺少对应`bars`更新将报错；删除数值标签时必须移除相应柱形。参考5页使用`ranking`时由排名组件统一重算。双时期比较的一个标签对应两根柱，必须同时绑定两期事实。轴线、刻度、均值线和结论的语义一致性仍须审阅者检查。

减少流程步骤时，同时移除不用的文字槽位和空形状。确有必要重新分布原生对象时，使用`geometry:{"对象ID":{"bounds_pt":[x,y,w,h],"reason":"具体重排原因"}}`，坐标单位pt，必须在960×540画布内；不允许借此改字体。超过已有版式容量优先拆页；重排后逐页审查。

指标槽位也可按内容数量删减和重排。只保留与当前结论有关的数字及其图标、单位、说明；对多余组的文字使用 `remove`，装饰图标使用 `remove_shapes`，随后对其余原生对象写有理由的 `geometry`。不要用同义重复、不同范围的弱指标或无关获奖分数填满固定格数。重排保留原字号和混合格式，并需在 PowerPoint 实际渲染后检查。

空白非图片叶形状可通过`remove_shapes:["对象ID"]`移除无用的线条、背景或流程箭头（仍保留其余原生对象）。文字槽位的移除使用slots中的remove以保留原因；不可用remove_shapes绕过文字/照片来源检查。`data_chart.replace_shapes` 只为同一图表区域的原生旧形状提供清理路径，旧文字仍需在 `slots` 显式 `remove`，并接受区域遮挡检查。

## 状态与错误

命令输出JSON，错误退出码2。build生成独立版本，状态`built_unverified`；verify结构与实际渲染通过后为`rendered_pending_visual_review`；finalize还要求逐页审查、无阻断缺口及指纹一致，最终为`verified`。

公开样例使用合成数据，只用于测试接口。不得把 `examples` 中的来源声明直接用于真实汇报。脚本不会自动写审美批准，也不自动忽略警告。
