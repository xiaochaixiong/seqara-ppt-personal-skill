# 意群转折与原生过渡页

此流程用于讲述逻辑发生转换时，而不是按固定页数增加深蓝页。先用 `storyboard` 定义意群与每页时长，再检查相邻意群的关系。`suggest-transitions` 只列出边界和审阅问题，默认建议微转场；程序不能仅凭页数或版式重复判断是否需要独立过渡页。

对每个需要处理的边界写明听众此刻的问题、前后论点关系、桥接句及选择理由。三级选择：

- `micro`：前页讲述备注加桥接语，后页标题完成承接，不加页。适用于顺畅的证据递进、同一主题继续展开。
- `content`：桥接本身是一个新判断、选择或比较，用参考页 2、18、23、26、29 的原生内容构图表达。需要完整填写该页所有历史文字槽位及来源；仅有“现在进入下一部分”不够。
- `pause`：重大论点/章节边界且删去会失去必要判断，用参考页 17 或 25。必须具体说明独立成页的必要性；系统重填标题、说明、证据行、PART 标签和大号章节号，或按明确决定移除章节号。

## 计划合同

顶层可选 `transitions`。旧计划不含该字段仍可构建。每条边必须连接 `storyboard.beats` 中相邻意群，并给出 `id`、`from_beat`、`to_beat`、`relation`、`audience_question`、`bridge`、`rationale`、`tier`。关系代码为 `question_answer`、`evidence_judgment`、`diagnosis_action`、`case_rule`、`known_unknown`、`speaker_shift`、`section_shift`、`other`。代码只方便检索，实际关系应由桥接句说明。

微转场示例：

```json
{"id":"t01","from_beat":"B01","to_beat":"B02","relation":"evidence_judgment","audience_question":"这些数字说明什么？","bridge":"数据证明了覆盖，接下来要判断质量。","rationale":"前页结尾和后页标题已经足以承接。","tier":"micro"}
```

独立页另需 `if_removed_loss`、`seconds`、`page_id` 和 `significance`。`content` 需 `significance:"new_judgment"` 或 `"major_pivot"`，且在 `page` 中提供已写好的原生页面对象。`pause` 需 `significance:"major_pivot"`、`template_slide:17|25`、`section` 与 `content`：

```json
{
  "id":"t02","from_beat":"B02","to_beat":"B03",
  "relation":"diagnosis_action","audience_question":"既然问题明确，先做什么？",
  "bridge":"先解决执行条件，再扩展项目。",
  "rationale":"这里由诊断转入实施原则。","tier":"pause",
  "significance":"major_pivot",
  "if_removed_loss":"会丢失从诊断推导出的实施优先级。",
  "seconds":12,"page_id":"implementation-turn","template_slide":17,
  "section":{"number":2,"label":"实施原则"},
  "content":{"headline":"先解决执行条件，\n再扩展项目",
    "explanation":"演示证据支持先明确职责与范围。","evidence_line":"依据：合成演示记录",
    "kind":"inference","sources":["manuscript:p012"]}
}
```

`section.number` 可省略；此时原生大号幽灵章节号会移除，标签只写 `section.label`。提供数字时，参考 17 页的幽灵号在槽位 3，25 页在槽位 4，统一写两位号；槽位 5 自动写 `PART NN · 标签`。这两组值会相互校验，避免残留模板演示标签。标题、说明、证据行必须绑定本期来源；无独立证据行可以不写，原槽位被明确移除。

参考 17/25 页的说明文字包含混合格式。可在 `content` 加 `explanation_highlight`，指定说明文字中恰好出现一次的短语，复用原页加粗片段；其余文字沿用原常规片段。两行标题的文本框高度必须在 PowerPoint 实际渲染后检查；出现超框时先调整文本框或分屏，不通过缩小字号硬塞。

`apply-transitions` 对微转场只写前页讲述桥接；对独立页把新增秒数从下一意群第一页的时长中分配出来，整套讲述时间不被悄悄拉长。若原来的页面承载了相邻两个意群，须先拆页再插入独立页。再次运行不会重复插页。生成后的页码由构建器统一更新，后续页内引用仍用 `target_page` 绑定稳定页面 ID。

验证只能检查结构、时间分摊、来源与章节标签。实际演讲是否更清楚，仍需对有/无过渡页版本进行完整 PowerPoint 渲染、串读标题、核对重复论点并试讲。若增加页只重复下一页标题或抬高总时长，就保留微转场。
