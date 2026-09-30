# 富图 Word 的项目照片

当口述稿 DOCX 有项目照片时，先读取 `scripts/media_assets.py` 产生的媒体账本，再决定照片页。图片是证据素材；alt 文本、文件名、视觉猜测和参考稿照片均不能单独证明它属于某项目。

## 提取与项目匹配

```powershell
python -X utf8 "<skill>/scripts/personal_ppt.py" media-inventory "<口述稿.docx>" --id speech --out "<项目>/media.json"
python -X utf8 "<skill>/scripts/personal_ppt.py" media-project-candidates "<项目>/media.json" --out "<项目>/projects.json"
python -X utf8 "<skill>/scripts/personal_ppt.py" media-match "<项目>/media.json" --projects "<项目>/projects.json" --out "<项目>/project-matches.json"
```

`media.json`记录每次图片出现的 `asset_id`、`occurrence_id`、源DOCX SHA、媒体SHA、图片段落锚点（即使该段落没有文字）、标题路径、图注候选、前后段落、表格单元格与行、alt、像素尺寸及提取文件。相同字节重复出现会共用提取文件，但保留不同 occurrence。读取内嵌、浮动 `wp:anchor` 和旧式 VML；JPEG/PNG可用于原生PPT图片替换，外链、缺失和不支持格式须列为缺口。账本和导出的媒体需随生成计划保存，修改DOCX后必须重抽取。

`project-candidates` 仅把图片附近章节标题列为待核对实体，默认 `reviewed:false`。运行 `match` 前，应依据口述稿或其它本期证据删除泛化标题、核对项目正式名称并添加别名；核对完成后才把 `reviewed` 改为 `true`。项目目录例如：

```json
{
  "source_doc_sha256": "media.json中的document_sha256",
  "reviewed": true,
  "projects": [
    {"id": "project-a", "name": "项目正式名称", "aliases": ["简称"]}
  ]
}
```

匹配报告列出各张图的候选、依据和 `auto`/`review`。明确图注、同一段/表格单元格或唯一章节标题可以形成自动建议；多个项目上下文冲突、同一媒体跨项目复用、只有相邻文字/文件名或图片不支持时进入 `project-matches.review.html` 缩略图清单。`auto` 表示文档上下文足够明确，仍须人工看最终PPT的图像主体与图注；`review` 不得悄悄作为事实照片使用。

## 绑定到生成计划

已有plan无媒体字段时完全兼容。新富图plan增加 `media_inventory`、`project_matches`（相对plan文件）和每张图的处置。项目匹配报告中的目录与账本均由输入SHA绑定。示例：

```json
{
  "media_inventory": "media.json",
  "project_matches": "project-matches.json",
  "asset_dispositions": {
    "speech:document-p0012-i01": {"status": "used"},
    "speech:document-p0014-i01": {"status": "omit", "reason": "同项目近似重复照片"},
    "speech:document-p0015-i01": {"status": "pending", "reason": "照片归属需确认"}
  },
  "pages": [{
    "id": "project-a-evidence",
    "pictures": {"34": {
      "action": "replace", "path": "media-media/<SHA前20位>.jpg",
      "sha256": "原媒体完整SHA", "role": "evidence", "generated": false,
      "sources": ["speech:p011"], "asset_id": "speech:document-p0012-i01",
      "project_id": "project-a", "crop": {"l": 0, "t": 0, "r": 0, "b": 0}
    }}
  }]
}
```

所有图片 occurrence 必须有 `used`、`duplicate`、`omit` 或 `pending` 处置；后三者需要理由，`pending` 是阻断交付的缺口。`review` 图片若经过人工核对，可在对应 `pictures` 中写 `confirmation_note`，并保留匹配报告、图注与确认依据。`validate_asset_bindings` 校验文件、媒体SHA、项目目录、自动/人工归属和全部处置；构建 manifest 需记录 `输出页/shape → asset_id → occurrence → 项目 → DOCX锚点 → 媒体SHA`。

```powershell
python -X utf8 "<skill>/scripts/personal_ppt.py" media-check "<项目>/plan.json"
```

## 一、二、三图排版

```powershell
python -X utf8 "<skill>/scripts/personal_ppt.py" media-layout-projects "<项目>/project-matches.json" --out "<项目>/project-photo-pages.json"
python -X utf8 "<skill>/scripts/personal_ppt.py" media-layout "<项目>/media.json" --assets "speech:document-p0012-i01,speech:document-p0013-i01" --out "<项目>/photo-layout.json"
```

`layout-projects` 将唯一明确归属的图片按项目分组，并按DOCX出现顺序生成1–3图页面候选；同一项目的重复媒体列为重复 occurrence，冲突和低置信度图片留在 review 列表，不自动进页。页次仍需按口述稿的论证关系调整，不能把文档顺序直接当讲述顺序。

单/双图基于参考11页，三图基于参考13页。单图建议将原 `34` 形状扩展到 `[30,150,876,240]`，移除照片 `37` 与图注槽位 `13`；它仍需填充其余文字槽位。双图使用 `34/37`，三图使用 `38/41/44`。脚本只按原图尺寸和框宽高比提出居中裁切，**不识别主体**；先按项目目标、现场、成果组织图片，再逐张审阅主体、裁切和图注。超过三图时按讲述关系分多页，不能因数量相同把不同项目混装。

合成DOCX回归可覆盖无文字图片段、表格、浮动图、VML、重复媒体和项目冲突；真实图像的项目身份与裁切质量仍须在使用者自己的材料上逐张核验。
