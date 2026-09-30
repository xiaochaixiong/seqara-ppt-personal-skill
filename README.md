# Seqara PPT Personal Skill

![Seqara PPT Personal 功能示意：数据、叙事与项目图片](docs/promo/hero.png)

*上图为功能示意，不是实际生成的汇报截图。*

把口述稿、数据表和 Word 项目图片整理为一套**有讲述节奏、来源可查、原生可编辑**的 PowerPoint 汇报。Skill 复用固定 38 页演示模板中的页面和对象，生成 PPTX、逐页备注及验收材料。

> **公开版提供的是中性、合成的演示模板。**原个人参考稿、真实照片、业务数据和口述稿不随仓库发布。当前实现绑定这份 38 页演示模板的对象结构，不能将任意 PPTX 放进 `assets/reference.pptx` 就直接使用；更换模板须重做对象映射和全页回归验收。

## 它解决什么问题

| 需求 | 做法 | 需要人判断的部分 |
| --- | --- | --- |
| 演讲稿逻辑与节奏 | 先写听众问题、结论和证据，再选页面；转场可放在备注、内容页或停顿页 | 哪个判断值得独立成页 |
| 数据与图表 | 建立 XLSX/CSV/Word 表格来源账本，记录口径和计算规则；生成五类带内嵌工作簿的原生图表 | 统计口径、比较是否成立、图上结论是否准确 |
| Word 多项目图片 | 逐次提取图片与上下文，提出项目候选及一到三图排版 | 照片是否真实属于该项目，主体裁切是否合适 |
| 可编辑与验收 | 复用 PPTX 文字、形状、图表和图像对象；检查实际文件并用 PowerPoint 渲染 | 字体、遮挡、整套叙事及最终视觉质量 |

图表类型包括趋势、比较、排名、构成和目标进度。`draft` 命令只生成待填写计划；模型需要结合材料补齐内容，脚本执行确定性的构建与检查。任何新报告都要重新核对来源和成片。

## 环境

- Windows、Python 3.10 或以上、本机 Microsoft PowerPoint。
- MiSans Normal。仓库不分发字体，也不会静默替换缺失字体。
- Python 依赖见 `requirements.txt`。Seqara 与 Codex 共用同一套 CLI；不需要另配模型 API 密钥。

```powershell
python -m pip install -r requirements.txt
python -X utf8 scripts/personal_ppt.py doctor
```

PowerPoint 用于 PDF/PNG 实际渲染及视觉验收。WPS 后续编辑兼容性没有作为独立渲染环境验收。

## 直接运行合成示例

仓库内的 `examples/synthetic-data/` 提供完整、虚构的 6 个月趋势与同期比较计划。安装依赖并通过 `doctor` 后，可直接构建，无须先写自己的稿件：

```powershell
python -X utf8 scripts/personal_ppt.py check-plan examples/synthetic-data/plan.json
python -X utf8 scripts/personal_ppt.py build examples/synthetic-data/plan.json --out demo-output
python -X utf8 scripts/personal_ppt.py verify "build 返回的版本目录"
```

查看已在本机 PowerPoint 实际渲染的[趋势页](docs/demo/slide-01.png)、[对比页](docs/demo/slide-02.png)、[可编辑 PPTX](docs/demo/presentation.pptx)、[PDF](docs/demo/presentation.pdf)和[逐页讲述备注](docs/demo/speaking-notes.md)。这些数字、标题和图片占位均是合成演示，不对应真实机构。

## 安装与试用

将仓库根目录作为 `seqara-ppt-personal-skill` 文件夹导入 Seqara 的 Skill 管理器，或复制到 `$CODEX_HOME/skills/`。也可用 `scripts/install_local.py --target codex` 安装到默认 Codex 目录；安装器不会覆盖同名目录。新开会话后明确调用 `$seqara-ppt-personal-skill`。

先用自己的 DOCX/TXT/MD 稿件在演示模板上试跑。以下示例中的项目目录由你创建；`1,4,11` 只示意封面、数据页和照片页，应按实际内容重新选择。页面目录见 `references/layout-catalog.md`。

```powershell
python -X utf8 scripts/personal_ppt.py extract "稿件.docx" --id manuscript --role primary --out "项目/manuscript.source.json"
python -X utf8 scripts/personal_ppt.py draft --source "项目/manuscript.source.json" --pages 1,4,11 --out "项目/plan.json"
```

接着依据 `references/plan-contract.md` 填写页面目的、槽位、来源、讲述备注和原文覆盖。`draft` 不会自动理解稿件，空计划无法直接生成合格成片。数据表需要 `extract-data`、`check-data`；Word 多图可用 `media-inventory`、`media-match`；转场用 `suggest-transitions`、`apply-transitions`。这些高级命令的字段见对应 `references/` 文档。

```powershell
python -X utf8 scripts/personal_ppt.py check-plan "项目/plan.json"
python -X utf8 scripts/personal_ppt.py build "项目/plan.json" --out "项目/output"
python -X utf8 scripts/personal_ppt.py verify "build 返回的版本目录"
python -X utf8 scripts/personal_ppt.py review-init "版本目录"
```

逐页查看实际输出的 PNG 与 PPTX，核对文字、图片、数字、图表、裁切和遮挡，再检查整套标题串与讲述节奏。填写真实的 `visual-review.json`，然后运行 `finalize`。不能因为脚本返回成功就批量批准审阅项。

```powershell
python -X utf8 scripts/personal_ppt.py finalize "版本目录"
```

完整流程及规则见 [`SKILL.md`](SKILL.md)；最短操作路径见 [`QUICKSTART.md`](QUICKSTART.md)。

## 验证与边界

公开版在 Windows、Python 3.14、MiSans Normal、Microsoft PowerPoint 16.0 环境下通过 **28 项自动测试**；中性模板全部 38 页的对象映射已复核。上述 2 页合成示例通过来源和图表数据检查，生成的 PPTX 经 PowerPoint 实际渲染，逐页审查后状态为 `verified`，0 错误、0 警告。记录见[公开样例验收](docs/demo/validation.md)。每份新报告仍须重新验证，不能沿用样例验收。

运行自动测试：

```powershell
python -B -X utf8 -m unittest discover -s tests -v
```

合成图表数据不是现实业务数据；中性照片占位不能证明真实项目图片识别效果。素材匹配与来源检查不能代替事实核实，结构检查也不能证明图表结论或设计质量。缺字体、缺关键材料或无法渲染时，输出只能标为未完成验收的草稿。

## 主要目录

| 路径 | 用途 |
| --- | --- |
| `SKILL.md` | 模型使用时的决策规则与交付边界 |
| `references/` | 叙事、页面、数据、图表、转场、富图接口 |
| `scripts/personal_ppt.py` | 双端共用 CLI |
| `scripts/pipeline.py`、`scripts/native_pptx.py` | 内容计划与原生对象复用 |
| `scripts/data_sources.py`、`scripts/native_data_charts.py` | 来源账本与原生图表 |
| `scripts/transitions.py`、`scripts/media_assets.py` | 转场与 Word 图片处理 |
| `scripts/quality.py` | 结构、渲染和交付验收 |
| `assets/` | 中性 38 页演示模板与对应对象映射 |

私有参考稿的内容不应因页面复用而进入新报告。演示模板只用于测试方法，实际汇报必须以本期材料重新建立事实、图片和图表来源。

代码采用 [MIT 许可](LICENSE)；演示模板、图片和预览的单独使用范围见 [ASSETS.md](ASSETS.md)。

需要分享项目时，可直接使用[推广图与图文文案](PROMOTION_COPY.md)。
