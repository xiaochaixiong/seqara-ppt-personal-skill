# 公开版演示验收（2026-09-30）

此记录仅适用于随仓库发布的中性模板与合成双图示例。源讲稿、表格及图中数字均为虚构数据。

| 项目 | 结果 |
|---|---|
| 环境 | Windows、Python 3.14、Microsoft PowerPoint 16.0、MiSans Normal |
| 模板 | 38 页、1,270 个原生对象；页面对象 ID、类型、位置及文字 Run 格式完成对照，0 处差异 |
| 模板渲染 | 38/38 页由 PowerPoint 导出 PNG/PDF；逐张检查 7 张联系表 |
| 公开单元测试 | 28/28 通过，包括 38 页对象映射、合成计划构建、图表工作簿与缓存、数据、转场及图片账本 |
| 合成示例 | 2 页；`check-plan`、构建、结构/数据检查、PowerPoint 实际渲染与逐页视觉审查完成；0 错误、0 警告；交付状态 `verified` |
| 可编辑性 | 页面文字、形状及两张原生图表保留；趋势图与双序列比较图各有内嵌工作簿，缓存值与工作簿一致 |

发布模板 SHA256：`4afd338cf1206489ec22527d02bfacc132cc7f8609adf787a3dbb66fa9fe4a13`。

[演示 PPTX](presentation.pptx) SHA256：`f03c846ce0e9cf8f843156a091d0301937a883ae1e751f7c94a437386c73dd75`；[PDF](presentation.pdf) SHA256：`bcb5514176d3a7212d84c7c7be9fbf15c3a5929b05b2c7d8c65d034178d259d0`。可查看[趋势页](slide-01.png)、[同期比较页](slide-02.png)及[讲述备注](speaking-notes.md)。模板联系表：[1](../template-preview/contact-01.jpg) · [2](../template-preview/contact-02.jpg) · [3](../template-preview/contact-03.jpg) · [4](../template-preview/contact-04.jpg) · [5](../template-preview/contact-05.jpg) · [6](../template-preview/contact-06.jpg) · [7](../template-preview/contact-07.jpg)。

模板中的 23 张原媒体已全部替换为自制几何占位图，历史备注、嵌入字体及外链均已移除。视觉设计仍沿用用户提供参考稿的页面结构。模板原有的部分文本框在 PowerPoint 测量中出现边界超框数值；与原稿对照无新增超框，联系表审阅未见截断。此演示未验证真实项目照片的归属、构图质量、任何现实业务事实，也不证明另一份 PPTX 可以直接替换本模板。
