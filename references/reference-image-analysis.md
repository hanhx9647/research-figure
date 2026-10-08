# Reference image analysis

> **核心原则：参考图只决定视觉风格，不能自动证明科研逻辑。**
> 从参考图里提取的是版式、配色、字体、线条；科学内容一律来自用户确认过的 `figure_contract.yaml`。
> 不得因为参考图里画了某个模块，就把它加进用户的流程图。

## 参考图存放位置

把风格参考图放在 `references/reference-styles/`（一张图一个文件，文件名即风格名）：

```text
references/reference-styles/
├── nature-clean.png
├── cell-method.png
└── README.md          # 逐张记录下面 10 项分析结果
```

每加一张参考图，就在 `references/reference-styles/README.md` 追加一条分析记录，供后续复用。
仓库 `templates/theme.yaml` 是默认主题；参考图分析结论可覆盖它，但必须记录覆盖理由。

## 十项分析维度

对每张参考图，逐项记录：

| # | 维度 | 要回答的问题 |
|---|---|---|
| 1 | 版式方向 | 横向 / 纵向 / 网格 / 中心辐射？阅读顺序如何？ |
| 2 | 模块数量 | 主模块几个、子模块几个？深度（层级）多少？ |
| 3 | 分区方式 | 有无分组框、色带、泳道、面板（a/b/c）？留白比例多少？ |
| 4 | 标题样式 | 有无总标题？字号层级、对齐、是否加粗、是否带编号？ |
| 5 | 配色提取 | 取 2–4 个主色 + 中性色，给出 HEX；是否存在语义配色（输入/处理/输出各一色）？ |
| 6 | 字体层级 | 拉丁字体估计、字号梯度（标题/模块/注释）、字重、是否全大写、中文回退字体 |
| 7 | 线条和箭头样式 | 粗细、虚实、箭头形状（三角/开口/V 形）、拐角方式（直角/圆角/曲线） |
| 8 | 是否存在图例 | 有无图例？位置？是否可复用为语义说明？ |
| 9 | 是否适合转换为 PPT | 结构是否规整、是否依赖位图、文字量是否过载、纵横比是否匹配 16:9 |
| 10 | 科学内容 vs 视觉风格 | 哪些元素是**科学内容**（模块语义、因果关系）？哪些只是**视觉风格**（配色、圆角、阴影）？ |

第 10 项是分界线，必须显式写出来。可提取的只有右列；左列只能来自用户确认。

## 分析输出格式

```yaml
reference: nature-clean.png
orientation: left-to-right
module_count: { main: 5, sub: 8, depth: 2 }
partition: { group_boxes: true, lanes: false, panels: false }
title: { present: true, size_pt: 26, weight: bold, align: left }
palette: { input: "DDEBF7", process: "E2F0D9", decision: "FCE4D6", output: "FFE699", text: "1F2937" }
fonts: { latin: "Aptos", cjk: "Microsoft YaHei", sizes: [26, 18, 12] }
connectors: { width_pt: 1.5, dash: solid, arrowhead: triangle, corner: right-angle }
legend: { present: true, position: bottom-right }
ppt_ready: true
scientific_content: []      # 空 —— 不参考
visual_style_only: true     # 本图仅用于风格
```

## 冲突处理

- 参考图风格与用户明确要求冲突 → **以用户要求为准**。
- 参考图风格与 `theme.yaml` 冲突 → 以参考图为准，但记录覆盖理由。
- 参考图不利于可访问性（如仅靠颜色区分语义、对比度不足）→ 保留配色，增加形状或文字冗余编码，不照搬。

## 禁止事项

- 禁止从参考图搬运科学模块、因果箭头、阶段名称
- 禁止把参考图中的示例占位文字当成领域术语
- 禁止引入参考图中的复杂效果（渐变网格、艺术字、位图纹理），它们是 Google Slides 兼容性地雷
