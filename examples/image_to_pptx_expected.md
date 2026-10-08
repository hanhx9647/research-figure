# Expected output of `image_to_pptx.py`

对照说明：跑完示例后，PPTX 里**什么可以编辑、什么不可以**。

## 原图 vs 转换结果

| | 文件 |
|---|---|
| 原图（位图输入） | `examples/image_to_pptx_input.png` |
| 转换产物 | `output/image_to_pptx_converted.pptx` |
| 编辑性报告 | `output/image_to_pptx_converted_editability.md` |

## 可编辑元素（原生重建，10 个区域中 9 个）

| 元素 | 在 PPTX 中的形态 | 可执行的操作 |
|---|---|---|
| 标题「肿瘤影像分析流程」 | 独立文本框 | 改字、换字体、移动 |
| 4 个流程模块 | 圆角矩形原生形状 + 独立文字 | 改色、改字、缩放、复制 |
| 3 条箭头 | 原生连接符（带真实箭头） | 拖动端点、改线型 |
| 图例 | 原生形状 | 编辑文字 |

## 嵌入图片元素（1 个区域）

| 元素 | 原因 |
|---|---|
| 左下角示例影像（灰度图） | 属于图像数据；重绘为形状会破坏科学保真度，按规范保留为嵌入图片 |

## 预期编辑性评分

`rebuild = 9 / (9 + 1) = 90%` → **高度可编辑**

## 本示例刻意演示的三件事

1. **区域划分是声明式的**：`image_to_pptx_regions.yaml` 里每个区域标了 `rebuild` / `embed`，转换器不做隐式猜测 —— 这是「不保证 100% 自动还原」原则的落地形式。
2. **嵌入 ≠ 偷懒**：影像缩略图保留为图片是规范要求（见 `references/image-to-pptx.md` 第 2 节），而不是转换失败。
3. **报告随文件走**：每次转换都产出编辑性报告，把「哪些能改」写清楚，收图的人不用猜。

## 想跑自己的图？

```powershell
# 第一步：自动检测生成区域草稿（全部标记 uncertain，不会直接出成品）
python scripts/convert_image_to_pptx.py --image path/to/your.png --auto

# 第二步：人工把草稿里的区域分类、填文字

# 第三步：正式转换
python scripts/convert_image_to_pptx.py --image path/to/your.png --regions path/to/your.regions.yaml
```
