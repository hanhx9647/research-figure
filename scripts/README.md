# scripts

| 脚本 | 作用 |
|---|---|
| `figure_kit.py` | 公共原语：形状 / 中文字体 / 带箭头连接符（避开两个已知 python-pptx 坑） |
| `validate_pptx.py` | 结构校验：箭头、中文字体、文字溢出、模块重叠、整页截图 |
| `convert_image_to_pptx.py` | 图片 → 可编辑 PPTX（区域声明式驱动 + 编辑性报告；`--auto` 生成待人工分类的草稿） |
| `render_preview.py` | PNG/PDF 预览：PowerPoint COM → LibreOffice → 明确提示无渲染器 |
