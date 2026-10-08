# ResearchFigureGPT 中文说明

**把科研方法、手绘草图和科学逻辑转换成可编辑的 PowerPoint 科研图。**

ResearchFigureGPT 是面向科研人员、研究生和工程师的开源 AI Skill，用于生成真正可继续修改的**科研流程图、实验方案图、技术路线图和系统架构图**。它可以从论文方法段落、科研描述、手绘草图或参考图片生成结构化 `.pptx`，输出原生形状、可编辑文字和真实连接符，而不是一张只能查看的图片。

> **ResearchFigureGPT 是 [Research Artifact Lab](https://github.com/hanhx9647/research-artifact-skills) 的首个独立 Skill。**

## 为什么使用 ResearchFigureGPT？

很多 AI 绘图工具只能生成 PNG，研究人员还需要重新绘制。ResearchFigureGPT 围绕科研交付物设计：

- **默认面向可编辑：** 文本、方框、箭头和布局元素使用原生 PowerPoint 对象。
- **先确认科学逻辑：** 先提取并确认研究目标、阶段、输入输出和关系，再生成最终图形。
- **支持图片转 PPTX：** 可以将指定区域重建为可编辑对象，同时明确记录仍保留为图片的区域。
- **区分事实与推断：** 区分用户确认的内容、假设、歧义和视觉参考。
- **支持中英文科研内容：** 支持中文标签、科学符号、多分支流程和反馈回路。
- **可复现：** 示例包含 Python 源码、主题配置、生成文件和验证报告。

ResearchFigureGPT 不宣称所有对象都能编辑。照片、显微图和复杂嵌入素材可能仍是图片，报告会明确列出这些对象。

## 安装

直接安装独立 Skill：

```bash
npx skills add hanhx9647/research-figure
```

也可以从 Skill 集合安装：

```bash
npx skills add hanhx9647/research-artifact-skills --skill research-figure
```

本地运行生成脚本需要 Python 3.10 或更高版本：

```bash
python -m pip install -r requirements.txt
```

生成 `.pptx` 不要求安装 PowerPoint。生成 PNG/PDF 预览时，需要本机有 PowerPoint 或 LibreOffice。

## 工作流程

```text
论文方法 / 科研描述 / 手绘草图 / 参考图片
                    ↓
              科学逻辑提取
                    ↓
              确认草案与审批
                    ↓
              视觉架构与对象规划
                    ↓
              原生 PowerPoint 生成
                    ↓
       PPTX + 预览图 + 源码 + 验证报告
```

对于论文和科研描述，Skill 会先提取目标、阶段、输入、输出、依赖关系、反馈回路和不确定关系，获得确认后才生成最终图形。对于图片转换，会将区域分为“原生重建”和“嵌入图片”，并记录可编辑性结果。

## 快速开始

生成原生科研流程图并检查 PPTX 结构：

```bash
python examples/simple_workflow.py
python scripts/validate_pptx.py output/simple_workflow.pptx
```

将参考图转换为可编辑 PPTX：

```bash
python examples/image_to_pptx.py
```

运行完整测试：

```bash
python tests/run_tests.py
```

## 适用场景

- 科研流程图和论文方法图
- 实验流程与技术路线图
- 算法流程和系统架构图
- 将论文方法段落转换为流程图
- 将手绘草图重建为可编辑 PPTX
- 中英文科研图和复杂反馈回路
- 同时包含可编辑结构与图片素材的科研图

## 目录结构

```text
SKILL.md                 # Agent Skill 入口
references/              # 工作流、审批、图片转换和质量规范
scripts/                 # 生成、转换、渲染和验证工具
templates/theme.yaml     # 可复用主题配置
examples/                # 可运行示例和预期成果
tests/                   # 回归测试和静默失败测试
```

## 验证与限制

验证器会检查 PPTX 结构、原生形状、文字对象、连接符、箭头和图片区域。项目会记录尚未执行的检查，例如人工视觉审查，以及当前环境无法进行的兼容性测试。科研结论和科学逻辑仍然需要研究人员审阅。

## 文档

- [Skill 规范](SKILL.md)
- [工作流契约](references/workflow-contract.md)
- [确认模板](references/confirmation-template.md)
- [图片转 PPTX 指南](references/image-to-pptx.md)
- [验证规则](references/validation-rules.md)
- [English README](README.md)
- [Research Artifact Lab 集合](https://github.com/hanhx9647/research-artifact-skills)

## 许可证

本项目采用自定义的非商业署名许可证。允许个人、学术、教育及其他非商业用途，但必须署名原作者 **左手骑车**。商业使用及修改版本再分发需要书面许可。详见 [LICENSE](LICENSE)。
