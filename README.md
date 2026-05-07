# 上海大学本科毕业设计论文

本仓库是我的本科毕业设计论文项目，论文主题为 **多模态混合专家模型的推理优化**，主要研究 Vision-Language Mixture-of-Experts（VL-MoE）模型在推理阶段的性能瓶颈与优化方法。

论文围绕多模态输入负载、Visual Encoder 前端开销和 MoE 专家调度代价展开，重点包括：

- 基于真实 VL-MoE 推理链路的负载特征分析；
- 面向 Visual Encoder 的语义感知视觉 Token 压缩与异步流水优化；
- 面向 CPU-GPU expert offload 场景的全局路由预测器；
- ExpertFlow 风格、DALI 风格和 Hybrid 融合专家调度策略；
- 任务效果保持性、TTFT、预测准确率和真实 expert offload 微基准实验。

## 仓库结构

```text
.
├── main.tex                 # 论文主入口
├── contents/                # 各章节正文、摘要、结论、附录
├── figures/                 # 论文图片与实验图表
├── reference/refs.bib       # 参考文献
├── fonts/                   # 本地编译所需中文字体
├── scripts/build.sh         # 编译脚本
├── build/main.pdf           # 已编译生成的论文 PDF
├── shuthesis.cls            # 上海大学本科论文模板类文件
├── shuthesis.cfg
└── shuthesis.sty
```

## 章节内容

- 第一章：绪论，介绍研究背景、意义、相关工作和本文主要工作；
- 第二章：VL-MoE 推理负载分析与问题建模；
- 第三章：面向 Visual Encoder 的多模态感知优化方法；
- 第四章：基于全局预测的专家提前调度方法；
- 第五章：系统实现与实验评估；
- 附录：第三章与第四章关键实现代码。

## 本地编译

本项目使用 XeLaTeX 编译。推荐直接运行：

```bash
bash scripts/build.sh xebib
```

也可以使用：

```bash
xelatex -interaction=nonstopmode -halt-on-error -file-line-error -output-directory=build main.tex
```

编译后的 PDF 位于：

```text
build/main.pdf
```

## 说明

本论文模板基于上海大学本科毕业论文 LaTeX 模板修改整理，感谢 SHU-Bachelor-Thesis-OSC 及相关开源模板作者的工作。

本仓库主要用于毕业设计论文撰写、版本管理和归档。
