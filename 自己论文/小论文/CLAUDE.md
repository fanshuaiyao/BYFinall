# CLAUDE.md

## 当前项目目标

将本目录下的毕业论文 PDF 的**第四章**改写为一篇面向 **Bioinformatics** 期刊投稿的**英文小论文**。

## 协作流程

1. 用户提供毕业论文 PDF，由 Claude 读取并提取第四章内容。
2. Claude 不做简单直译，而是按照 **Bioinformatics** 期刊论文风格对第四章内容进行英文化重组与改写。
3. 目标产出应符合常规英文学术论文结构，通常包括：
   - Title
   - Abstract
   - Introduction
   - Materials and Methods / Methods
   - Results
   - Discussion
   - Conclusion
   - References
4. 用户会先从 Overleaf 找到目标期刊模板，并将模板的 `.tex` 文件提供给 Claude。
5. Claude 直接在本地编辑该 `.tex` 文件，与用户共同迭代修改。
6. Claude 完成编辑后，用户再将生成或修改后的 TeX 内容复制回 Overleaf。

## 写作原则

- 不做逐段直译，优先做论文化改写。
- 以 **Bioinformatics** 期刊风格为目标，而不是毕业论文风格。
- 优先保留第四章中的核心：研究问题、方法、实验设计、结果、结论。
- 必要时可重组章节顺序，使内容更符合期刊投稿表达。
- 输出语言默认使用英文；与用户沟通默认使用中文。

## 后续工作默认顺序

1. 先读取并理解期刊模板 `.tex`
2. 再读取毕业论文 PDF 的第四章
3. 先整理英文论文大纲
4. 再将内容写入并迭代修改 `.tex`
