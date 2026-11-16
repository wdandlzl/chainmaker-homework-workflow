# 长安链训练营作业工作流

用于在 Codex 中完成单份长安链训练营实验作业，从登记题目和验收条件，到生成报告、导出 PDF 和逐项核对证据。

## 使用

1. 将仓库内容放到实验项目根目录，让 `AGENTS.md` 和 `.agents/skills/` 位于根目录。项目规则默认使用 `D:\长安链作业`，采用其他路径时更新本地路径约定。
2. 准备你自己的实验手册和原始 Word 报告模板，按项目规则放置。
3. 在 Python 虚拟环境中运行 `python -m pip install pymupdf python-docx`。导出 PDF 的 PowerShell 脚本使用 Windows 上已安装的 Microsoft Word。
4. 在 Codex 中打开项目根目录，指定要完成的一份作业，按 `AGENTS.md` 和项目 Skill 执行。
5. 首次运行时生成并核对 `.workflow/template_rules.json`。将本地核实的技术问题记录到 `.workflow/corrections.md`。

## 文件

- `AGENTS.md` 规定课程资料范围、实验验证、纠错和交付要求。
- `.workflow/TASK_TEMPLATE.md` 提供任务矩阵、恢复记录和验收清单。
- `.agents/skills/chainmaker-homework/` 包含执行步骤，以及扫描 PDF、解析 Word 模板、生成与验证报告的脚本。

每份作业的课程资料、实验记录和交付文件均由使用者在本地准备。
