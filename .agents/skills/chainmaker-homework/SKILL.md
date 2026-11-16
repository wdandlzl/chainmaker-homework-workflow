---
name: chainmaker-homework
description: 完成长安链训练营单份作业；从当前 PDF 建立可恢复任务，复用模板缓存，生成并验证真实实验报告。
---

# 单份作业执行
遵守根目录 AGENTS.md。脚本路径均相对本 Skill 的 scripts/；命令使用可用 Python（Windows 的 python 别名可能无效，可通过工作区依赖工具定位运行时）。

1. 先读 `.workflow/tasks/<作业名称>/TASK.md`。核对已验证项的输入、输出、证据哈希和环境，失效项及其依赖退回待验证；仍有效的不重跑。不存在则按 `.workflow/TASK_TEMPLATE.md` 创建，禁止覆盖既有任务。
2. 仅对指定 PDF 执行 `scan_pdf.py --pdf <路径> --output-dir <缓存目录>`，定位【实验作业】及跨页续文，按原题登记 ID、来源页码、依赖、成果、验收和证据。遇图片公式须视觉核对。
3. 执行 `template_cache.py --docx <原模板> --cache .workflow/template_rules.json`。缓存命中不重复模型分析。首次、SHA256 变化或无效时，缓存保存全部 Word XML（含样式、页眉页脚、图形引用）、段落和表格；再用 `inspect_template.py --docx <模板> --out <扫描.md>` 并实际渲染核实图形、格式、说明及规则。新缓存 requirements 为空且待核实，必须重新填写，不自动沿用旧规则。完整视觉检查后才将 review_status 标为已验证；修改版模板还需核对生成器段落/单元格映射。
4. 独立执行真实实验。查错先关键词检索唯一 `.workflow/corrections.md`，检查适用环境。定位页码、实际现象及根因；结合官方文档/帮助/运行结果做最小修正，不改原任务。原则上最多三种实质不同方案；失败记阻塞并继续独立项。仅实际错误或阻塞新增记录，区分已确认/待核实，日志另存，禁止二级纠错库。
5. 报告输入 JSON 包含 fields、summary，优先使用 sections。每节包含 task_ids（对应 TASK 矩阵）、title、paragraphs，可按需提供 commands（实际命令数组）、results（headers 和 rows，单元格用字符串）、figures（path 为真实图片绝对路径、caption 说明验证结果）。旧输入继续使用 steps 和 images。fields 必含 course/team/name/school/contact/team_en/date/declaration_date/partners1/partners2/project/equipment/division；个人信息来自用户，无同实验者按实际填写。正文按任务写明关键配置、运行命令或附件入口、实际结果和对应证据。排障写清现象、原因、修正和复验，未确认的原因明确标注。同一次运行的结果相互核对后再写入报告。截图紧跟相关步骤，可使用保留关键输出和身份的裁剪副本，原图留作证据。模拟测试不得成为作业证据。
6. `report_generator.py --template <原模板> --rules .workflow/template_rules.json --data <输入.json> --out <课程+队伍+姓名+学校.docx> --pdf` 基于原模板生成，保留样式和表格；拒绝覆盖已有输出。PDF 使用现有 Word COM，不自行安装软件。导出失败修复或记录阻塞，不能以 DOCX 代替最终 PDF。
7. `validate_report.py --report <DOCX> --rules .workflow/template_rules.json --data <输入.json> --json-out <验收.json>` 检查实际报告中的字数、图片、字段、说明及 PDF 命名。原 `validate_delivery.py --task <TASK.md> --report <DOCX> --output-dir <交付目录> --json-out <结构验收.json>` 仅作为补充结构检查，不能替代专项验收。
8. 实际渲染 DOCX/PDF，逐页检查文字、图片、空白及分页，确认截图中的关键输出和要求的身份可辨认，图注与截图相邻。按最终 PDF 逐题核对复现入口、正文结论和证据，在 TASK 矩阵登记页码及小节、图号或表号，记录输入输出哈希及结论。修订后重新渲染并检查受影响页面及其分页。程序检查有序结构，实验真实性、人员信息真实性和任务覆盖由内容验收确认。全部通过才已验证。独立完成后可选少量同题报告查漏。状态统一：待执行、执行中、已生成待验证、已验证、阻塞。

纠错记录字段：问题 ID、所属作业/任务 ID、类型、确认状态、PDF 页码（适用时）、现象/原因、修正方案、环境版本、验证结果、证据路径。相同问题更新原记录，不重复新增。任务仅链接问题 ID。

## 总结与承诺声明专项要求
- 总结以学生本人视角描述实验结果、遇到的问题及处理方法、心得和不足；仅依据真实记录，不虚构体验。不把Codex生成、截图核验和证据管理过程写成学生心得。证据限制写入交付说明或验收记录。
- 原样保留模板三句承诺声明及格式，填写独立字段declaration_date；日期采用本次报告确认日期，不篡改实验执行日期。声明由用户提交前审核。
- 只删除已确认的填写说明及示例段落，禁止整体删除总结后的内容；模板SHA变化时必须重新核对删除索引与声明槽位。
- 程序验收检查声明完整、有序且唯一、声明日期真实有效且与输入一致；总结字数仅计算声明之前正文。内容验收检查学生视角、实验结果、问题处理及心得，不以关键词测试替代内容审核。
