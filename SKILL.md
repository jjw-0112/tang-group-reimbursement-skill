---
name: tang-group-reimbursement
description: 为课题组学生从发票和附件文件夹制作可补件续跑的个人集中报销 Excel、缺件 PDF 与 ZIP；适用于首次整理、更正发票和补充材料，不用于课题组总表或财务系统填报。
---

# 课题组个人集中报销

使用本 skill 时，先读 [报销规则](references/rules-v6.md)。V6 手册优先；语音整理稿仅补充明确且适用于学生交件的规则。所有金额、票号、抬头和附件关系都要有原件依据；不要从旧文件名猜测或承诺财务最终批准。

## 运行环境

使用 Python 3.10+，安装本目录的 `requirements.txt`。扫描件 OCR 需要本地 Tesseract 与 `chi_sim`、`eng` 语言包；RAR 需要 7-Zip/`7z`。运行 `python scripts/reimburse.py doctor` 检查环境。缺失可选工具时继续处理可读取文件，将受影响文件列为待核实，并告知学生如何补装。脚本使用路径对象，支持 Windows、macOS 和 Linux；不访问外部服务。

## 首次整理

1. 询问学生原始材料目录、姓名、学号、电话、**报销截止日期**，以及可选的历史报销目录。原目录不得作为输出目录；不要移动或改名原件。
2. 运行 `python scripts/reimburse.py scan --input <原目录> [--output <输出目录>] [--history <历史目录>]`。默认输出目录是原目录的同级目录 `<原目录名>_报销输出`。检查生成的 `review.json`，特别是 `needs_review`、`unreviewed_files` 和 `replacement_candidates`。无法读取的图片、未识别附件须人工确认；确认为无关文件时，把路径写入 `ignored_attachment_refs`，不要直接忽略。
3. 将学生信息、截止日期写入 `review.json` 的 `profile`；向学生**一次性询问哪些发票属于对公转账**，把票号或记录 ID 写入 `public_invoice_numbers` / `public_record_ids`。其余视为个人垫付。不要逐张询问付款方式。
4. 逐项复核模型和脚本提取的票面金额、商品明细、类别、日期、抬头以及附件关联。只在看到证据后填写记录的 `manual` 字段：`category`、`confirm_mixed_category`、`description`、`invoice_number`、`issue_date`、`claim_amount`、`confirm_filename_amount`、`confirm_claim_amount`、`confirm_payment_amount`、`buyer_name`、`buyer_tax_id`、`confirmed_buyer_absent`、`confirmed_wrong_buyer`、`confirmed_nonreimbursable`、`paper_original_confirmed`、`attachment_refs`、`replacement_for`、`notes`。不清楚的保持待核实；不要为赶进度填造确认值。
5. 运行 `python scripts/reimburse.py build --review <review.json>`。检查终端摘要、Excel 暂计金额、PDF 缺件事项、未匹配文件、文件对照表和 ZIP。缺件发票进入 Excel 并标“缺件”；待核实和不可报销项不计入金额。PDF 显示“尚不可提交”时，提醒学生补齐或核实后续跑。

## 补件续跑

学生向同一原目录追加附件或更正发票后，重新运行 `scan`，指向原输出目录。脚本读取 `state.json`，沿用个人信息和已确认判断，重新匹配附件，并指出可能替换的旧票。对替换关系向学生确认，在新记录的 `manual.replacement_for` 填旧记录 ID；不要仅凭相似名称自动替换。再次 `build` 会生成 v2、v3 等新包、重新连续编号，并附旧新编号对照。旧版本保留。第一版不解析负责人或财务的退回意见。

## 复核重点

- Excel 保持“一张发票一行”、全局连续编号；材料费中含达到门槛的商品行的整张发票排在组后部，备注指出具体商品。个人信息只填在明细页现有首行。
- 普通增值税发票抬头确认为缺失或不符时列“不可报销”；识别不清则“待核实”。火车、机票和出租车等按票种规则检查。
- 票面与旧文件名或付款证明不一致时，使用票面提取值并先列待核实。满 1000 元个人垫付发票按 V6 列不可报销；所有对公票需要付款依据。
- 逐票检查适用附件和纸票原件。对未匹配附件、疑似历史重复、跨年度票据、截止日期之后票据写出具体原因。仅做离线检查，PDF 必须写明未做税务平台验真。

若需要修改规则，先更新 [报销规则](references/rules-v6.md) 中的版本和来源，再更新脚本与测试；不要把某次报销的例外永久写成通用规则。

同学安装与操作可参照 [使用说明](使用说明.md)。
