# 课题组个人集中报销 Skill

从学生自己的发票与附件文件夹生成个人报销明细 Excel、补件提醒 PDF、文件对照表和带版本号的 ZIP。补充材料后可以在同一原材料目录续跑，保留旧版并提供编号变更对照。

本 Skill 根据 Tang Group 集中报销手册 V6 整理规则，只做**离线材料核对**；不访问学校财务系统或税务验真平台，也不保证最终报销获批。仓库不包含学生发票、个人信息表或任何实际报销结果。

## 安装

在 Codex 中使用 `$skill-installer`，让它从本仓库根目录安装，技能名称为 `tang-group-reimbursement`：

> `$skill-installer` 请安装 `https://github.com/jjw-0112/tang-group-reimbursement-skill` 仓库根目录的 Skill，名称设为 `tang-group-reimbursement`。

也可以手动下载仓库 ZIP，把解压后的整个文件夹放入 Codex 个人技能目录。安装 Python 3.10+ 后执行：

```bash
python -m pip install -r requirements.txt
python scripts/reimburse.py doctor
```

若要处理扫描件，另需 Tesseract OCR 及 `chi_sim`、`eng` 语言包；读取 RAR 需要 7-Zip/`7z`。缺少可选工具时，受影响材料会进入待核实。

完整步骤见 [使用说明](使用说明.md)。

## 使用

在 Codex 中输入：

> `$tang-group-reimbursement` 请整理我的集中报销材料，原始文件夹在 `<路径>`。先收集姓名、学号、电话、报销截止日期，并一次性确认对公转账发票。

第一次生成的 `review.json` 需要复核识别结果。`build` 后会产生 v1 ZIP。向**同一个原材料目录**追加补件，再调用本 Skill，会复查未结问题并生成 v2、v3 等新版本。原始文件不会移动或改名。

规则摘要见 [V6 规则](references/rules-v6.md)。
