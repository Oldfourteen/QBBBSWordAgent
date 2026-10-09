# 训练样本说明

`approved/` 只保存经人工确认合格的脱敏 JSON 与评语；`rejected/` 保存不合格 JSON 或脱敏 DOCX、触发的校验项和修正建议。禁止放入姓名、学号、成绩单原件、联系方式或其他个人信息。

`rejected/error_correction_constraints/` 用于保存真实生成错误的负面示范。每个错误单独建立目录，至少包含原样错误文件、`sample_manifest.json`、`constraint_node.json` 和 `failure_report.md`。`constraint_node.json` 是供 Agent 优先读取的机器约束节点；错误文件只用于识别和拒绝同类问题，不得作为正确格式来源。
