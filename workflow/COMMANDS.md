# QBWordAgent 固定命令手册（Windows 原生 PowerShell）
所有命令由Agent执行。工作目录为解压后的项目根；不需要Python。每条命令代表独立步骤，不批量伪造用户确认。固定方法仍只在 WORKFLOW.md，本文件不修改方法。

| 用途 | 命令 |
| --- | --- |
| 展示完整工作流 | `./agent.ps1 workflow` |
| 只读环境检查 | `./agent.ps1 doctor` |
| 查看材料清单（不读内容） | `./agent.ps1 input-status` |
| 允许读取／跳过材料 | `./agent.ps1 choose-input --state dir/日期/workflow_state.json --mode read --confirm`（不读取用skip） |
| 真实排版能力探测 | `./agent.ps1 doctor --probe` |
| 查看学年/学期建议 | `./agent.ps1 intake-options` |
| 新建任务 | `./agent.ps1 start --topic "题目"` |
| 记录学年学期 | `./agent.ps1 set-term --state dir/日期/workflow_state.json --academic-year 2025-2026 --semester 1 --confirm` |
| 提交大纲 | `./agent.ps1 save-outline --state dir/日期/workflow_state.json --source dir/日期/outline.draft.md` |
| 闸门1批准 | `./agent.ps1 approve-outline --state dir/日期/workflow_state.json --pages 8 --confirm` |
| 默认题目确认 | `./agent.ps1 set-report-title --state dir/日期/workflow_state.json --mode default --project-name "图书管理系统" --confirm` |
| 自定义题目确认 | `./agent.ps1 set-report-title --state dir/日期/workflow_state.json --mode custom --project-name "图书管理系统" --title "图书借阅系统设计与实现" --confirm` |
| 闸门2代码选择 | `./agent.ps1 choose-code --state dir/日期/workflow_state.json --mode none --confirm` |
| 正文骨架 | `./agent.ps1 prepare-content --state dir/日期/workflow_state.json` |
| 预检，不启动Office | `./agent.ps1 generate --workflow-state dir/日期/workflow_state.json --input dir/日期/report_content.json --output 报告.docx --check-only` |
| 自动生成并排版 | `./agent.ps1 generate --workflow-state dir/日期/workflow_state.json --input dir/日期/report_content.json --output 报告.docx` |
| 闸门3 | `./agent.ps1 review-decision --state dir/日期/workflow_state.json --approve --confirm`（退回用 --reject） |
| T3最终检查 | `./agent.ps1 final-check --state dir/日期/workflow_state.json` |
| 闸门4发布 | `./agent.ps1 finish --state dir/日期/workflow_state.json --output 报告.docx --confirm` |
| 状态 | `./agent.ps1 status --state dir/日期/workflow_state.json` |
| 回退 | `./agent.ps1 rollback --state dir/日期/workflow_state.json --to t2 --reason "原因"` |
| 打包 | `./agent.ps1 package` |
| 清理预览/执行 | `./agent.ps1 clean` / `./agent.ps1 clean --apply` |
| 发布后清空本任务材料 | `./agent.ps1 clean --input --state dir/日期/workflow_state.json`，展示计划后 `./agent.ps1 clean --input --state dir/日期/workflow_state.json --apply --confirm` |
| 用户已自行移走，验证为空 | `./agent.ps1 clean --input --state dir/日期/workflow_state.json --manual`，通过后同命令加 `--apply --confirm` |

--engine auto 为默认，另有 word/wps；引擎通过真实探测后绑定任务。--word-timeout 默认为180秒，范围10至600；不盲目延长。--force-rebuild 忽略缓存，但不能越过闸门3修改已交审内容，先回退T2。
rollback支持t1/gate1/gate2/t2。prepare-content拒绝覆盖正文。--confirm只能记录用户本轮真实确认。
旧render-job队列不再是默认流程；保留旧文件并用原任务执行generate。旧Python入口已退役清理。
正文修改后必须重新验证。最终文档和验证回执的哈希都绑定在状态中，批准后修改会阻止发布。
权限错误由Agent使用客户端合法授权执行机制处理，不要求用户手动复制终端命令；无此能力时报告环境限制。

维护清理：`./agent.ps1 clean --retired` 预览本机明确列出的退役文件，再 `./agent.ps1 clean --retired --apply` 移入Windows回收站。未列明文件不删除，保留任务、报告、模板和错误约束样本。

新任务先停在T1_TERM_REQUIRED，set-term后才可提交大纲；approve-outline后停在REPORT_TITLE_WAITING_CHOICE，set-report-title后才可choose-code。题目确认不是代替四道闸门。旧任务缺少这些信息时必须显式回退T1补充，已有文件保留。

非空input的新任务先choose-input才可set-term；为空不提问、不需要choose-input。choose-input只记录权限及指纹，实际项目理解由AI完成，materials.md放在本任务目录。finish仅发布报告，不删除input；返回next_action提示材料收尾。clean --input只允许已COMPLETE且有待处理材料的所属任务，必须先预览；manual验证用户已移走，非manual移入回收站。清理未完成禁止start新任务。旧任务没有input记录时不猜测归属、不自动清理。
