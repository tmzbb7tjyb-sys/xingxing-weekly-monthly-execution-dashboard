# 星星电台周月执行进度看板

这是一个 Codex 技能，用于按月份和任务清单生成、更新或审阅星星电台 PM 执行进度 Excel 看板。工作簿包含月度汇总、每个产品的任务明细、自然周计划与实际、阶段反馈和进度预警。

## 安装与依赖

将本仓库放在 Codex 的技能目录 `~/.codex/skills/xingxing-weekly-monthly-execution-dashboard/`。生成脚本需要 Node.js、工作区依赖 `@oai/artifact-tool`，以及用于最终公式重算的 LibreOffice 或 Excel。`@oai/artifact-tool` 必须能从运行脚本时的当前目录解析；可在临时运行目录中把 `node_modules` 链接到已提供该依赖的工作区目录。

仓库内容：

- `SKILL.md`：技能工作流程和验收要求。
- `references/dashboard_contract.md`：输入 JSON、工作簿结构和公式口径。
- `scripts/generate_dashboard.mjs`：新建看板的生成脚本。
- `agents/openai.yaml`：Codex 技能展示配置。

## 新建看板

先根据 `references/dashboard_contract.md` 准备 JSON。未提供的目标、周计划、实际和负责人应留空，不要编造正式数据。以下是最小配置示意：

```json
{
  "month": "202611",
  "checkpoint": { "mode": "auto" },
  "tasks": [
    {
      "name": "示例产品",
      "brand": "",
      "product": "",
      "pm": "",
      "task_id": "",
      "monthly_targets": {
        "samples": null,
        "scripts": null,
        "publishes": null,
        "amount": null
      },
      "weekly_plans": [],
      "actuals": [],
      "feedback": []
    }
  ],
  "output": "/absolute/path/dashboard.xlsx"
}
```

在可以解析 `@oai/artifact-tool` 的目录中运行：

```bash
node /path/to/xingxing-weekly-monthly-execution-dashboard/scripts/generate_dashboard.mjs \
  --config /absolute/path/config.json
```

脚本不会默认覆盖已有输出；仅在确认可以覆盖时使用 `--force`。默认自动检查点会参考 `TODAY()` 和已填写的阶段数值；只有明确需要手动控制时，才在配置中使用 `checkpoint.mode="manual"`。

## 交付前核验

生成的汇总页使用 `INDIRECT` 引用产品名称工作表。当前 `@oai/artifact-tool` 不能完整重算该函数，因此必须用 LibreOffice 或 Excel 打开并全量重算、另存最终工作簿，再核对公式缓存、空白与真实零值、进度状态、条件格式及错误值。详细检查项见 `SKILL.md` 和看板契约。
