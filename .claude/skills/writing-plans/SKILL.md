---
name: writing-plans
description: Use when the user explicitly asks for an implementation plan, task breakdown, or written technical approach before implementation.
---

# 编写实施计划

## 原则

只在用户明确要求计划时使用。计划应帮助当前项目落地，不引入额外的代理、测试、Git 或文档流程。

## 编写流程

1. 阅读需求及直接相关的项目上下文。
2. 明确目标、范围、约束和必要假设。
3. 按依赖顺序拆分步骤；可以并行的步骤要明确标注。
4. 为每一步写清涉及文件、预期改动、边界情况和静态 review 要点。
5. 列出会影响方案的风险或待确认事项。

## 输出格式

```markdown
# [任务名称] 实施计划

## 目标
[一句话描述结果]

## 实施步骤
1. [步骤名称]
   - 文件：`path/to/file`
   - 改动：[具体行为]
   - 注意：[边界情况或依赖]

## 静态 Review
- [需要重读的逻辑和引用]

## 风险
- [仅列出真实存在的风险]
```

## 约束

- 遵循现有架构和代码模式，避免无关重构。
- 不写 `TODO`、"适当处理"、"后续补充"等占位内容。
- 除非用户明确要求，否则不创建计划文件，只在回复中给出计划。
- 除非用户明确要求，否则不加入测试、构建、lint 或类型检查命令。
- 不安排 commit、push、分支、PR 或 worktree 操作。
- 用户同时要求实施时，计划确认后由当前会话直接执行，不要求调用其他流程 skill。
