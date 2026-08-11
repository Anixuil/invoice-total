---
name: writing-skills
description: Use when creating or revising a project-local skill whose reusable instructions are specific to this repository.
---

# 编写项目 Skill

## 原则

只保留模型无法从项目代码和 `AGENTS.md` 直接推断、且会重复使用的专业流程。项目级硬性规则优先放在 `AGENTS.md`，不要复制到多个 skill。

## 工作流程

1. 明确触发场景和不适用场景。
2. 检查是否已有同职责的 skill 或项目规则。
3. 保留最少的必要步骤、约束和领域知识。
4. 将触发条件写入 frontmatter 的 `description`。
5. 重读 skill，确认没有失效引用、重复流程或与 `AGENTS.md` 冲突的要求。

## 结构

```markdown
---
name: skill-name
description: Use when [具体触发条件].
---

# Skill 标题

## 原则
[核心判断]

## 流程
[必要步骤]

## 约束
[项目特有边界]
```

## 约束

- 名称使用小写字母、数字和连字符，并与目录名一致。
- `SKILL.md` 默认保持在 200 行以内；详细参考仅在确有需要时拆分。
- 不为标准工程常识创建 skill。
- 不引入测试、子智能体、commit、push、PR 或 worktree 流程。
- 修改后只对本次变更做一次静态逻辑 review。
