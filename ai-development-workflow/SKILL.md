---
name: ai-development-workflow
description: 跨阶段、多组件软件项目的需求澄清、计划、实现、系统性根因修复、验证和交付工作流。用于新项目从模糊需求落地、恢复未完成开发、复杂功能全链路补全、反复缺陷或错误设计的根因修复、分阶段优化、项目验收与安全 Git 交付；尤其适合“先理解/讨论”“生成开发计划”“继续完成剩余阶段”“检查遗漏并修复”“确保前后端对接和运行”。不用于一次性小改动、单文件简单修复或纯知识问答。
---

# AI Development Workflow

Move from intent to verified delivery without repeatedly making the user restate the same workflow.

## Operating contract

- Inspect the repository, instructions, documentation, configuration, and current Git state before choosing an approach.
- Reuse existing code, project conventions, native capabilities, installed dependencies, and relevant skills before adding anything.
- Preserve unrelated user changes. Do not commit or push unless the user explicitly requests it.
- Delete or move existing user files, discard changes, reset history, deploy, or mutate external systems only when explicitly authorized.
- Treat “只讨论/不要实现” as a hard phase boundary. Discuss and clarify without editing until the user authorizes implementation.
- Anchor “全部完成/继续” to the latest active plan, acceptance criteria, and explicitly requested artifacts. Treat old TODOs and wish lists as evidence to assess, not automatic authorization.
- Continue through safe, authorized work until the requested outcome is complete. Ask only when a missing choice materially changes the result.
- For defects and rejected designs, reconstruct the positive invariant and authoritative baseline before editing. Do not turn the reported symptom, correction wording, or discarded approach into the target architecture. Repair the component that owns the violated invariant, replace incorrect logic, remove obsolete compensations, and then verify equivalent paths instead of layering a guard onto the reported location.

## 1. Establish intent

Classify the request as discussion, planning, implementation, audit, verification, delivery, or resumption. Combine phases only when the request authorizes them.

1. Read local context first; do not ask for facts that files or tools can answer.
2. Resolve material ambiguity until roughly 90% confident about the goal, users, constraints, acceptance criteria, and non-goals.
3. Ask one focused question at a time when blocked. Otherwise state a reasonable assumption and proceed.
4. Internalize an improved task brief. Output a rewritten “perfect prompt” only when the user asks for one.

Read `references/development-profile.md` when calibrating defaults or interpreting a short request such as “继续”.

## 2. Baseline the project

- Trace only the flow the request touches: entry point, callers, data model, API, UI, persistence, configuration, tests, and deployment path as applicable. For discussion-only work, inspect only enough context to support the discussion.
- Check repository instructions and existing plans before writing a new plan.
- Separate authoritative project state from working-session history. Rejected proposals and temporary attempts are diagnostic evidence, not requirements; shipped behavior, persisted data, compatibility contracts, executed external operations, and pre-existing user changes remain real baseline facts.
- For defect, rework, inconsistent-behavior, or repeated-workaround tasks, trace the visible symptom back to the earliest violated invariant and its owning source of truth. Read `references/root-cause-repair.md` before choosing the repair level.
- For changing, niche, security-sensitive, or data-backed claims, use primary or authoritative sources and distinguish facts from inference.
- Reject fake, placeholder, simulated, or unexplained production data. Record the source or derivation of formulas, thresholds, and constants that affect product behavior.
- Report a blocker only after exhausting safe local checks and in-scope alternatives.

## 3. Plan only as much as needed

For non-trivial work, prefer an existing plan or a short in-conversation task list. Create or update one living project document only when the user requests documentation or cross-session continuity materially benefits. Include only what execution needs:

- goal, scope, non-goals, constraints, and acceptance criteria;
- current-state findings and affected components;
- proposed architecture, data/API changes, migration and compatibility notes;
- ordered phases with dependencies and verification per phase;
- deployment, rollback, and unresolved risks when relevant.

Skip new documentation for a small isolated fix when an existing issue, plan, or concise task list is enough. Split work between agents only when tasks are independent and cannot overwrite the same files.

## 4. Implement vertical slices

- Implement the smallest complete path that proves the behavior, then extend it.
- Express the accepted behavior positively in the domain model, state transition, contract, or owning abstraction. When that representation is wrong, replace it and update its consumers; do not preserve the rejected design through inverted flags, caller-specific exceptions, fallback chains, or comments that merely say what no longer happens.
- Keep a compatibility bridge only for a real released, persisted, or external contract. Give a temporary bridge an explicit reason, removal condition, and verification; session-only experiments do not justify permanent compatibility code.
- Keep database, backend, frontend, validation, permissions, documentation, and deployment configuration synchronized when a field or workflow crosses layers.
- Follow the repository's established patterns. Avoid speculative abstractions, duplicate helpers, unnecessary dependencies, and unrelated redesigns.
- Preserve existing calibration and configuration knobs for hardware, model, network, or environment-dependent behavior; add new knobs only when variability is demonstrated.
- If a living plan exists, update it after meaningful phase completion so a resumed session can continue without reconstruction.
- Default to local, test, or isolated environments. Accessing production data, mutating external services, deploying, or running costly third-party operations requires explicit authorization.

## 5. Verify and repair

Verification is part of implementation, not a separate optional handoff.

1. Run the narrowest relevant tests, then broader build, integration, or browser checks only when the affected path and risk justify them.
2. Exercise the changed user path with realistic data. Verify failure states as well as the happy path.
3. For UI work, inspect common viewport sizes, text wrapping, overflow, overlap, loading, empty, error, disabled, keyboard, and focus states.
4. For a root-cause repair, test the positive invariant directly, rerun the original reproduction, exercise sibling and counterexample paths, and search for remaining copies of the faulty rule or obsolete workaround. Passing only the reported example does not prove eradication.
5. Review the final diff for omissions, security issues, secrets, temporary files, generated artifacts, stale code, accidental unrelated edits, and complexity that exists only to compensate for the replaced logic.
6. Fix failures within scope and rerun the check that exposed them.

Read `references/delivery-checklists.md` during verification and before delivery. Do not claim a check passed unless it actually ran.

## 6. Deliver safely

- Lead with the outcome, then list changed areas, verification performed, and any concrete remaining limitation.
- Derive titles, commit or PR text, release notes, and the handoff from the task-owned final diff and observed state. Do not identify the result by a rejected proposal or temporary repair path. Name a removal, migration, compatibility break, safety constraint, or partial external action when it is a real baseline fact the audience needs.
- Before finishing, privately challenge the result: lowest-confidence assumption, biggest omission, likely three-month failure mode, and whether one high-value improvement belongs in scope. Fix in-scope issues; mention only actionable out-of-scope items.
- When Git delivery is requested, inspect all tracked and untracked changes, ignore or remove only confirmed disposable files, and scan for secrets.
- Partition authorized changes by independently reviewable and revertible intent before staging. Keep implementation with its tests, migration, and required documentation; do not split mechanically by file or file type.
- Create multiple commits when independent concerns exist, in dependency order. Inspect the staged diff before each commit and the remaining worktree after it; do not mix unrelated user changes into any commit.
- Write clear Chinese commit messages when the user asks for Chinese. Push only when explicitly requested, after all requested commits pass verification, and report the resulting branch and commit sequence or failure.
