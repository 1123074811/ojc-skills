# Development profile

Use these as defaults only when the current request and repository do not say otherwise.

## Observed working style

- The user favors persistent execution: “继续”, “完成剩余所有内容”, and “全部完成后再结束” recur often.
- The user separates discussion from implementation when exploring product direction.
- Non-trivial work should begin with repository understanding and an executable plan; write it to a project document only when the user requests it or cross-session continuity materially benefits.
- Acceptance means the feature is complete across affected layers, compiles or runs, and has been tested—not merely that code was written.
- Reviews should actively find omissions, obvious bugs, weak error handling, stale files, and opportunities to simplify.
- UI quality includes sensible sizing, no overlap or overflow, robust text wrapping, and real browser verification when available.
- Production data, formulas, and conclusions need traceable sources. Avoid fake metrics, filler data, unsupported constants, and hard-coded secrets.
- Git delivery includes tracked and untracked files, `.gitignore`, secret scanning, coherent Chinese commit messages, and pushing only when requested.
- When the user asks to preserve a successful workflow, distill it into a skill instead of repeatedly pasting prompts.

## Default interaction

- Inspect before questioning.
- Ask only about choices that materially alter the result; do not repeatedly seek confirmation for safe work already authorized.
- Give concise progress updates during long work.
- Treat “继续” as resuming the latest unfinished plan after checking current state, not as permission for unrelated scope expansion.
- Treat explicit “只讨论/不要实现” as read-only even if implementation seems obvious.
- When a prompt rewrite would add no value, internalize the clarified brief and proceed.

## Typical environments

The local portfolio contains Windows/PowerShell projects and many Java/Spring, Vue/Vite/React, Python, Docker, AI, document, and visualization workflows, commonly under `E:\code`. Detect the actual stack from each repository; never force these technologies onto a project.

## Representative trigger requests

- “先了解这个项目并和我讨论，不要实现。”
- “根据需求生成可执行的开发文档，再按阶段完成。”
- “继续把剩余功能全部补全，不要反复问我。”
- “检查有没有功能遗漏和 bug，修复后确保前后端能正确运行。”
- “这个字段要同步数据库、后端、前端和部署配置。”
- “检查样式的遮挡、溢出和尺寸问题，并用浏览器验证。”
- “审查所有 Git 修改，检查密钥和临时文件，写中文提交信息并按要求推送。”
