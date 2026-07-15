# Delivery checklists

Apply only the sections relevant to the task.

## Discovery

- Read repository instructions, README, plans, package/build files, environment examples, and Git status.
- Find existing implementations and trace callers before changing shared behavior.
- Record scope, acceptance criteria, constraints, non-goals, and material assumptions.

## Plan

- Map affected UI, API, service, persistence, configuration, tests, deployment, and documentation.
- Order work by dependency and attach a runnable verification step to each phase.
- Note migration, compatibility, rollback, security, and data provenance where relevant.

## Implementation

- Reuse existing patterns and dependencies.
- Validate inputs at trust boundaries and preserve authorization checks.
- Synchronize cross-layer names, types, defaults, validation, and error handling.
- Avoid unrelated refactors, duplicate helpers, fake data, unsupported constants, and silent fallbacks.

## Verification

- Run focused tests, then the relevant build and integration or browser flow.
- Exercise success, empty, loading, validation, permission, network, and failure states.
- Check responsive layout, overflow, overlap, wrapping, focus, keyboard use, and accessible labels for UI changes.
- Verify migrations and deployment configuration against a realistic environment when applicable.
- Rerun failed checks after fixes; record commands and results.

## Git delivery

- Review `git status`, diff, and untracked files without discarding unrelated user work.
- Check for secrets, credentials, personal data, debug logs, generated output, caches, and temporary files.
- Add `.gitignore` entries only for reproducibly disposable artifacts.
- Commit by coherent concern with the requested message language.
- Push only when explicitly requested; report branch, commit, and remote result.

## Completion

- Confirm every acceptance criterion or name the exact blocker.
- State what changed and what verification actually ran.
- Surface only actionable residual risks; do not pad the handoff with speculative feature ideas.
