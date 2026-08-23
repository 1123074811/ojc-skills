# Root-cause repair

Use this mode when a task involves a defect, inconsistent behavior, a rejected implementation, repeated workarounds, or the same failure appearing across components. The goal is to make the accepted behavior native to the system, not to make one reproduction stop failing.

## Reconstruct the accepted model

Before proposing code, form an internal repair contract:

- **Positive invariant:** what must always be true, stated without defining the result as “not the rejected approach.”
- **Authoritative owner:** the domain model, schema, state machine, service, boundary, configuration, or other source of truth responsible for that invariant.
- **Observed baseline:** shipped behavior, persisted state, public contracts, task-start repository state, and pre-existing user changes that the repair must preserve or migrate.
- **Affected surfaces:** callers, APIs, UI, storage, tests, deployment configuration, documentation, and delivery metadata that consume or describe the invariant.
- **Required exceptions:** only real safety, policy, compatibility, migration, audit, or user-approved variation.

Treat rejected proposals, correction wording, failed assistant edits, and temporary experiments as diagnostic history. They do not become product concepts, configuration switches, compatibility requirements, comments, or names merely because they appeared in the working session.

## Trace the first divergence

Work backward from the visible symptom until the first state or decision differs from the positive invariant. Determine whether the owner is wrong or a consumer violates an otherwise correct contract. Useful cause classes include:

- an incorrect or missing domain invariant;
- a data or state model that represents the wrong concept or permits invalid states;
- duplicated or conflicting sources of truth;
- an incorrect transition, ordering, lifecycle, concurrency, or control-flow rule;
- missing validation, authorization, or normalization at a trust boundary;
- API, schema, type, default, or error-semantics drift between layers;
- incomplete migration or real backward-compatibility handling;
- an isolated implementation mistake below a correct shared contract.

Stop at the narrowest owner that can express and enforce the invariant for every affected consumer. Do not manufacture a shared abstraction when the evidence shows a genuinely isolated error, but do not call a caller-level guard a root-cause fix when the owner remains wrong.

## Choose replacement over compensation

Prefer a change that makes the correct path ordinary and deletes the need for special handling.

| Signal | Patch-shaped response | Root-cause response |
|---|---|---|
| The same condition appears at many callers | Add another caller guard | Enforce the invariant at the shared owner or boundary |
| Frontend and backend translate the same field differently | Add another adapter | Repair the schema or contract and synchronize consumers |
| Invalid state is repaired late | Add cleanup after each failure | Prevent or reject the invalid state at creation or ingress |
| A rejected UI or workflow survives as a flag | Invert the flag and keep both branches | Model the accepted workflow directly and remove the session-only branch |
| A label, comment, or test is framed around the rejected idea | Rename it to “no old idea” | Regenerate it from the accepted behavior and current invariant |
| A timing failure is hidden by retries or longer waits | Increase retries or timeout | Repair lifecycle, readiness, ordering, cancellation, or backpressure unless real environmental variance requires tuning |

Do not preserve the faulty semantic frame through a synonym, negative flag, wrapper, fallback, or explanatory comment. A real released API, persisted format, external client, staged rollout, or operational rollback requirement may justify a compatibility bridge; document why it exists, test both sides, and define how it ends.

## Implement the new invariant

Adapt the sequence to the repository and risk:

1. Add or update a focused test that observes the accepted invariant when doing so will not freeze known-wrong behavior.
2. Change the authoritative owner and its representation or transition.
3. Update affected consumers as a complete vertical slice.
4. Remove obsolete guards, flags, adapters, fallbacks, dead branches, stale comments, and tests whose only purpose was compensating for the replaced logic.
5. Add an explicit migration or compatibility path only when the observed baseline requires it.
6. Update documentation and user-facing text from the resulting behavior, not from the correction narrative.

Avoid running old and new models side by side by default. If a temporary dual path is required, make selection unambiguous, observable, reversible, and bounded by a removal condition.

## Prove eradication and preservation

Verification must show both that the cause is gone and that required behavior remains:

- run the original reproduction;
- test the invariant directly, including counterexamples and failure states;
- exercise sibling callers, equivalent implementations, and every affected layer;
- search for duplicated faulty rules and artifacts of the obsolete model, then inspect matches semantically rather than treating zero keyword matches as proof;
- confirm public contracts, persisted data, safety and authorization boundaries, and unrelated user changes were preserved or deliberately migrated;
- review whether the repair reduced accidental complexity instead of moving it elsewhere.

If only the reported symptom could be repaired or the owning cause could not be established, describe the result accurately as a scoped repair and name the remaining uncertainty. Do not claim a root-cause fix from one passing reproduction.

## Finalize from the accepted state

Generate comments, documentation, commit and PR text, release notes, and the final handoff from the authoritative diff and observed final state. Omit alternatives that existed only in discussion or temporary edits. Preserve real removals, migrations, compatibility changes, security facts, executed external operations, and user-requested comparisons when the audience needs them.

When a rejected alternative and its rationale must remain to prevent a material recurrence, put that decision in the appropriate ADR, migration guide, or audit record. Do not spread it into unrelated identifiers, titles, comments, or delivery labels.
