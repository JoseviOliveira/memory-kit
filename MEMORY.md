# Memory Convention

A portable, four-level memory convention for any coding project. It keeps the knowledge produced while coding — context, decisions, dead ends, durable lessons — captured, refined, and preserved instead of lost to context resets.

This is the **single source of truth**. Everything an agent needs is here: the model, the rules, and three verbs (**capture / recall / curate**). No other files are required.

---

## Project settings

Edit these two values when dropping this kit into a project:

- `<notes-dir>` = `notes/`  ← where L0/L1 working notes live (git-ignore `L0_*`)
- `<docs-dir>` = `docs/`     ← where L3 curated docs live

---

## Philosophy

1. **Capture fast, curate later.** Never lose signal because writing it down was expensive. Dump raw, refine afterward.
2. **Promotion must be earned.** Knowledge moves toward durable only when validated, reusable, and stable. Speculation stays hot.
3. **Cleanup over accumulation.** Stale, duplicated, low-signal notes are liabilities. Pruning is a first-class activity.
4. **Failures are valuable.** A tested approach that failed prevents the same mistake twice. Preserve anti-patterns; never silently delete them.
5. **Authority is explicit.** When sources disagree, follow a strict order of trust. Raw chat never overrides curated docs or live code.
6. **Honesty over confidence.** If a source was unavailable, say so and lower confidence. Never fabricate coverage.

---

## The model (canonical — defined once)

Knowledge flows from **hot/raw** to **cold/curated**:

```text
L0 ──► L1 ──► L2 ──► L3
hot                  cold
```

| Level | Name | Temperature | Location | Lifespan |
|-------|------|-------------|----------|----------|
| **L0** | Scratchpad | hot / raw | `<notes-dir>/L0_*.md` (git-ignored) | ephemeral — triage soon or delete |
| **L1** | Task memory | warm / working | `<notes-dir>/L1_*.md` | lives per task, cleaned at task end |
| **L2** | Module memory | cool / stable | `L2_*.md` beside the code | durable, module-scoped |
| **L3** | Global memory | cold / curated | `<docs-dir>/` | durable, cross-cutting |

**Authority order** (highest wins on conflict):

```text
L3 > L2 > L1 > L0 > Chat
```

Chat is useful but never authoritative until validated against files. When docs and code disagree, trust the running code first, then propose fixing the docs.

---

## Promotion & cleanup rules

**Promote upward only when knowledge is:** validated by usage or current code · reusable beyond one case · stable enough to survive a reset · not contradicted by higher-authority docs or implementation.

**Routing:** single-task detail → `L1` · module-specific durable lesson → `L2` · cross-cutting project truth → `L3`.

**Cleanup actions (pick one per target):** `keep · merge · summarize · promote · archive · delete · reset`.

Prefer cleanup over accumulation. **Never delete useful anti-pattern knowledge** — relocate it to L2/L3 as a "known failure" instead.

---

## The three verbs

### CAPTURE — save volatile context (capture-only)

Trigger words: *save, checkpoint, before reset, dump context, handoff.*

1. Append (do not overwrite) a raw dump to `<notes-dir>/L0_<topic-or-date>.md`.
2. Use the template below. Capture only — no analysis, no cleanup, no promotion.
3. Preserve contradictions and raw snippets. If context is partial, state it.
4. Confirm the file path and whether it was created or appended.

```md
## Checkpoint — <timestamp>
### Topic
<short description>
### Current state
- what is being worked on / current objective
### Recent actions
- key prompts, commands, outputs, errors
### Observations
- what works / what does not / surprises
### Hypotheses
- possible explanations being tested
### Next steps
- immediate planned actions
### Raw snippets (optional)
- code, errors, partial output (keep raw)
```

### RECALL — retrieve what we already know (read-only)

Trigger words: *recall, what do we know, context for, analyze memory, catch me up.*

1. **Inventory** sources for the topic, in order: chat → `<notes-dir>/L0_*` → `L1_*` → `L2_*` beside code → `<docs-dir>/` → the always-loaded instructions → current code needed to validate claims.
2. State which sources were unavailable; lower confidence accordingly.
3. **Return the briefing** (do not edit any files):
   - **Topic**
   - **Context used** (sources read)
   - **What we already know** (confirmed facts vs. uncertainty — keep them separate)
   - **Prior work and memory**
   - **Failed-but-tested paths**
   - **Conflicts and drift** (surface them; do not flatten)
   - **Recommended starting context**
   - **Suggested next step**

### CURATE — clean and promote memory (approval-first)

Trigger words: *clean up, optimize memory, promote, triage L0, tidy notes.*

1. **Inventory** memory across L0–L3 (as in RECALL).
2. **Detect problems:** drift between chat and repo, duplication across levels, stale assumptions, failed branches polluting L1, oversized noise, missing anti-patterns.
3. **Propose a plan.** For each target list: file · level · action · reason · expected benefit · risk if skipped. Include a reset recommendation (`reset now` / `later` / `none`) with a ≤15-line carry-forward summary if resetting.
4. **STOP for approval.** Do not edit, delete, archive, promote, replace, or reset until the user approves.
5. **Execute only the approved subset.** Then report: actions executed · files changed · actions skipped · per-level summary · follow-ups.

---

## Safety rules

- **RECALL is read-only.** It never edits files.
- **CURATE is approval-first.** It always stops before any destructive change.
- **CAPTURE is capture-only.** No analysis or promotion.
- **Preserve anti-patterns** — relocate, never silently delete.
- **Never overwrite higher-authority memory** with lower-authority claims.
- **Declare unavailable sources** and lower confidence; never fabricate coverage.

---

## Quick reference

| Request | Verb | Writes? |
|---------|------|---------|
| "Save this before we reset" | CAPTURE | appends L0 |
| "What do we know about X?" | RECALL | no |
| "Clean up / promote the notes" | CURATE | only after approval |

**Lifecycle in one line:** capture raw in **L0** → promote working signal to **L1** → keep durable local lessons in **L2** → publish cross-cutting truth in **L3**, deleting low-signal material at every step.
