---
description: Capture volatile working context into an L0 scratchpad (capture-only).
argument-hint: '<topic>'
name: 'mem-capture'
---

# mem-capture

Follow the **CAPTURE** verb in `memory-kit/MEMORY.md`.

- Topic or filename hint: `${input:topic}`

Append a raw dump to `<notes-dir>/L0_<topic-or-date>.md` using the checkpoint template.
Capture only — no analysis, no cleanup, no promotion. Confirm the file path when done.
