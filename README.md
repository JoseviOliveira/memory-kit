# Human-written introduction

The idea raised naturally while building a web app from scratch using Codex and Claude's AI agents.
I realised that models focused too hard on creating new code but not so much on thinking-first and gathering knowledge before taking decisions.

The primary reasons were to:
- 1st: Window chat session: it grows fast, gets dirty quickly while iterating, and introduces noise into the development process. AI compacting can reduce the size but it doesn't mean it keeps ONLY the relevant context.
- 2nd: Move important information out from chat windows and keep it organized in a durable memory system.
- 3rd: Guide -or better, "force"- models to build a structured multi-level knowledge system, including immediate context, module-specific knowledge, and long-term documentation.
- 4th: Ensure that the captured knowledge is consistently used and integrated into the development workflow.
- 5th: Last but not least: Token economy.

This small framework organise technical and functional information from hot (L0) to cold (L3):
- Lowest levels capture immediate context and transient information: It keeps what didn't work well during dev iterations, and what did work well.
- Medium levels keep concern for a module or specific part of the app.
- The highest level builds a long-term, durable knowledge documentation.
- All levels must be used by AI agents to ensure comprehensive context awareness and knowledge integration.

Format: All levels stand as Markdown files except the highest one, which can be built in a kinder format such as HTML. 

This memory kit becomes a good interface for exchange between AI agents and humans.
**Provocative final thought**: Why spend time reviewing code when you can validate knowledge directly from the memory kit?

-- The rest of the content in this repository is fully agentic-AI generated.

# memory-kit

A portable, four-level memory convention (**L0 → L1 → L2 → L3**, hot → cold) for any coding project. Capture context fast, recall what you already know, and curate durable lessons — without losing signal to context resets.

The entire system is one file: [`MEMORY.md`](MEMORY.md). The `prompts/` folder is optional slash-command sugar.

## Install (any project, ~1 minute)

1. Copy this `memory-kit/` folder into the project (root or `.github/`).
2. Add one line to the project's always-loaded instructions (`AGENTS.md`, `.github/copilot-instructions.md`, `CLAUDE.md`, or `.cursorrules`):
   > Follow the memory convention in `memory-kit/MEMORY.md`.
3. Open `MEMORY.md` and set `<notes-dir>` and `<docs-dir>` (the only project-specific edit). Git-ignore `L0_*`:
   ```gitignore
   notes/L0_*.md
   ```

No build step, no dependencies, no tool-specific runtime.

## Use

Plain language works in any agent (Copilot, Claude Code, Cursor, Aider, …):

- "Save the current context before we reset." → **CAPTURE**
- "What do we already know about the auth flow?" → **RECALL**
- "Clean up and promote the useful notes." → **CURATE** (asks for approval first)

If your tool supports prompt files, the optional shortcuts in `prompts/` route to the same verbs:

- `/mem-capture <topic>`
- `/mem-recall <topic>`
- `/mem-curate <scope>`

## Why it's small

The model is defined **once** (in `MEMORY.md`). The three verbs are inline procedures, not a chain of agents/skills/hooks. One source of truth → no drift, full portability.
