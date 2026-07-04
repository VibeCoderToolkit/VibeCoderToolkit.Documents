# Contributing to VibeCoderToolkit

**You solved an edge case that kept sending your AI in loops? Ship it.**

VibeCoderToolkit is a community of vibe coders who've been burned by the same gotchas — encrypted files that won't open, CSV quoting that breaks on row 10,000, JSON that silently deserializes wrong. We turn those scars into zero-dependency libraries so the next AI (yours or someone else's) gets it right first try.

## The rules

1. **Zero dependencies.** If your solution needs a NuGet package, it doesn't belong here. BCL only.
2. **AI-first API.** One method call, auto-detection where possible. If an AI needs to read docs to use it, simplify it.
3. **AiExample attributes.** Every model property gets `[AiExample("value")]` so LLMs can inspect intent via reflection.
4. **Cross-platform.** Your code should work on Windows, macOS, and Linux.
5. **Tests.** Include tests that prove your edge case is actually solved. Show the before (the trap) and the after (the fix).

## How to contribute

1. **Fork** the repo you want to add to (or create a new one under your fork)
2. **Add your solution** — a single focused library that does one thing well
3. **Open a PR** — tell us which edge case you solved and how many loops it cost you

## What we're looking for

- Encrypted document handling (Office, PDF, etc.)
- Encoding and charset nightmares
- Date/time parsing across locales and formats
- CSV/TSV/DSV edge cases (embedded quotes, multiline fields, BOM)
- File format detection that actually works
- Schema inference from messy data
- Anything that made you say "WHY IS THIS SO HARD" at 2am

## Org-level repos

If your solution is big enough to be its own library (not a fit for an existing repo), we'll create a new repo under the org. PRs for new repo proposals welcome — just open an issue describing the edge case and your approach.

---

**You fought the edge case and won. Let the next AI benefit from your battle scars.**
