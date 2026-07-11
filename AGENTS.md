# AGENTS.md

> Rules for AI contributors working on WordVim.

---

## Before Writing Any Code

1. **Read `docs/architecture.md`** — understand why the project exists and how it's structured.
2. **Read `docs/mvp.md`** — understand what's been built and what hasn't.
3. **Read the file you're about to modify** — understand its current state before changing it.
4. **Run `dotnet build`** — make sure the project compiles before you touch anything.
5. **Run `dotnet test`** — make sure all tests pass before you start.

---

## Rules

### Think Before Coding

- Understand the problem before reaching for a solution.
- If you're unsure about an approach, say so. Don't guess.
- If the request is ambiguous, ask for clarification before writing code.

### Read Before Modifying

- Always read the full file before editing it.
- Read neighboring files to understand conventions and patterns.
- Check if similar functionality already exists elsewhere in the codebase.

### Keep Commits Small

- One logical change per commit.
- Each commit should compile and pass tests.
- Never mix formatting changes with logic changes.
- Never mix refactoring with feature work.

### Explain Architectural Decisions

- When introducing a new pattern, explain why in the commit message.
- When choosing between two approaches, document the tradeoff.
- Reference `docs/architecture.md` or `docs/research.md` when the decision is documented there.

### Prefer Maintainability Over Cleverness

- Write code that a new contributor can understand.
- Avoid clever one-liners that obscure intent.
- Prefer explicit over implicit.
- Prefer simple names over abbreviated ones.

### Never Refactor Without Justification

- Don't reorganize files, rename classes, or restructure projects unless there's a concrete problem being solved.
- If you think a refactor is needed, open an issue first explaining why.
- Large refactors require explicit approval.

### Follow Conventional Commits

Format: `type(scope): description`

Types:
- `feat` — new feature
- `fix` — bug fix
- `refactor` — code change that neither fixes a bug nor adds a feature
- `test` — adding or updating tests
- `docs` — documentation only
- `chore` — build, CI, or tooling changes
- `style` — formatting, whitespace, no logic change

Scope: `core`, `hooks`, `word`, `addin`, `tests`, `docs`, `scripts`

Examples:
- `feat(core): add w/b/e word motions`
- `fix(hooks): prevent delegate GC collection`
- `test(core): add VimState transition tests`
- `docs: update milestones with Phase 3`

### Keep the Project Beginner-Friendy

- Write clear, descriptive variable and method names.
- Add XML doc comments on public interfaces and methods.
- Don't assume the reader knows Vim internals — explain Vim concepts briefly.
- Don't assume the reader knows Win32 hooks or COM interop — point to relevant docs.

---

## Project Structure

```
wordvim/
├── src/
│   ├── WordVim.Core/           # Vim logic — NO Word dependency
│   ├── WordVim.Hooks/          # Win32 keyboard hook
│   ├── WordVim.Word/           # Word COM adapter
│   └── WordVim.Addin/          # COM Add-in entry point
│
├── tests/
│   ├── WordVim.Core.Tests/     # Unit tests (no Word dependency)
│   └── WordVim.Integration.Tests/
│
├── scripts/                    # Build, test, register scripts
├── docs/                       # Architecture, milestones, research
└── installer/                  # Inno Setup + registry
```

**Key rule:** `WordVim.Core` must NEVER reference Word, COM, or Win32 APIs. It is pure C# logic. If you need to add a Word dependency, it belongs in `WordVim.Word`.

---

## Build & Test

```powershell
dotnet build                  # Build everything
dotnet test                   # Run all unit tests
dotnet test --filter "ClassName=VimStateTests"  # Run specific tests
```

**Every change must pass both `dotnet build` and `dotnet test` before committing.**

---

## Code Conventions

| Element | Convention | Example |
|---------|-----------|---------|
| Namespace | `WordVim.{Layer}` | `WordVim.Core`, `WordVim.Hooks` |
| Class | PascalCase | `VimState`, `KeyHandler` |
| Interface | I + PascalCase | `IWordAdapter` |
| Method | PascalCase | `ProcessKey`, `MoveCursor` |
| Private field | _camelCase | `_currentMode` |
| Parameter | camelCase | `keyCode` |
| Constant | PascalCase | `MaxRegisterCount` |
| Test method | `Method_Scenario_Expected` | `Transition_NormalToInsert_ReturnsPrevious` |

---

## Testing Requirements

- `WordVim.Core` tests: run on every commit, no external dependencies
- Integration tests: run manually, require Word installed
- **Never commit code that breaks existing tests.** Fix the tests first or explain why they need to change.

---

## What NOT to Do

- Don't add dependencies without checking if they already exist in the project.
- Don't use `#pragma warning disable` to hide warnings — fix the warning.
- Don't use `var` when the type isn't obvious from the right side.
- Don't add comments unless explaining non-obvious why (not what).
- Don't rename files or classes without discussing it first.
- Don't change the .NET target framework.
- Don't add NuGet packages to `WordVim.Core` — it must stay lightweight.
