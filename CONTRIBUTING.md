# Contributing to WordVim

Thanks for your interest in contributing!

## Getting Started

1. Fork the repository
2. Clone your fork
3. Create a feature branch
4. Make your changes
5. Submit a pull request

## Development Setup

### Requirements

- Windows 10/11
- .NET SDK 8.0+
- Git
- Microsoft Word 365 (for testing)

### Building

```bash
git clone https://github.com/YOUR_USERNAME/wordvim.git
cd wordvim/poc
dotnet build
```

### Testing

1. Register the add-in: `.\register.ps1` (run PowerShell as administrator)
2. Open Microsoft Word
3. Verify `-- NORMAL --` appears in the title bar
4. Press `i` to switch to Insert mode
5. Press `Escape` to switch back to Normal mode
6. Unregister when done: `.\unregister.ps1`

## Code Style

| Element | Convention | Example |
|---------|-----------|---------|
| Namespace | `WordVim.{Layer}` | `WordVim.Core` |
| Class | PascalCase | `VimState` |
| Interface | I + PascalCase | `IWordAdapter` |
| Method | PascalCase | `ProcessKey` |
| Private field | _camelCase | `_currentMode` |
| Parameter | camelCase | `keyCode` |

## Commit Convention

```
type(scope): description
```

Types:
- `feat` — new feature
- `fix` — bug fix
- `docs` — documentation only
- `refactor` — code change that neither fixes a bug nor adds a feature
- `test` — adding or updating tests
- `chore` — build, CI, or tooling changes

Examples:
```
feat(core): add w/b/e word motions
fix(hooks): prevent delegate GC collection
docs: update milestones with Phase 3
```

## Project Structure

```
wordvim/
├── src/
│   ├── WordVim.Core/           # Vim logic — NO Word dependency
│   ├── WordVim.Hooks/          # Win32 keyboard hook
│   ├── WordVim.Word/           # Word COM adapter
│   └── WordVim.Addin/          # COM Add-in entry point
├── tests/
│   ├── WordVim.Core.Tests/     # Unit tests
│   └── WordVim.Integration.Tests/
├── scripts/
├── docs/
└── installer/
```

## Pull Request Guidelines

- Keep PRs small and focused
- One logical change per PR
- Each PR should compile and pass tests
- Write clear commit messages
- Update documentation if needed

## Questions?

Open an issue on GitHub.
