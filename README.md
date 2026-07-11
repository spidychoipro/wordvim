# WordVim

Vim keybindings for Microsoft Word.

> **Status:** Early development — proof of concept stage.  
> **Platform:** Microsoft Word 365 on Windows (.NET Framework 4.7.2)

---

## What is WordVim?

WordVim brings Vim's modal editing experience to Microsoft Word. Press `Escape` to enter Normal mode, `i` to return to Insert mode — just like in Vim.

### Current Features

- **Modal editing** — Insert and Normal modes with automatic switching
- **Mode indicator** — Window title shows current mode (`-- NORMAL --` / `-- INSERT --`)
- **Key suppression** — Normal mode swallows all keys except recognized commands

### Planned Features

- `h`/`j`/`k`/`l` cursor movement
- `w`, `b`, `e` word motions
- `x`, `dd`, `yy` editing commands
- Visual mode
- Command-line mode (`:`)
- And more...

---

## Installation

### Prerequisites

- Windows 10/11
- Microsoft Word 365 (desktop version)
- .NET Framework 4.7.2 or later (pre-installed on Windows 10 1803+)

### Quick Install

1. Download the latest release from [Releases](https://github.com/spidychoipro/wordvim/releases)
2. Extract the ZIP file
3. Right-click `install.bat` and select **Run as administrator**
4. Done! Open Microsoft Word — you'll see `-- NORMAL --` in the title bar

### Uninstall

1. Right-click `uninstall.bat` and select **Run as administrator**
2. Done! WordVim is removed

---

## Usage

| Key | Mode | Action |
|-----|------|--------|
| `Escape` | Insert | Switch to Normal mode |
| `i` | Normal | Switch to Insert mode |
| *any other key* | Normal | Ignored (swallowed) |

**Starting in Normal mode** — when Word opens, WordVim starts in Normal mode. Press `i` to begin typing.

---

## Building from Source

### Requirements

- [Git](https://git-scm.com/)
- [.NET SDK 8.0+](https://dotnet.microsoft.com/download)

### Steps

```bash
# Clone the repository
git clone https://github.com/spidychoipro/wordvim.git
cd wordvim

# Build
cd poc
dotnet build

# Register the add-in (run as administrator)
.\register.ps1
```

### Project Structure

```
wordvim/
├── poc/                    # Proof of concept (current)
│   ├── Connect.cs          # COM Add-in entry point
│   ├── PoC.csproj          # Project file
│   ├── register.ps1        # Registration script
│   └── unregister.ps1      # Unregistration script
├── docs/                   # Design documents
│   ├── architecture.md     # Architecture analysis
│   ├── research.md         # API research
│   ├── milestones.md       # Development roadmap
│   └── mvp.md              # MVP specification
├── AGENTS.md               # Contributor guidelines
└── LICENSE                 # MIT License
```

---

## How It Works

WordVim uses a **native COM Add-in** (not VSTO) with:

1. **`IDTExtensibility2`** — COM Add-in interface for loading into Word
2. **`WH_KEYBOARD` hook** — Thread-level keyboard hook to intercept keystrokes
3. **Word COM API** — `Application.Selection` for cursor movement and text manipulation

This is the only approach that enables true Vim modal editing in Word. Office.js (Microsoft's modern add-in platform) cannot intercept document keystrokes — an [architectural limitation](https://github.com/nickebbitt/office-dev-journey/issues/6454) that cannot be worked around.

---

## Contributing

See [AGENTS.md](AGENTS.md) for contributor guidelines.

### Commit Convention

```
type(scope): description

feat(core): add w/b/e word motions
fix(hooks): prevent delegate GC collection
docs: update milestones with Phase 3
```

---

## License

[MIT](LICENSE)

---

## Acknowledgments

- Inspired by [Vim](https://www.vim.org/) — the legendary text editor
- Built with the [Word Object Model](https://learn.microsoft.com/en-us/office/vba/api/word.word)
- Keyboard hooks via Win32 API
