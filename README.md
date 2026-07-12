<div align="center">

# ⌨️ WordVim

**Vim keybindings for Microsoft Word**

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-blue.svg)]()
[![Word](https://img.shields.io/badge/Word-365%20Desktop-orange.svg)]()
[![.NET](https://img.shields.io/badge/.NET-Framework%204.7.2-purple.svg)]()
[![Status](https://img.shields.io/badge/Status-Proof%20of%20Concept-brightgreen.svg)]()
[![Release](https://img.shields.io/github/v/release/spidychoipro/wordvim.svg)](https://github.com/spidychoipro/wordvim/releases)

**English** · [한국어](README.ko.md) · [日本語](README.ja.md) · [中文](README.zh.md)

---

**WordVim** brings Vim's modal editing experience to Microsoft Word.

Press `Escape` to enter Normal mode, `i` to return to Insert mode — just like in Vim.

**No Visual Studio required.** Builds with `dotnet build`. One-click install.

[**⬇️ Download v0.1.0**](https://github.com/spidychoipro/wordvim/releases/tag/v0.1.0) · [**📖 Documentation**](#documentation) · [**🚀 Quick Start**](#quick-start)

---

</div>

## 🎯 Why WordVim?

Microsoft Word lacks Vim's powerful modal editing. **Office.js cannot intercept keystrokes** — this is an [architectural limitation](https://github.com/nickebbitt/office-dev-journey/issues/6454) that cannot be worked around.

WordVim solves this with **native COM Add-in + Win32 keyboard hooks** — the only approach that enables true Vim modal editing in Word.

<div align="center">

```
┌─────────────────────────────────────────────────────────┐
│  Word: -- NORMAL --                                     │
│                                                         │
│  vim → Word integration                                │
│                                                         │
│  ✅ Modal editing (Insert/Normal/Visual)                │
│  ✅ Keyboard interception via Win32 hooks               │
│  ✅ Cursor movement via Word COM API                    │
│  ✅ One-click install — no Visual Studio needed         │
└─────────────────────────────────────────────────────────┘
```

</div>

## ✨ Features

| Feature | Status |
|---------|--------|
| Insert / Normal mode switching | ✅ Done |
| Mode indicator (title bar) | ✅ Done |
| Key suppression in Normal mode | ✅ Done |
| `h`/`j`/`k`/`l` cursor movement | 🔜 Next |
| `w`/`b`/`e` word motions | 🔜 Planned |
| `x`/`dd`/`yy` editing commands | 🔜 Planned |
| Visual mode | 🔜 Planned |
| Command-line mode (`:`) | 🔜 Planned |

## 🚀 Quick Start

### Option 1: Download (Recommended)

> **Step 1:** Download [`wordvim-v0.1.0.zip`](https://github.com/spidychoipro/wordvim/releases/download/v0.1.0/wordvim-v0.1.0.zip)

> **Step 2:** Extract the ZIP file anywhere

> **Step 3:** Right-click `install.bat` → **Run as administrator**

> **Step 4:** Open Microsoft Word → See `-- NORMAL --` in title bar 🎉

```
📁 wordvim-v0.1.0.zip
   └── 📁 wordvim
       ├── 📄 install.bat      ← Double-click this (as admin)
       ├── 📄 uninstall.bat
       ├── 📄 register.ps1
       ├── 📄 unregister.ps1
       ├── 📄 PoC.dll
       └── 📄 PoC.pdb
```

### Option 2: Build from Source

```bash
# Prerequisites: .NET SDK 8.0+ and Git
git clone https://github.com/spidychoipro/wordvim.git
cd wordvim/poc
dotnet build

# Register (run PowerShell as administrator)
.\register.ps1
```

## 📖 Documentation

| Document | Description |
|----------|-------------|
| [Architecture Analysis](docs/architecture.md) | Why COM Add-in + Win32 hooks, not VSTO or Office.js |
| [API Research](docs/research.md) | Word COM API capabilities and limitations |
| [MVP Specification](docs/mvp.md) | What the minimum viable product includes |
| [Milestones](docs/milestones.md) | 30-day development roadmap |
| [PoC Plan](docs/poc-plan.md) | Proof of concept validation plan |
| [Environment Guide](docs/environment.md) | Development environment setup |
| [Contributing](AGENTS.md) | Contributor guidelines and code conventions |

## 🏗️ Architecture

```
┌──────────────────────────────────────────────────────────────┐
│                      WordVim Architecture                    │
├──────────────────────────────────────────────────────────────┤
│                                                              │
│  ┌─────────────┐    ┌─────────────┐    ┌─────────────┐     │
│  │   Win32     │    │   Word      │    │   Vim       │     │
│  │   Hook      │───▶│   COM API   │───▶│   State     │     │
│  │  (keyboard) │    │  (cursor)   │    │   Machine   │     │
│  └─────────────┘    └─────────────┘    └─────────────┘     │
│         │                  │                   │             │
│         ▼                  ▼                   ▼             │
│  ┌─────────────────────────────────────────────────────┐    │
│  │              COM Add-in (IDTExtensibility2)         │    │
│  │                  WordVim.Connect                    │    │
│  └─────────────────────────────────────────────────────┘    │
│                                                              │
└──────────────────────────────────────────────────────────────┘
```

| Component | Technology | Purpose |
|-----------|------------|---------|
| **Add-in** | COM + IDTExtensibility2 | Load into Word process |
| **Keyboard** | WH_KEYBOARD hook | Intercept keystrokes before Word |
| **Cursor** | Word COM API | Move cursor, edit text |
| **State** | C# state machine | Track Insert/Normal mode |

## 🎮 Usage

### Key Bindings

| Key | Mode | Action |
|-----|------|--------|
| `Escape` | Insert → Normal | Switch to Normal mode |
| `i` | Normal → Insert | Switch to Insert mode |
| `h`/`j`/`k`/`l` | Normal | Cursor movement *(planned)* |
| `w`/`b`/`e` | Normal | Word motions *(planned)* |
| `x` | Normal | Delete character *(planned)* |
| `dd` | Normal | Delete line *(planned)* |
| `yy` | Normal | Yank line *(planned)* |

### Mode Indicator

The window title bar displays the current mode:

```
-- NORMAL --  ← All keys except 'i' are swallowed
-- INSERT --  ← Normal typing mode
```

## ❓ FAQ

<details>
<summary><b>Does this work with Word Online?</b></summary>
<br>
No. WordVim requires the desktop version of Word 365 on Windows. Word Online uses a different architecture that doesn't support COM Add-ins.
</details>

<details>
<summary><b>Do I need Visual Studio?</b></summary>
<br>
No. WordVim builds with `dotnet build` — just install the .NET SDK.
</details>

<details>
<summary><b>Is this safe?</b></summary>
<br>
Yes. WordVim is open source (MIT License). The installer only registers a COM Add-in — no system files are modified.
</details>

<details>
<summary><b>How do I uninstall?</b></summary>
<br>
Right-click `uninstall.bat` → Run as administrator. Done.
</details>

## 🤝 Contributing

See [AGENTS.md](AGENTS.md) for contributor guidelines.

```bash
# Commit convention
feat(scope): description
fix(scope): description
docs: description

# Examples
feat(core): add w/b/e word motions
fix(hooks): prevent delegate GC collection
docs: update milestones with Phase 3
```

## 📊 Project Status

```
✅ Research & Architecture      ████████████ 100%
✅ PoC: Add-in Loading          ████████████ 100%
✅ PoC: Keyboard Hook           ████████████ 100%
✅ PoC: Mode Switching          ████████████ 100%
🔜 Phase 0: Foundation          ░░░░░░░░░░░░   0%
🔜 Phase 1: Core Engine         ░░░░░░░░░░░░   0%
🔜 Phase 2: Word Integration    ░░░░░░░░░░░░   0%
🔜 Phase 3: Commands            ░░░░░░░░░░░░   0%
🔜 Phase 4: Visual Mode         ░░░░░░░░░░░░   0%
🔜 Phase 5: Release v1.0        ░░░░░░░░░░░░   0%
```

## 📄 License

[MIT License](LICENSE) — free to use, modify, and distribute.

## 🙏 Acknowledgments

- [Vim](https://www.vim.org/) — the legendary text editor that inspired this project
- [Word Object Model](https://learn.microsoft.com/en-us/office/vba/api/word.word) — Microsoft's COM API for Word automation
- [WH_KEYBOARD Hook](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setwindowshookexw) — Win32 API for keyboard interception

---

<div align="center">

**Built with ❤️ for Vim users who are stuck with Word**

[![Buy Me A Coffee](https://img.shields.io/badge/Buy%20Me%20A%20Coffee-ffdd00?style=flat&logo=buy-me-a-coffee&logoColor=black)](https://www.buymeacoffee.com/spidychoi)

[⬆️ Back to top](#-wordvim)

</div>
