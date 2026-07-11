# WordVim

Vim keybindings for Microsoft Word.

> **Status:** Early development — proof of concept  
> **Platform:** Microsoft Word 365 on Windows

---

## What is this?

WordVim brings Vim's modal editing to Microsoft Word. Press `Escape` for Normal mode, `i` for Insert mode — just like Vim.

### Current Features

- Insert and Normal modes
- Mode shown in window title (`-- NORMAL --` / `-- INSERT --`)
- Normal mode swallows all keys except `i`

### Planned

- `h`/`j`/`k`/`l` movement
- `w`, `b`, `e` word motions
- `x`, `dd`, `yy` commands
- Visual mode
- Command-line mode (`:`)

---

## Installation (super easy)

1. Download [wordvim-v0.1.0.zip](https://github.com/spidychoipro/wordvim/releases/download/v0.1.0/wordvim-v0.1.0.zip)
2. Extract the ZIP
3. Right-click `install.bat` → **Run as administrator**
4. Done! Open Word → you'll see `-- NORMAL --` in the title bar

### Uninstall

Right-click `uninstall.bat` → **Run as administrator**

---

## Usage

| Key | Mode | Action |
|-----|------|--------|
| `Escape` | Insert | → Normal mode |
| `i` | Normal | → Insert mode |
| any other key | Normal | swallowed |

Starts in Normal mode. Press `i` to type.

---

## Build from source

```bash
git clone https://github.com/spidychoipro/wordvim.git
cd wordvim/poc
dotnet build
.\register.ps1  # run as admin
```

Requires: [Git](https://git-scm.com/), [.NET SDK 8.0+](https://dotnet.microsoft.com/download)

---

## How it works

- **COM Add-in** using `IDTExtensibility2` (not VSTO)
- **WH_KEYBOARD** hook intercepts keystrokes
- **Word COM API** for cursor/text control

Office.js can't intercept keystrokes — [architectural limitation](https://github.com/nickebbitt/office-dev-journey/issues/6454).

---

## License

[MIT](LICENSE)
