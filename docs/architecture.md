# WordVim Architecture

## 1. Problem Analysis

### What we're building

A Vim keybinding extension for Microsoft Word that enables modal editing (Normal, Visual, Insert, Command-line modes) within the Word document editor.

### The fundamental challenge

Vim requires intercepting **every keystroke** before the application processes it, and then **suppressing** the default behavior when the keystroke is a Vim command. This is the single hardest technical constraint of the project.

Microsoft Word processes keystrokes in its native process. Extension add-ins run in sandboxed runtimes. The keystroke event never reaches the add-in before Word has already handled it.

### Core requirements for Vim emulation

| Requirement | Why it matters |
|---|---|
| Intercept all keystrokes before Word processes them | Normal mode must capture `d`, `y`, `x`, motions, etc. |
| Suppress default keystroke behavior | Pressing `d` in Normal mode must NOT insert the letter "d" |
| Move cursor programmatically | `h/j/k/l`, `w/b/e`, `0/$`, `gg/G` |
| Read and manipulate selection text | `diw`, `ci"`, `yl`, visual mode selections |
| Read/write document content | `dd` (delete line), `yy` (yank line), `p` (paste) |
| Display current mode | User must see whether they're in Normal/Insert/Visual mode |

---

## 2. Extension Mechanism Comparison

### Option A: VSTO + Win32 Hooks (C# / .NET)

How it works: A C# add-in runs inside the Word process. It installs a Win32 keyboard hook via `SetWindowsHookEx(WH_KEYBOARD, ...)` to intercept keystrokes at the thread level.

| Criterion | Assessment |
|---|---|
| Keystroke interception | **Full** — intercepts every keystroke before Word |
| Keystroke suppression | **Full** — return non-zero from hook callback |
| Cursor control | **Full** — Word Object Model: `Selection.Move`, `SetRange`, etc. |
| Text manipulation | **Full** — `Selection.Text`, `Range.Text`, full read/write |
| Deployment | MSI/ClickOnce installer, requires .NET Framework 4.8 |
| Cross-platform | **Windows only** |
| Long-term outlook | **Maintenance mode** — Microsoft recommends Office.js migration |

**Pros:** Only approach that achieves true modal editing with full Vim fidelity.
**Cons:** Windows-only, declining platform, complex hook management, crash risk (bad hook = Word crash).

### Option B: VBA Macros

How it works: VBA `KeyBindings.Add` maps specific key combinations to macros. The macros use the Word Object Model for cursor and text manipulation.

> **Critical correction:** `Application.OnKey` does **NOT exist in Word VBA**. It is Excel-only ([Word docs return 404](https://learn.microsoft.com/en-us/office/vba/api/word.application.onkey)). Word VBA has no equivalent method for intercepting arbitrary keystrokes.

| Criterion | Assessment |
|---|---|
| Keystroke interception | **No** — `KeyBindings.Add` only reassigns registered combos (same as Tools > Customize Keyboard). Cannot intercept arbitrary keystrokes. |
| Keystroke suppression | **No** — no `OnKey` with empty string equivalent. `KeyBinding.Disable` exists but only for registered bindings. |
| Cursor control | **Full** — same Word Object Model as VSTO |
| Text manipulation | **Full** — same Word Object Model as VSTO |
| Deployment | `.dotm` template in Startup folder, trivial distribution |
| Cross-platform | **Windows + Mac** (VBA built into both desktop apps) |
| Long-term outlook | **Legacy but stable** — VBA is not going away |

**Pros:** Simplest deployment, cross-platform desktop, no external dependencies.
**Cons:** No keystroke interception mechanism. `KeyBindings.Add` only handles registered combos. True modal editing is impossible. A helper DLL for Win32 hooks would be needed, which defeats VBA's simplicity advantage.

### Option C: Office.js (Web Add-in)

How it works: A JavaScript/TypeScript add-in runs in a sandboxed web runtime. Can define custom keyboard shortcuts via the manifest and read/write document content via the Word JavaScript API.

| Criterion | Assessment |
|---|---|
| Keystroke interception | **None** — Word intercepts all keyboard events before the add-in |
| Keystroke suppression | **None** — cannot prevent Word from handling keystrokes |
| Cursor control | **Severely limited** — no cursor position API, `range.select()` only |
| Text manipulation | **Good** — `document.getSelection()`, `range.insertText()`, etc. |
| Deployment | Microsoft AppSource or sideload, HTTPS-hosted web files |
| Cross-platform | **Windows, Mac, Web** |
| Long-term outlook | **Active development** — Microsoft's strategic direction |

**Pros:** Cross-platform including web, modern technology, active Microsoft support.
**Cons:** Fundamentally cannot achieve Vim modal editing. Keyboard events never reach the add-in. This is an architectural limitation, not a gap to be filled.

### Option D: Hybrid VBA + Office.js

How it works: VBA handles keyboard hooks and cursor manipulation (the things Office.js cannot do). Office.js provides a task pane UI for mode display, configuration, and visual feedback.

| Criterion | Assessment |
|---|---|
| Keystroke interception | **Partial** — same VBA limitations, or Win32 hooks via helper DLL |
| Keystroke suppression | **Partial** — same VBA limitations |
| Cursor control | **Full** — VBA Word Object Model |
| Text manipulation | **Full** — VBA + Office.js |
| Deployment | `.dotm` template + Office.js manifest, complex dual install |
| Cross-platform | **Windows + Mac** (VBA portion), Web (Office.js UI only) |
| Long-term outlook | **Mixed** — VBA stable, Office.js active |

**Pros:** Modern UI layer, VBA for the hard parts.
**Cons:** Two technology stacks to maintain, complex installation, VBA still cannot do true modal editing.

---

## 3. Existing Projects

| Project | Approach | Result |
|---|---|---|
| [VimWord](https://github.com/cxw42/VimWord) | VBA with `KeyBindings.Add` | Hotkey-triggered command mode, not true modal editing |
| [Vimifier](https://github.com/strictlymike/vimifier) | Win32 low-level hooks (any app) | System-wide, translates to standard edit shortcuts, limited rich-text |
| [windows-vim-mode](https://github.com/vieuxfontsize/windows-vim-mode) | AutoHotkey | System-wide, limited rich-text support |

**No project achieves true modal Vim editing in Word.** The closest is Vimifier, but it operates system-wide and cannot leverage Word's rich-text API.

---

## 4. Recommendation

### Recommended Architecture: VSTO with Win32 Hooks (C# / .NET)

This is the **only approach that can achieve the project's stated goal** of Vim keybindings in Word. The alternatives fundamentally cannot intercept keystrokes.

### Why not Office.js?

Office.js is the future, but it has a hard technical limitation: **it cannot intercept keyboard events in the Word document.** The document keystroke pipeline is entirely controlled by Word's native process. No amount of clever engineering can work around this. Microsoft has not announced plans to expose this capability (see [office-js#6454](https://github.com/OfficeDev/office-js/issues/6454)).

Choosing Office.js would mean accepting that true modal editing is impossible, and the result would be a hotkey-triggered command palette (like VimWord) rather than a Vim emulation.

### Why not VBA?

VBA has **no keystroke interception mechanism in Word**. `Application.OnKey` does not exist in Word VBA (it is Excel-only). `KeyBindings.Add` only reassigns specific key combinations to commands — it cannot intercept arbitrary keystrokes needed for Normal mode (`dw`, `ciw`, `f{char}`, etc.). To intercept arbitrary keystrokes from VBA, you would need a helper DLL that provides Win32 hooks, which defeats the purpose of VBA's simplicity advantage.

### Architecture design

```
WordVim
├── src/
│   ├── WordVim.Core/           # Vim state machine, command parser
│   │   ├── VimState.cs          # Mode tracking (Normal, Insert, Visual, Command)
│   │   ├── KeyHandler.cs        # Keystroke → action mapping
│   │   ├── CommandParser.cs     # Parse Vim commands (operator + motion + count)
│   │   ├── Motions.cs           # h/j/k/l, w/b/e, 0/$, gg/G, f/t
│   │   ├── Operators.cs         # d/y/c/x/p/indent
│   │   └── Registers.cs         # Register system ("a-"z, "0-"9, etc.)
│   │
│   ├── WordVim.Hooks/          # Win32 keyboard hook management
│   │   ├── KeyboardHook.cs      # SetWindowsHookEx wrapper
│   │   └── HookInstaller.cs     # Install/uninstall lifecycle
│   │
│   ├── WordVim.Word/           # Word Object Model interface
│   │   ├── SelectionHelper.cs   # Cursor movement, selection read/write
│   │   ├── DocumentHelper.cs    # Document content manipulation
│   │   └── WordBridge.cs        # Adapter between VimCore and Word COM
│   │
│   └── WordVim.Addin/          # VSTO entry point
│       ├── ThisAddIn.cs         # Add-in lifecycle
│       ├── Ribbon.cs            # Ribbon UI (mode indicator, settings)
│       └── StatusBar.cs         # Mode display in Word status bar
│
├── tests/
│   ├── WordVim.Core.Tests/      # Unit tests for state machine, parser
│   └── WordVim.Integration.Tests/ # Word COM integration tests
│
└── installer/
    └── Setup.iss                # Inno Setup or WiX installer script
```

### Key design decisions

1. **Clean separation between Vim logic and Word API.** The `Core` project knows nothing about Word. It receives keystrokes and returns actions. This makes the state machine testable without Word.

2. **WordBridge adapter pattern.** All Word COM interaction goes through an interface (`IWordAdapter`). This isolates the crash-prone COM calls and makes it possible to swap the Word backend if Microsoft ever exposes keyboard events in Office.js.

3. **Hook management in a separate assembly.** Win32 hooks are dangerous (bad hook = application crash). Isolating this code reduces blast radius and makes it independently testable.

4. **Ribbon-based mode indicator.** Use the VSTO Ribbon designer to show current mode in the Word Ribbon. Also write to the Word status bar (`Application.StatusBar`).

5. **Installer required.** VSTO add-ins require installation (registry entries, CAS policy). Include a simple installer from day one.

### Risk mitigation

| Risk | Mitigation |
|---|---|
| Hook crashes Word | Hook callback must be minimal (< 1ms). Post message to main thread for processing. |
| Word COM threading | All COM calls on main thread only. Use `System.Windows.Forms.BeginInvoke`. |
| User has conflicting add-ins | Graceful fallback: detect hook failure, disable WordVim, show error. |
| Word update breaks COM | Pin to documented COM interfaces. Avoid undocumented APIs. |
| Windows-only frustration | Document clearly in README. Maintain an issue for "Office.js when possible" as a long-term goal. |

---

## 5. Development Roadmap

### Phase 0: Foundation (Week 1-2)

**Goal:** Skeleton project that compiles and loads into Word.

- Initialize .NET Framework 4.8 class library project
- Set up VSTO project with Ribbon UI
- Create basic installer (Inno Setup)
- Verify add-in loads in Word without errors
- Display "WordVim loaded" in status bar
- Set up CI (GitHub Actions: build + unit test)

**Deliverable:** Empty add-in that loads and shows status.

### Phase 1: Core Engine (Week 3-5)

**Goal:** Vim state machine and command parser, fully testable without Word.

- Implement `VimState` (Normal/Insert/Visual modes)
- Implement key mapping table for Normal mode motions: `h/j/k/l`, `w/b/e`, `0/$`, `gg/G`
- Implement `CommandParser`: operator + motion + count composition (`d2w`, `3j`, `ci"`)
- Implement operators: `d` (delete), `y` (yank), `c` (change), `x` (delete char)
- Implement Visual mode: character and line selection
- Unit tests for all core logic (no Word dependency)

**Deliverable:** All core logic with 90%+ test coverage, no Word dependency.

### Phase 2: Word Integration (Week 6-8)

**Goal:** Connect the engine to Word via Win32 hooks and COM.

- Implement `KeyboardHook` (WH_KEYBOARD thread-level hook)
- Implement `WordBridge` (cursor movement, text read/write via COM)
- Wire: keystroke → VimState → action → Word COM call
- Basic Normal mode working: `h/j/k/l` movement, `i/a/o` enter insert, `dd/yy/p`
- Mode indicator in Ribbon and status bar

**Deliverable:** Normal mode working in Word for basic editing.

### Phase 3: Command Completion (Week 9-12)

**Goal:** Full Normal mode command set.

- Complete motion repertoire: `f/t/F/T`, `{`/`}`, `%`, `0`/`^`/`$`
- Complete operator repertoire: `>`/`<` (indent), `~` (toggle case)
- Search: `/` and `?` with `n/N`
- Command-line mode (`:`) for `:w`, `:q`, `:wq`, `:e`
- Undo/redo integration (`u`, `Ctrl+r`)
- Count support: `3dd`, `5j`, `2dw`
- Dot repeat (`.`)

**Deliverable:** Feature-complete Normal mode.

### Phase 4: Visual Mode and Polish (Week 13-16)

**Goal:** Visual mode, edge cases, user experience.

- Visual mode: characterwise and linewise
- Visual mode operations: `d`, `y`, `c`, `>`, `<`
- `V` (visual line mode)
- Registers: unnamed, named (`"a`-`"z`), numbered (`"0`-`"9`)
- Paste from register, yank to register
- Search and replace (`:s/pattern/replacement/g`)
- Configuration: key remapping, option toggles
- Settings UI in Ribbon

**Deliverable:** Feature-complete Vim emulation.

### Phase 5: Testing and Release (Week 17-20)

**Goal:** Stability, documentation, v1.0 release.

- Integration testing with Word (automated where possible)
- Manual testing matrix: Word 365 on Windows 10/11
- Performance testing: hook latency < 5ms, no UI lag
- User documentation: installation guide, command reference
- GitHub release with installer
- README with screenshots/GIFs

**Deliverable:** v1.0 release.

---

## 6. Technology Stack

| Component | Technology | Reason |
|---|---|---|
| Core engine | C# (.NET Framework 4.8) | VSTO requirement, strong typing for state machine |
| Word integration | VSTO + Word COM Object Model | Only way to achieve full Vim fidelity |
| Keyboard hooks | Win32 API via P/Invoke | Required for keystroke interception |
| Unit testing | NUnit or xUnit | Standard .NET testing |
| Build | MSBuild / Visual Studio | VSTO project type requires VS |
| Installer | Inno Setup | Simple, free, widely used |
| CI | GitHub Actions | Build + test on push |
| License | MIT | Already decided |

---

## 7. What Would Change If Microsoft Exposes Keyboard Events in Office.js

If Microsoft adds a `DocumentKeyDown` event to the Word JavaScript API (currently not planned), the architecture should be designed to swap the Word backend:

1. Replace `WordVim.Hooks` + `WordVim.Word` with an Office.js adapter
2. `WordVim.Core` remains unchanged (it's already backend-agnostic)
3. Redistribute as Office.js add-in (cross-platform: Windows, Mac, Web)

The `IWordAdapter` interface makes this swap feasible without rewriting the core.

---

## 8. Decision Log

| Decision | Choice | Rationale |
|---|---|---|
| Extension type | VSTO (C#) | Only approach with full keystroke interception |
| Hook type | WH_KEYBOARD (type 2, thread-level) | WH_KEYBOARD_LL (type 13) does NOT fire for keystrokes in the hosting Office app (Word 2013+). WH_KEYBOARD (type 2) is confirmed working. |
| Hook module handle | `IntPtr.Zero` | In-process hook, no DLL injection needed |
| Thread ID | `GetCurrentThreadId()` | `Thread.ManagedThreadId` does NOT work with SetWindowsHookEx |
| Core language | C# .NET 4.8 | VSTO requirement, strong typing |
| Testing | Unit tests (NUnit) | Core is backend-agnostic, testable without Word |
| Installer | Inno Setup | Simple, free, no dependency on WiX Toolset |
| License | MIT | Already decided, permissive for adoption |
| Cross-platform | Windows only (initial) | Technical limitation of the approach |
| VBA OnKey | **Does not exist in Word** | Excel-only. Word VBA has no keystroke interception mechanism. |
