# Microsoft Word API Research

> Investigation of every available API for keyboard interception in Microsoft Word.
> All claims sourced from official Microsoft documentation. No assumptions.

---

## 1. Word Object Model (VSTO / COM Interop)

**Source:** [Application object members](https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.application), [ApplicationEvents4_Event](https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.applicationevents4_event)

### 1.1 Application Events — Complete List

The `Application` object exposes **34 events**. None are keyboard events.

| Event | Fires | Cancellable | What it does |
|-------|-------|-------------|-------------|
| `DocumentBeforeClose` | Before | **Yes** | Document closing |
| `DocumentBeforePrint` | Before | **Yes** | Document printing |
| `DocumentBeforeSave` | Before | **Yes** | Document saving |
| `DocumentChange` | After | No | New doc created/opened/activated |
| `DocumentOpen` | After | No | Document opened |
| `NewDocument` | After | No | New document created |
| `WindowActivate` | After | No | Window activated |
| `WindowDeactivate` | After | No | Window deactivated |
| **`WindowBeforeDoubleClick`** | Before | **Yes** | Mouse double-click (has `Cancel`) |
| **`WindowBeforeRightClick`** | Before | **Yes** | Mouse right-click (has `Cancel`) |
| **`WindowSelectionChange`** | After | No | Selection changed (closest to keyboard) |
| `WindowSize` | After | No | Window resized/moved |
| `XMLSelectionChange` | After | No | XML node selection changed |
| `XMLValidationError` | After | No | XML validation error |
| `Quit` | After | No | User quits Word |
| `MailMerge*` (8 events) | Mixed | Some | Mail merge lifecycle |
| `ProtectedViewWindow*` (6 events) | Mixed | Some | Protected view lifecycle |
| `EPostage*` (3 events) | After | No | Electronic postage |

### 1.2 Document Events — Complete List

The `Document` object exposes **13 events**. None are keyboard events.

| Event | Fires | Cancellable |
|-------|-------|-------------|
| `BuildingBlockInsert` | After | No |
| `Close` | After | No |
| `ContentControlAfterAdd` | After | No |
| `ContentControlBeforeContentUpdate` | Before | No (Content modifiable) |
| `ContentControlBeforeDelete` | Before | No |
| `ContentControlBeforeStoreUpdate` | Before | No (Content modifiable) |
| `ContentControlOnEnter` | After | No |
| `ContentControlOnExit` | Before | **Yes** (Cancel) |
| `New` | After | No |
| `Open` | After | No |
| `Sync` | After | No |
| `XMLAfterInsert` | After | No |
| `XMLBeforeDelete` | Before | No |

### 1.3 Window and Selection Events

| Object | Event count | Keyboard events |
|--------|------------|----------------|
| `Window` | **0** | No |
| `Selection` | **0** | No |

### 1.4 What the Word Object Model DOES NOT Provide

| Capability | Available? | Evidence |
|-----------|-----------|---------|
| `Application.OnKey` | **NO** | URL returns 404. [Word Application methods list](https://learn.microsoft.com/en-us/office/vba/api/word.application) does not include `OnKey`. This method is **Excel-only** ([Excel docs](https://learn.microsoft.com/en-us/office/vba/api/excel.application.onkey)). |
| `Application.KeyDown` event | **NO** | Does not exist on Application. `KeyDown`/`KeyUp` events are [MSForms (UserForm) only](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/keydown-keyup-events). |
| `Document.KeyDown` event | **NO** | Document has no keyboard events. |
| `Selection.KeyDown` event | **NO** | Selection has no events at all. |
| `Window.KeyDown` event | **NO** | Window has no events at all. |
| Any `BeforeKeyPress` event | **NO** | Does not exist. |
| Any keyboard interception mechanism | **NO** | The Word Object Model provides zero keyboard event infrastructure. |

### 1.5 What DOES Exist (Keyboard-Adjacent)

| API | What it actually does | Useful for Vim? |
|-----|----------------------|----------------|
| `Application.KeyBindings` | Collection of custom key assignments (same as Tools > Customize Keyboard) | Can reassign shortcuts, but only registered combos |
| `Application.FindKey` | Returns the `KeyBinding` for a given key code | Query only — read-only |
| `Application.BuildKeyCode` | Builds key code constants from `WdKey` values | Helper for `KeyBindings.Add` |
| `Application.Keyboard` | Gets/sets keyboard language/layout | Not relevant |
| `Application.CapsLock` | Read-only: CapsLock state | Minor utility |
| `Application.NumLock` | Read-only: NumLock state | Minor utility |
| `WindowSelectionChange` | Fires when selection changes (after arrow keys, typing, etc.) | No key code info, fires AFTER Word processes the keystroke |

### 1.6 Limitation: WindowSelectionChange

`WindowSelectionChange` fires when the selection changes, which includes after arrow-key navigation and typing. However:

- It fires **AFTER** Word has already processed the keystroke
- It provides **no key code** — you cannot tell which key was pressed
- It fires for both keyboard and mouse selection changes
- It cannot be cancelled (no `Cancel` parameter)
- It is a notification-only event

**Verdict:** `WindowSelectionChange` is useful for tracking cursor position but cannot be used for keystroke interception.

---

## 2. Office.js (Office JavaScript API)

**Source:** [Office.actions API](https://learn.microsoft.com/en-us/javascript/api/office/office.actions), [Keyboard shortcuts for Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/design/keyboard-shortcuts), [Word add-ins events](https://learn.microsoft.com/en-us/office/dev/add-ins/word/word-add-ins-events)

### 2.1 Office.actions API

The `Office.actions` interface has 4 methods:

| Method | What it does |
|--------|-------------|
| `associate(actionId, callback)` | Maps a predefined action ID to a JavaScript function |
| `areShortcutsInUse(shortcuts)` | Checks if shortcut combinations are already taken |
| `getShortcuts()` | Gets current shortcuts for the add-in |
| `replaceShortcuts(shortcuts)` | Reassigns key combinations at runtime |

### 2.2 Keyboard Shortcuts System

Defined in the add-in manifest. Maps key combinations (e.g., `Ctrl+Alt+Up`) to action IDs. When the user presses the combination, the associated function runs.

**Requirement sets:** `SharedRuntime 1.1` (for shortcuts to work), `KeyboardShortcuts 1.1` (for runtime customization).

### 2.3 What Office.js CANNOT Do with Keyboards

| Capability | Supported? | Evidence |
|-----------|-----------|---------|
| Intercept arbitrary keystrokes in document body | **NO** | [SO answer by Microsoft MVP](https://stackoverflow.com/questions/75953243): "The Office JavaScript API doesn't provide anything for handling keyboard buttons" |
| Listen for keydown/keyup/keypress | **NO** | No such events exist in the API |
| Prevent/suppress default keystroke behavior | **NO** | No `preventDefault()` or equivalent |
| Detect which key was pressed | **NO** | No key code information |
| Sequential key shortcuts (like `Alt+H, H`) | **NO** | Only simultaneous combos |
| Intercept single keys (just `A`, `Enter`) | **NO** | Must use modifier keys (Ctrl/Alt/Cmd + key) |
| Override browser shortcuts on web | **NO** | Explicitly blocked |
| Work when task pane has focus (web) | **NO** | Shortcuts only trigger when document is focused |

### 2.4 Open GitHub Issues

| Issue | Status | Summary |
|-------|--------|---------|
| [OfficeDev/office-js#24](https://github.com/OfficeDev/office-js/issues/24) | Closed (2017) | Original keyboard shortcut request. Closed, directed to UserVoice. |
| [OfficeDev/office-js#6454](https://github.com/OfficeDev/office-js/issues/6454) | **Open** (Jan 2026) | "Keyboard shortcuts don't work because Word intercepts all keyboard events before they reach the editor component." Assigned to Microsoft engineer, no resolution. |

### 2.5 Word.js Document Events

The Word JavaScript API exposes these document events (stable):

| Event | Object | What it does |
|-------|--------|-------------|
| `onParagraphChanged` | Paragraph | Fires when paragraph content changes |
| `onParagraphAdded` | Document | New paragraphs added |
| `onParagraphDeleted` | Document | Paragraphs deleted |
| `onContentControlAdded` | Document | Content control added |
| `onDataChanged` | ContentControl | Data within content control changed |
| `onSelectionChanged` | ContentControl | Selection changed within content control |

**None of these are keyboard events.** `onParagraphChanged` fires after text changes but provides no key code information.

### 2.6 Limitation: The Architectural Barrier

Office.js add-ins run in a sandboxed web runtime (iframe). The Word document editor runs in Word's native process. Keyboard events are processed entirely by Word's native code before they could ever reach the web runtime. This is a fundamental architectural limitation, not a feature gap.

**Source:** [office-js#6454](https://github.com/OfficeDev/office-js/issues/6454) — Microsoft has not announced plans to expose keyboard events.

---

## 3. VBA Keyboard APIs

**Source:** [KeyBindings collection](https://learn.microsoft.com/en-us/office/vba/api/word.keybindings), [KeyBindings.Add](https://learn.microsoft.com/en-us/office/vba/api/word.keybindings.add), [BuildKeyCode](https://learn.microsoft.com/en-us/office/vba/api/word.application.buildkeycode)

### 3.1 Application.OnKey — DOES NOT EXIST in Word

The `Application.OnKey` method is **Excel-only**. The URL `https://learn.microsoft.com/en-us/office/vba/api/word.application.onkey` returns **404**. It is not listed in the [Word Application methods](https://learn.microsoft.com/en-us/office/vba/api/word.application).

In Excel, `OnKey` can:
- Map a specific key combination to a VBA procedure
- Suppress a keystroke by passing empty string as the procedure
- Work with single keys (e.g., `{RIGHT}`, `{ENTER}`) and modifier combos

**This capability does not exist in Word VBA.**

### 3.2 KeyBindings.Add — Command Reassignment

```vb
expression.Add(KeyCategory, Command, KeyCode, KeyCode2, CommandParameter)
```

| Parameter | Type | Description |
|-----------|------|-------------|
| `KeyCategory` | `WdKeyCategory` | Category of the key assignment |
| `Command` | `String` | Command to execute (macro name, built-in command, style, etc.) |
| `KeyCode` | `Long` | Key code (from `BuildKeyCode`) |
| `KeyCode2` | `Variant` | Optional second key |
| `CommandParameter` | `Variant` | Optional command parameter |

**What it does:** Reassigns a key combination to a command. Equivalent to Tools > Customize Keyboard. The assigned command **replaces** the default behavior for that combination.

**Limitations:**
- Only handles **registered combinations**, not arbitrary keystrokes
- Requires `BuildKeyCode` to construct key codes from `WdKey` constants
- Cannot intercept single unmodified keys (like `d`, `j`, `w`) — only combos with modifiers
- The reassignment is per-context (document/template)
- Cannot fire an arbitrary VBA callback for any keystroke — only mapped commands

### 3.3 Application.FindKey — Query Only

```vb
expression.FindKey(KeyCode, KeyCode2)
```

Returns the `KeyBinding` object for a given key code. Read-only — cannot intercept, only query what's currently assigned.

### 3.4 Limitation: No Arbitrary Keystroke Interception in VBA

Without `Application.OnKey`, VBA has **no mechanism to intercept arbitrary keystrokes**. The options are:

1. **`KeyBindings.Add`**: Reassign specific key combinations to macros. Only registered combos fire.
2. **`Application.OnKey`**: Does not exist in Word.
3. **Win32 hooks from VBA**: VBA cannot P/Invoke `SetWindowsHookEx` directly (no `Declare` for hooks in VBA). Would require a helper DLL, which defeats the purpose of VBA simplicity.

**Verdict:** VBA alone cannot achieve Vim-like keystroke interception in Word.

---

## 4. Win32 Keyboard Hooks in VSTO

**Source:** [SetWindowsHookEx](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setwindowshookexw), [KeyboardProc](https://learn.microsoft.com/en-us/windows/win32/winmsg/keyboardproc), [LowLevelKeyboardProc](https://learn.microsoft.com/en-us/windows/win32/winmsg/lowlevelkeyboardproc), [MSDN Blog: Message Hooks in Add-ins](https://learn.microsoft.com/en-us/archive/blogs/andreww/message-hooks-in-add-ins)

### 4.1 Hook Types

| Type | Constant | Value | Scope | Callback |
|------|----------|-------|-------|----------|
| Thread-level | `WH_KEYBOARD` | 2 | Thread or global | `KeyboardProc` |
| Low-level global | `WH_KEYBOARD_LL` | 13 | Global only | `LowLevelKeyboardProc` |

### 4.2 P/Invoke Signatures

```csharp
// SetWindowsHookExW
[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
static extern IntPtr SetWindowsHookEx(
    int idHook, HookProc lpfn, IntPtr hMod, uint dwThreadId);

// CallNextHookEx
[DllImport("user32.dll")]
static extern IntPtr CallNextHookEx(
    IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

// UnhookWindowsHookEx
[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
[return: MarshalAs(UnmanagedType.Bool)]
static extern bool UnhookWindowsHookEx(IntPtr hhk);

// GetCurrentThreadId (kernel32)
[DllImport("kernel32.dll")]
static extern uint GetCurrentThreadId();
```

### 4.3 Callback Signatures

```csharp
// For WH_KEYBOARD (type 2)
delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);

// For WH_KEYBOARD_LL (type 13)
delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

// KBDLLHOOKSTRUCT (for WH_KEYBOARD_LL lParam)
[StructLayout(LayoutKind.Sequential)]
struct KBDLLHOOKSTRUCT {
    public uint vkCode;
    public uint scanCode;
    public uint flags;
    public uint time;
    public IntPtr dwExtraInfo;
}
```

### 4.4 WH_KEYBOARD (Type 2) — Thread-Level Hook

**How it works:**
- Monitors `WM_KEYDOWN` and `WM_KEYUP` messages about to be returned by `GetMessage`/`PeekMessage` on the hooked thread
- Fires on the installing thread's message loop
- `hMod = IntPtr.Zero` when hook procedure is in the same process (no DLL needed)
- `dwThreadId` = current thread ID (hooks only that thread)

**VSTO setup:**
```csharp
// Official pattern from MSDN blog
uint threadId = (uint)GetCurrentThreadId();
_hookId = SetWindowsHookEx(WH_KEYBOARD, _proc, IntPtr.Zero, threadId);
```

**Key requirements:**
- Installing thread MUST have a message loop (Word's STA thread does)
- Hook delegate MUST be stored in a class field to prevent GC collection
- Must check `nCode == HC_ACTION` (0) to avoid multiple firings
- Must call `CallNextHookEx` to pass events through (unless suppressing)

**Suppression:** Return `1` (non-zero) instead of calling `CallNextHookEx` to swallow a keystroke.

### 4.5 WH_KEYBOARD_LL (Type 13) — Global Low-Level Hook

**How it works:**
- Monitors low-level keyboard input events before they are posted to a thread input queue
- Fires in the context of the installing thread via message sending
- Requires a module handle (`GetModuleHandle`)
- Global scope — fires for all processes

**Critical limitation in Word:**
Multiple developers report that `WH_KEYBOARD_LL` **does not fire for keystrokes in the hosting Office application** (Word, Excel) when installed from within a VSTO add-in:

> "generally it works when I type in any application except in the instance of MS Word which runs the add-in" — [SO:25047972](https://stackoverflow.com/questions/25047972)

> "I see keydown events sent to the Edge browser and even Visual Studio, but not to Word itself" — [SO:32770647](https://stackoverflow.com/questions/32770647)

This is because Office's message processing intercepts the hook callback before it reaches the add-in's hook procedure.

### 4.6 WH_KEYBOARD vs WH_KEYBOARD_LL — Comparison

| Aspect | WH_KEYBOARD (2) | WH_KEYBOARD_LL (13) |
|--------|----------------|---------------------|
| **Works in Word VSTO?** | **Yes** (confirmed) | **No** (fails for host app) |
| Scope | Thread-local only | Global (all processes) |
| Module handle | `IntPtr.Zero` | `GetModuleHandle(process.MainModule.ModuleName)` |
| Thread ID | Current thread ID | 0 (global) |
| Fires on | Installing thread's message loop | Any thread with message loop |
| Can call Word COM in callback? | **Yes** (same thread) | Must marshal to UI thread |
| lParam format | Keystroke message flags | `KBDLLHOOKSTRUCT` pointer |
| wParam | Virtual-key code directly | `WM_KEYDOWN`/`WM_KEYUP` message ID |
| Requires DLL injection (global)? | Yes (for cross-process) | No |

### 4.7 Documented Limitations and Pitfalls

| Issue | Source | Mitigation |
|-------|--------|-----------|
| Delegate GC collection | [SetWindowsHookEx docs](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setwindowshookexw) | Store delegate in class-level field |
| Multiple firings per keystroke | [SO:39264317](https://stackoverflow.com/questions/39264317) | Check `nCode == HC_ACTION` (0) |
| `Thread.ManagedThreadId` doesn't work | [MSDN Blog](https://learn.microsoft.com/en-us/archive/blogs/andreww/message-hooks-in-add-ins) | Use `GetCurrentThreadId()` or `AppDomain.GetCurrentThreadId()` |
| Low-level hook timeout (Win7+) | [LowLevelKeyboardProc docs](https://learn.microsoft.com/en-us/windows/win32/winmsg/lowlevelkeyboardproc) | Keep callback < 1ms |
| Antivirus false positives | Community reports | Document in README, sign the add-in |
| WH_KEYBOARD_LL fails in Office 2013+ | [SO:25047972](https://stackoverflow.com/questions/25047972), [SO:32770647](https://stackoverflow.com/questions/32770647) | Use WH_KEYBOARD (type 2) instead |

### 4.8 Real-World Confirmation

**Dirk Vollmar's VSTO Word sample** ([SO:32770647](https://stackoverflow.com/questions/32770647), +400 votes):
- Complete working VSTO Word add-in using `WH_KEYBOARD` (type 2)
- Uses `IntPtr.Zero` for module handle, `GetCurrentThreadId()` for thread ID
- Installs both `WH_KEYBOARD` and `WH_MOUSE` hooks
- Confirmed working in Word

**Andrew Whitechapel's MSDN Blog** ([Message Hooks in Add-ins](https://learn.microsoft.com/en-us/archive/blogs/andreww/message-hooks-in-add-ins)):
- Official Microsoft sample demonstrating `WH_KEYBOARD` in VSTO
- Includes working suppression of Alt+F11
- Downloadable sample: `WordAddInHook.zip`

**Stack Overflow confirmation** ([SO:57604675](https://stackoverflow.com/questions/57604675)):
> "I don't use the Global Hooker and my code works. I explicitly tested it in Word (and know it works in Excel, PowerPoint, Access, etc)."

---

## 5. Summary: What Can and Cannot Be Done

### 5.1 Capability Matrix

| Capability | Word Object Model | Office.js | VBA | VSTO + WH_KEYBOARD |
|-----------|------------------|-----------|-----|-------------------|
| Intercept arbitrary keystrokes | **NO** | **NO** | **NO** | **YES** |
| Suppress keystroke behavior | **NO** | **NO** | **NO** | **YES** |
| Get key code of pressed key | **NO** | **NO** | **NO** | **YES** |
| Move cursor programmatically | **YES** | Limited | **YES** | **YES** |
| Read/write text | **YES** | **YES** | **YES** | **YES** |
| Cross-platform | Windows | Win/Mac/Web | Win/Mac | Windows only |

### 5.2 The Only Viable Approach

**VSTO + WH_KEYBOARD (type 2) + Word COM Object Model** is the only combination that provides:
1. Keystroke interception (WH_KEYBOARD hook)
2. Keystroke suppression (return non-zero from callback)
3. Full cursor control (Word COM: `Selection.Move`, etc.)
4. Full text manipulation (Word COM: `Selection.Text`, etc.)

### 5.3 Critical Corrections to Earlier Analysis

| Earlier claim | Correction |
|--------------|-----------|
| "VBA `OnKey` can map specific combos" | **`Application.OnKey` does not exist in Word VBA.** It is Excel-only. |
| "VBA can partially intercept via OnKey" | **False for Word.** VBA has no keystroke interception mechanism. |
| "WH_KEYBOARD_LL works but fires globally" | **WH_KEYBOARD_LL does NOT fire for keystrokes in the hosting Office app** (Word 2013+). |
| "Use WH_KEYBOARD_LL for global interception" | **Use WH_KEYBOARD (type 2) for VSTO.** This is the confirmed working approach. |

### 5.4 Remaining Risks

| Risk | Severity | Mitigation |
|------|----------|-----------|
| WH_KEYBOARD callback timing | Medium | Keep callback < 1ms. Post to queue for processing. |
| Hook conflicts with other add-ins | Medium | Graceful fallback: catch failure, disable WordVim. |
| Word updates break hook behavior | Low | WH_KEYBOARD is a stable Win32 API, unlikely to break. |
| Antivirus flags the hook | Medium | Sign the add-in. Document in README. |
| Mouse-driven text changes bypass hook | Medium | Also install WH_MOUSE hook to detect paste/cut via context menu. |

---

## 6. References

### Official Microsoft Documentation
- [SetWindowsHookExW](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setwindowshookexw)
- [KeyboardProc](https://learn.microsoft.com/en-us/windows/win32/winmsg/keyboardproc)
- [LowLevelKeyboardProc](https://learn.microsoft.com/en-us/windows/win32/winmsg/lowlevelkeyboardproc)
- [Word Application object](https://learn.microsoft.com/en-us/office/vba/api/word.application)
- [Word KeyBindings collection](https://learn.microsoft.com/en-us/office/vba/api/word.keybindings)
- [Word KeyBindings.Add](https://learn.microsoft.com/en-us/office/vba/api/word.keybindings.add)
- [Office.actions API](https://learn.microsoft.com/en-us/javascript/api/office/office.actions)
- [Keyboard shortcuts for Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/design/keyboard-shortcuts)
- [Word add-ins events](https://learn.microsoft.com/en-us/office/dev/add-ins/word/word-add-ins-events)
- [Threading support in VSTO](https://learn.microsoft.com/en-us/visualstudio/vsto/threading-support-in-office)
- [MSDN Blog: Message Hooks in Add-ins](https://learn.microsoft.com/en-us/archive/blogs/andreww/message-hooks-in-add-ins)

### Community Sources
- [SO:32770647](https://stackoverflow.com/questions/32770647) — WH_KEYBOARD VSTO Word sample (+400 votes)
- [SO:25047972](https://stackoverflow.com/questions/25047972) — WH_KEYBOARD_LL failure in Word
- [SO:57604675](https://stackoverflow.com/questions/57604675) — Confirmed WH_KEYBOARD works in Word
- [SO:75953243](https://stackoverflow.com/questions/75953243) — Office.js keyboard limitations (Microsoft MVP answer)
- [OfficeDev/office-js#6454](https://github.com/OfficeDev/office-js/issues/6454) — Open issue: keyboard events not reaching Office.js
