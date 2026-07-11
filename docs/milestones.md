# WordVim Milestone Plan

> Each milestone = 1 day max. Each milestone = working prototype. No steps skipped.

## Prerequisites

- Visual Studio 2022 (Community is fine) with **Office/SharePoint development** workload
- .NET Framework 4.8 Developer Pack
- Microsoft Word 365 (desktop, not web)
- GitHub account (for CI later)

---

## Overview

| Day | Milestone | Working Prototype |
|-----|-----------|-------------------|
| 1 | Solution structure | `msbuild` succeeds, 4 projects compile |
| 2 | VimState | Unit tests: mode transitions work |
| 3 | KeyHandler | Unit tests: keystroke → action mapping works |
| 4 | Motions (basic) | Unit tests: h/j/k/l/w/b/e/0/$ produce correct Move actions |
| 5 | Operators (basic) | Unit tests: d/y/c/x produce correct Edit actions |
| 6 | CommandParser | Unit tests: `d2w`, `3dd`, `ci"` parse correctly |
| 7 | Motions (advanced) | Unit tests: f/t/F/T, gg/G, {/}, %, ^ produce correct actions |
| 8 | Operators (advanced) | Unit tests: >/<, ~, dot-repeat logic produce correct actions |
| 9 | State machine integration | Unit tests: full input → complete action pipeline works |
| 10 | VSTO add-in skeleton | Word loads add-in, status bar shows "WordVim loaded" |
| 11 | Ribbon mode indicator | Ribbon tab shows current mode (NORMAL) |
| 12 | Keyboard hook | Keystrokes logged to Debug output when typing in Word |
| 13 | Hook + Core wired | Pressing `i` changes Ribbon to INSERT, `Esc` back to NORMAL |
| 14 | Cursor movement (h/j/k/l) | h/j/k/l moves cursor in Word document |
| 15 | Insert mode entry/exit | `i` → type freely → `Esc` → back to Normal mode |
| 16 | Delete operator (d) | `dw` deletes a word, `dd` deletes a line |
| 17 | Yank (y) + Put (p) | `yw` then `p` duplicates a word |
| 18 | Change operator (c) | `cw` deletes word and enters Insert mode |
| 19 | Remaining operators | `x`, `>>`, `<<`, `~` work in Word |
| 20 | Word motions (w/b/e) | `w`/`b`/`e` navigate between words in Word |
| 21 | Line/char motions | `0`/`$`/`^`/`f`/`t` work in Word |
| 22 | Search (/ ? n N) | `/word` finds and highlights, `n`/`N` navigates |
| 23 | Command-line mode | `:w` saves, `:q` closes, `:wq` saves and closes |
| 24 | Visual mode | `v` enters visual, selection follows movement, `d` deletes selection |
| 25 | Registers + counts | `"a`/`"p` named registers, `3dd`/`5j` counts work |
| 26 | Undo/redo | `u` undoes, `Ctrl+r` redoes via Word COM |
| 27 | Dot repeat + paragraph motions | `.` repeats last change, `{`/`}`/`%` work |
| 28 | Edge cases + polish | Tab/Enter/Backspace in Insert, empty doc, multi-doc, error handling |
| 29 | Installer + CI | Inno Setup installs add-in, GitHub Actions builds on push |
| 30 | Documentation + v1.0 | README complete, GitHub release with installer |

---

## Phase A: Core Engine (Days 1-9)

All code lives in `WordVim.Core`. Zero Word dependency. Verified by unit tests only.

### Day 1 — Solution Structure

**Goal:** Build a solution with 4 projects that compiles.

**Deliverable:** `msbuild WordVim.sln` succeeds with 0 errors.

**Tasks:**
1. Create `WordVim.sln` at repository root
2. Create `src/WordVim.Core/WordVim.Core.csproj` — .NET Framework 4.8 Class Library
3. Create `src/WordVim.Hooks/WordVim.Hooks.csproj` — .NET Framework 4.8 Class Library
4. Create `src/WordVim.Word/WordVim.Word.csproj` — .NET Framework 4.8 Class Library
5. Create `src/WordVim.Addin/WordVim.Addin.csproj` — VSTO Word Add-in
6. Set project references: Hooks → Core, Word → Core, Addin → Core + Hooks + Word
7. Add `tests/WordVim.Core.Tests/WordVim.Core.Tests.csproj` — NUnit test project, references Core
8. Add placeholder `.cs` files in each project (empty namespaces)
9. Verify build: `msbuild WordVim.sln`

**Verification:** Clean build, 0 errors, 0 warnings. Test project runs (0 tests, green).

**Dependencies:** None.

---

### Day 2 — VimState (Mode Tracking)

**Goal:** State machine that tracks current Vim mode and handles transitions.

**Deliverable:** Unit tests pass: mode transitions are correct.

**Tasks:**
1. Create `src/WordVim.Core/VimMode.cs` — enum: `Normal`, `Insert`, `Visual`, `VisualLine`, `CommandLine`, `Search`
2. Create `src/WordVim.Core/VimState.cs`:
   - Property: `VimMode CurrentMode`
   - Method: `VimMode Transition(VimMode target)` — returns previous mode, sets new mode
   - Method: `bool IsInsertMode()` — true if Insert (keystrokes pass through to Word)
   - Method: `bool IsNormalMode()` — true if Normal (keystrokes are Vim commands)
3. Create `tests/WordVim.Core.Tests/VimStateTests.cs`:
   - Test: starts in Normal mode
   - Test: Normal → Insert (returns Normal)
   - Test: Insert → Normal via Esc
   - Test: Normal → Visual → Normal
   - Test: Normal → CommandLine → Normal
   - Test: Normal → Search → Normal
   - Test: Visual → VisualLine toggle
   - Test: cannot transition from Insert to Visual directly (must go through Normal)

**Verification:** `dotnet test` — all tests pass.

**Dependencies:** Day 1.

---

### Day 3 — KeyHandler (Keystroke → Action Mapping)

**Goal:** Map keystrokes to abstract actions, independent of what the actions do.

**Deliverable:** Unit tests pass: every Normal mode key produces the correct action.

**Tasks:**
1. Create `src/WordVim.Core/VimAction.cs`:
   - Class hierarchy (or record types):
     - `MoveAction` — direction: Left/Right/Up/Down, count: int
     - `MoveToAction` — target: WordStart/WordEnd/LineStart/LineEnd/DocStart/DocEnd/CharFind/CharTill
     - `EditAction` — type: Delete/Yank/Change/Put/Indent/Dedent/ToggleCase, scope: Char/Selection
     - `ChangeModeAction` — target mode: VimMode
     - `SearchAction` — direction: Forward/Reverse
     - `CommandLineAction` — command string
     - `RepeatAction` — repeat last change
     - `UndoAction`, `RedoAction`
     - `NoOp` — ignored keystroke
2. Create `src/WordVim.Core/KeyHandler.cs`:
   - Method: `VimAction HandleKey(VimState state, int keyCode, bool shift, bool ctrl, bool alt)`
   - Normal mode: `h`→MoveLeft, `j`→MoveDown, `k`→MoveUp, `l`→MoveRight
   - Normal mode: `i`→ChangeMode(Insert), `a`→ChangeMode(Insert), `o`→ChangeMode(Insert)
   - Normal mode: `v`→ChangeMode(Visual), `V`→ChangeMode(VisualLine)
   - Normal mode: `/`→Search(Forward), `?`→Search(Reverse), `:`→CommandLine
   - Insert mode: returns `NoOp` for all keys (pass through to Word)
   - Esc key in any mode except Normal → ChangeMode(Normal)
3. Create `tests/WordVim.Core.Tests/KeyHandlerTests.cs`:
   - Test each Normal mode key produces the correct action
   - Test Insert mode: all keys → NoOp
   - Test Esc returns to Normal from any mode
   - Test Ctrl/Ctrl+key combos (Ctrl+r → Redo)

**Verification:** `dotnet test` — all tests pass.

**Dependencies:** Day 2.

---

### Day 4 — Motions (Basic)

**Goal:** Normal mode motions: character, word, and line navigation.

**Deliverable:** Unit tests pass: h/j/k/l, w/b/e, 0/$ produce correct Move/MoveTo actions.

**Tasks:**
1. Extend `KeyHandler.HandleKey` in Normal mode:
   - `h` → `MoveAction(Left, 1)`
   - `l` → `MoveAction(Right, 1)`
   - `j` → `MoveAction(Down, 1)`
   - `k` → `MoveAction(Up, 1)`
   - `w` → `MoveToAction(WordStart, Forward)`
   - `b` → `MoveToAction(WordStart, Backward)`
   - `e` → `MoveToAction(WordEnd, Forward)`
   - `0` → `MoveToAction(LineStart)`
   - `$` → `MoveToAction(LineEnd)`
   - `^` → `MoveToAction(FirstNonBlank)`
2. Create `tests/WordVim.Core.Tests/MotionTests.cs`:
   - Test each motion produces the correct action type and parameters
   - Test: `h` in Normal → MoveAction(Left, 1)
   - Test: `w` in Normal → MoveToAction(WordStart, Forward)
   - Test: `0` in Normal → MoveToAction(LineStart)
   - Test: motions in Insert mode → NoOp (pass through)

**Verification:** `dotnet test` — all tests pass.

**Dependencies:** Day 3.

---

### Day 5 — Operators (Basic)

**Goal:** Normal mode operators: delete, yank, change, delete-char.

**Deliverable:** Unit tests pass: d/y/c/x produce correct Edit actions with pending state.

**Tasks:**
1. Create `src/WordVim.Core/PendingOperator.cs`:
   - Tracks: operator type (Delete/Yank/Change) + waiting for motion
   - When operator is pressed, KeyHandler stores pending state
   - Next keystroke is interpreted as motion (not as standalone key)
2. Extend `KeyHandler.HandleKey`:
   - `d` → store pending Delete, return `NoOp` (wait for motion)
   - `y` → store pending Yank, return `NoOp`
   - `c` → store pending Change, return `NoOp`
   - `x` → `EditAction(Delete, Char)` (immediate, no pending)
   - When pending + motion received: return `EditAction(operator, Selection)` with the motion range
   - `d` + `d` → `EditAction(Delete, Line)` (operator doubled = operate on line)
   - `y` + `y` → `EditAction(Yank, Line)`
   - `c` + `c` → `EditAction(Change, Line)`
3. Create `tests/WordVim.Core.Tests/OperatorTests.cs`:
   - Test: `d` alone → pending state (NoOp returned)
   - Test: `d` then `w` → EditAction(Delete, Selection, WordForward)
   - Test: `dd` → EditAction(Delete, Line)
   - Test: `yw` → EditAction(Yank, Selection, WordForward)
   - Test: `yy` → EditAction(Yank, Line)
   - Test: `cw` → EditAction(Change, Selection, WordForward)
   - Test: `cc` → EditAction(Change, Line)
   - Test: `x` → EditAction(Delete, Char) immediately
   - Test: pending state cleared after action completed
   - Test: Esc while pending → cancels pending, no action

**Verification:** `dotnet test` — all tests pass.

**Dependencies:** Day 4.

---

### Day 6 — CommandParser (Operator + Motion + Count)

**Goal:** Parse compound commands: count prefix, operator+motion composition.

**Deliverable:** Unit tests pass: `3dd`, `d2w`, `ci"`, `y$` parse correctly.

**Tasks:**
1. Create `src/WordVim.Core/CommandParser.cs`:
   - Maintains: `int pendingCount` (numeric prefix), `PendingOperator pendingOp`
   - Method: `VimAction ParseKey(VimState state, int keyCode, bool shift, bool ctrl)`
   - Digit keys `1`-`9` (or `0` if count already started): accumulate count
   - Motion keys with count: repeat motion N times
   - Operator + count + motion: apply operator to N repetitions of motion
   - Operator + operator (dd/yy/cc): apply to N lines
2. Extend for text-object style commands:
   - `c` + `i` + `w` → Change word (inner word)
   - `d` + `i` + `w` → Delete word
   - `y` + `i` + `w` → Yank word
   - `c` + `i` + `"` → Change inside quotes
   - `d` + `i` + `(` → Delete inside parens
   - (Store these as compound sequences, not immediate actions)
3. Create `tests/WordVim.Core.Tests/CommandParserTests.cs`:
   - Test: `3j` → MoveAction(Down, 3)
   - Test: `5l` → MoveAction(Right, 5)
   - Test: `d2w` → EditAction(Delete, Selection, 2×WordForward)
   - Test: `3dd` → EditAction(Delete, 3×Line)
   - Test: `ciw` → EditAction(Change, InnerWord)
   - Test: `di"` → EditAction(Delete, InnerQuotes)
   - Test: `y$` → EditAction(Yank, Selection, LineEnd)
   - Test: `2yy` → EditAction(Yank, 2×Line)
   - Test: count resets after action completes
   - Test: invalid count sequence resets

**Verification:** `dotnet test` — all tests pass.

**Dependencies:** Day 5.

---

### Day 7 — Motions (Advanced)

**Goal:** Character search, jumps, paragraph, bracket matching.

**Deliverable:** Unit tests pass: f/t/F/T, gg/G, {/}, % produce correct actions.

**Tasks:**
1. Extend `KeyHandler` / `CommandParser` for:
   - `f{char}` → `MoveToAction(CharFind, Forward, char)`
   - `t{char}` → `MoveToAction(CharTill, Forward, char)` (stop before char)
   - `F{char}` → `MoveToAction(CharFind, Backward, char)`
   - `T{char}` → `MoveToAction(CharTill, Backward, char)`
   - `gg` → `MoveToAction(DocStart)` (or line N if prefixed: `10gg`)
   - `G` → `MoveToAction(DocEnd)` (or line N if prefixed: `10G`)
   - `{` → `MoveToAction(ParagraphBack)`
   - `}` → `MoveToAction(ParagraphForward)`
   - `%` → `MoveToAction(MatchingBracket)`
2. Create `tests/WordVim.Core.Tests/AdvancedMotionTests.cs`:
   - Test: `fx` → MoveToAction(CharFind, Forward, 'x')
   - Test: `t;` → MoveToAction(CharTill, Forward, ';')
   - Test: `Fa` → MoveToAction(CharFind, Backward, 'a')
   - Test: `gg` → MoveToAction(DocStart)
   - Test: `G` → MoveToAction(DocEnd)
   - Test: `5G` → MoveToAction(GoToLine, 5)
   - Test: `}` → MoveToAction(ParagraphForward)
   - Test: `%` → MoveToAction(MatchingBracket)
   - Test: `df;` → EditAction(Delete, Selection, CharFind(';', Forward))

**Verification:** `dotnet test` — all tests pass.

**Dependencies:** Day 6.

---

### Day 8 — Operators (Advanced) + Dot Repeat

**Goal:** Indent, toggle case, dot repeat logic.

**Deliverable:** Unit tests pass: >/<, ~, . produce correct actions.

**Tasks:**
1. Extend `KeyHandler`:
   - `>` + motion → `EditAction(Indent, Selection, motion)`
   - `<` + motion → `EditAction(Dedent, Selection, motion)`
   - `>>` → `EditAction(Indent, Line)`
   - `<<` → `EditAction(Dedent, Line)`
   - `~` → `EditAction(ToggleCase, Char)` (immediate, no pending)
2. Create `src/WordVim.Core/RepeatManager.cs`:
   - Stores: last `VimAction` (or sequence of actions for compound commands)
   - Method: `Record(VimAction action)` — store after each completed action
   - Method: `VimAction GetRepeatAction()` — return last stored action
   - Only records change actions (Delete, Change, Indent, Dedent, ToggleCase, Put)
   - Does not record moves or mode changes
3. Extend `KeyHandler`:
   - `.` in Normal mode → returns `RepeatAction` (which resolves via RepeatManager)
4. Create `tests/WordVim.Core.Tests/AdvancedOperatorTests.cs`:
   - Test: `>w` → EditAction(Indent, Selection, WordForward)
   - Test: `>>` → EditAction(Indent, Line)
   - Test: `~` → EditAction(ToggleCase, Char)
   - Test: repeat manager records Delete action
   - Test: repeat manager does NOT record Move action
   - Test: `.` returns repeat of last recorded action
   - Test: `.` after no changes → NoOp

**Verification:** `dotnet test` — all tests pass.

**Dependencies:** Day 7.

---

### Day 9 — State Machine Integration (Full Pipeline)

**Goal:** Complete input→state→action pipeline, including Visual mode, registers, search/CLI state.

**Deliverable:** Unit tests pass: end-to-end pipeline handles all command types.

**Tasks:**
1. Create `src/WordVim.Core/VimEngine.cs` — the single entry point:
   - Constructor takes: `VimState`, `KeyHandler`, `CommandParser`, `RepeatManager`
   - Method: `VimAction ProcessKey(int keyCode, bool shift, bool ctrl, bool alt)`
   - Orchestrates: state check → parser update → action generation → repeat recording
2. Create `src/WordVim.Core/RegisterManager.cs`:
   - 26 named registers (`"a`-`"z`), unnamed register (`"`), numbered registers (`"0`-`"9`)
   - Method: `Store(char register, string text)`
   - Method: `string Retrieve(char register)`
   - Method: `StoreUnnamed(string text)` — default register
   - Method: `string RetrieveUnnamed()`
3. Extend `KeyHandler` for Visual mode:
   - In Visual mode, motion keys extend selection (not move cursor)
   - `d`/`y`/`c` in Visual mode operate on current selection
   - `v` toggles Visual → Normal
4. Extend `KeyHandler` for Search mode:
   - `/` → enters Search mode, keystrokes accumulate into search string
   - Enter → commits search, returns SearchAction
   - Esc → cancels search, returns to Normal
5. Extend `KeyHandler` for CommandLine mode:
   - `:` → enters CommandLine mode, keystrokes accumulate into command string
   - Enter → commits command, returns CommandLineAction
   - Esc → cancels, returns to Normal
6. Create `tests/WordVim.Core.Tests/EngineIntegrationTests.cs`:
   - Test: full sequence `d2w` → single Delete action with 2×WordForward
   - Test: Visual mode `vllld` → Delete action with 3-char selection
   - Test: `/test` then Enter → SearchAction("test", Forward)
   - Test: `:wq` then Enter → CommandLineAction("wq")
   - Test: `"ayy` → yanks line to register 'a'
   - Test: `"ap` → puts from register 'a'
   - Test: `.` after `dw` → repeats Delete(WordForward)

**Verification:** `dotnet test` — all tests pass. Core engine is complete.

**Dependencies:** Day 8.

---

## Phase B: Word Integration (Days 10-28)

Connect Core engine to Word via VSTO + Win32 hooks + COM. Each day produces a feature demonstrable in Word.

### Day 10 — VSTO Add-in Skeleton

**Goal:** Word loads the add-in without crashing.

**Deliverable:** Opening Word shows "WordVim loaded" in the status bar.

**Tasks:**
1. In `WordVim.Addin`:
   - Implement `ThisAddIn.cs`:
     - `ThisAddIn_Startup`: set `Application.StatusBar = "WordVim loaded"`
     - `ThisAddIn_Shutdown`: clear status bar
2. Set VSTO project properties:
   - Host: Microsoft Word
   - .NET Framework 4.8
   - Register for COM interop: checked
3. Build and run (F5 opens Word with debugger attached)
4. Verify status bar shows "WordVim loaded"
5. Close Word, verify clean shutdown (no error dialog)

**Verification:** Word starts, shows status message, closes cleanly.

**Dependencies:** Day 1, Day 9.

---

### Day 11 — Ribbon Mode Indicator

**Goal:** Ribbon shows current Vim mode.

**Deliverable:** Word Ribbon has a "WordVim" tab with a mode label showing "NORMAL".

**Tasks:**
1. Add Ribbon XML to `WordVim.Addin`:
   - Custom tab: "WordVim"
   - Group: "Mode"
   - Label control: `lblMode` with text "NORMAL"
   - Button: "Settings" (placeholder, does nothing yet)
2. Create `src/WordVim.Addin/ModeIndicator.cs`:
   - Static method: `UpdateMode(string modeName)` — updates Ribbon label text
   - Static method: `UpdateStatusBar(string text)` — updates Word status bar
3. Wire `ThisAddIn_Startup` to show Ribbon and set initial mode to "NORMAL"
4. Build and run: verify Ribbon tab appears with mode label

**Verification:** Ribbon tab visible, mode label shows "NORMAL".

**Dependencies:** Day 10.

---

### Day 12 — Keyboard Hook

**Goal:** Intercept keystrokes in Word before they are processed.

**Deliverable:** Typing in Word logs each keystroke to Debug output.

**Tasks:**
1. In `WordVim.Hooks`, create `KeyboardHook.cs`:
   - P/Invoke declarations:
     - `SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hMod, uint dwThreadId)`
     - `UnhookWindowsHookEx(IntPtr hhk)`
     - `CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam)`
     - `GetModuleHandle(string lpModuleName)`
   - Delegate: `HookProc(int nCode, IntPtr wParam, IntPtr lParam)`
    - Constant: `WH_KEYBOARD = 2`
    - P/Invoke: `GetCurrentThreadId()` from kernel32 (NOT `Thread.ManagedThreadId` which doesn't work with SetWindowsHookEx)
    - Thread-level hook: `dwThreadId = GetCurrentThreadId()` (hooks only the Word UI thread)
2. Hook callback (static method to prevent GC collection):
    - Store delegate in class-level field (GC will collect otherwise — silent failure)
    - If `nCode == HC_ACTION` (0): extract key code from `wParam`, log `Debug.WriteLine($"Key: {key}")`
    - Always call `CallNextHookEx` (don't suppress yet — just observe)
    - Do NOT process when `nCode < 0` (must pass through per Win32 contract)
3. Create `HookManager.cs`:
   - `Install()` — installs the hook, stores handle
   - `Uninstall()` — removes the hook
   - Called from `ThisAddIn_Startup` / `ThisAddIn_Shutdown`
4. Build and run: type in Word, verify keystrokes appear in Visual Studio Output window

**Verification:** Every keystroke typed in Word appears in Debug output.

**Dependencies:** Day 11.

---

### Day 13 — Hook + Core Wired

**Goal:** Keystrokes flow through the Core engine. Mode transitions are reflected in the Ribbon.

**Deliverable:** Press `i` → Ribbon changes to "INSERT". Press `Esc` → Ribbon changes to "NORMAL".

**Tasks:**
1. Create `src/WordVim.Addin/VimOrchestrator.cs`:
   - Holds: `VimEngine`, `VimState`, `ModeIndicator`
   - Method: `OnKeyStroke(int keyCode, bool shift, bool ctrl, bool alt)`:
     - Call `engine.ProcessKey(...)`
     - If result is `ChangeModeAction` → update `state`, call `ModeIndicator.UpdateMode(...)`
     - Other actions: logged to Debug (not yet connected to Word)
2. Modify `KeyboardHook` callback:
   - Instead of just logging, call `VimOrchestrator.OnKeyStroke(...)`
3. Modify `ThisAddIn_Startup`:
   - Create `VimOrchestrator` instance
   - Pass it to `HookManager`
4. Build and run:
   - Press `i` → Ribbon shows "INSERT"
   - Press `Esc` → Ribbon shows "NORMAL"
   - Press `v` → Ribbon shows "VISUAL"
   - Press `V` → Ribbon shows "VISUAL LINE"
   - Press `:` → Ribbon shows "COMMAND LINE"
   - Press `Esc` → Ribbon shows "NORMAL"

**Verification:** All mode transitions visible in Ribbon. Keystrokes no longer typed into Word in Normal mode (they are consumed by the hook — but note: we are NOT suppressing yet, so characters will appear AND be processed; suppression comes next).

**Dependencies:** Day 9, Day 12.

---

### Day 14 — Keystroke Suppression + Cursor Movement (h/j/k/l)

**Goal:** In Normal mode, keystrokes are suppressed (not typed into Word) and h/j/k/l move the cursor.

**Deliverable:** In Normal mode, `h`/`j`/`k`/`l` move the cursor. Letters do not appear in the document.

**Tasks:**
1. Create `src/WordVim.Word/WordAdapter.cs`:
   - P/Invoke: `GetForegroundWindow()` to verify Word is focused
   - COM interop: get `Application` object from VSTO `ThisAddIn.Application`
   - Method: `MoveCursor(int direction, int count)`:
     - `Selection.MoveLeft(wdCharacter, count)` for Left
     - `Selection.MoveRight(wdCharacter, count)` for Right
     - `Selection.MoveDown(wdLine, count)` for Down
     - `Selection.MoveUp(wdLine, count)` for Up
2. Modify `VimOrchestrator.OnKeyStroke`:
   - If `state.IsNormalMode()` AND action is Move/MoveTo → call `WordAdapter.MoveCursor(...)`, suppress keystroke
   - If `state.IsInsertMode()` → let keystroke pass through (call `CallNextHookEx`)
3. Modify `KeyboardHook` callback:
   - Return `1` (non-zero) to suppress keystroke when in Normal mode
   - Return `CallNextHookEx(...)` result to pass through in Insert mode
4. Build and run:
   - Press `i` → type freely → `Esc` → Normal mode
   - In Normal mode: `h`/`j`/`k`/`l` move cursor, no characters inserted
   - Press `i` again → type freely

**Verification:** Modal editing works for basic cursor movement. No stray characters in Normal mode.

**Dependencies:** Day 13.

---

### Day 15 — Insert Mode Entry/Exit (i/a/o)

**Goal:** `i`, `a`, `o` enter Insert mode. `Esc` returns to Normal.

**Deliverable:** `i` positions cursor and enters Insert. `a` moves one right and enters Insert. `o` opens line below and enters Insert.

**Tasks:**
1. Extend `WordAdapter`:
   - `EnterInsertAtCursor()` — collapse selection to insertion point (already there)
   - `EnterInsertAfterCursor()` — `Selection.MoveRight(wdCharacter, 1)` then collapse
   - `OpenLineBelow()` — move to end of line, `Selection.TypeParagraph`, enter Insert
2. Extend `VimOrchestrator` to handle `ChangeModeAction` with specific sub-types:
   - `i` → `WordAdapter.EnterInsertAtCursor()` + switch to Insert
   - `a` → `WordAdapter.EnterInsertAfterCursor()` + switch to Insert
   - `o` → `WordAdapter.OpenLineBelow()` + switch to Insert
   - `A` → move to end of line + switch to Insert
   - `O` → open line above + switch to Insert
3. Build and run:
   - `i` → type text → `Esc` → Normal
   - `a` → cursor moves one right → type text → `Esc`
   - `o` → new line created → type text → `Esc`
   - `A` → cursor at end of line → type text → `Esc`
   - `O` → new line above → type text → `Esc`

**Verification:** All Insert entry methods work correctly. Text typed in Insert mode appears in document. Esc returns to Normal.

**Dependencies:** Day 14.

---

### Day 16 — Delete Operator (d)

**Goal:** `d` + motion deletes text. `dd` deletes line.

**Deliverable:** `dw` deletes a word. `dd` deletes current line. `d$` deletes to end of line.

**Tasks:**
1. Extend `WordAdapter`:
   - `DeleteRange(startOffset, endOffset)` — select range, call `Selection.Delete()`
   - `DeleteLine()` — select current line (`Selection.HomeKey(wdLine)` + `Selection.EndOf(wdLine, wdExtend)`), delete
   - Helper: `GetSelectionRange(motion)` — convert motion to absolute offsets
2. Extend `VimOrchestrator` to handle `EditAction(Delete, ...)`:
   - Compute target range from motion
   - Call `WordAdapter.DeleteRange(...)` or `DeleteLine()`
3. Build and run:
   - Place cursor on a word → `dw` → word deleted
   - Place cursor on a line → `dd` → line deleted
   - Place cursor mid-line → `d$` → rest of line deleted
   - `d0` → delete to start of line

**Verification:** Delete operator works for all basic motions.

**Dependencies:** Day 15.

---

### Day 17 — Yank (y) + Put (p)

**Goal:** `y` + motion yanks text. `p` pastes after cursor. `P` pastes before.

**Deliverable:** `yw` yanks a word, `p` pastes it. `yy` yanks a line, `p` pastes it below.

**Tasks:**
1. Extend `WordAdapter`:
   - `string GetText(range)` — read text from range
   - `InsertText(string text, bool afterCursor)` — paste text at cursor position
2. Extend `VimOrchestrator` to handle `EditAction(Yank, ...)`:
   - Compute target range from motion
   - Read text via `WordAdapter.GetText(range)`
   - Store in `RegisterManager` (unnamed register by default)
3. Extend `VimOrchestrator` to handle `EditAction(Put, ...)`:
   - Retrieve text from `RegisterManager`
   - Call `WordAdapter.InsertText(text, afterCursor: true)` for `p`
   - Call `WordAdapter.InsertText(text, afterCursor: false)` for `P`
4. Build and run:
   - `yw` → cursor on word → `p` → word duplicated after cursor
   - `yy` → `p` → line duplicated below
   - `yl` → `p` → single character duplicated

**Verification:** Yank and Put work for words, lines, and characters.

**Dependencies:** Day 16.

---

### Day 18 — Change Operator (c)

**Goal:** `c` + motion deletes text and enters Insert mode.

**Deliverable:** `cw` deletes word and enters Insert. `cc` replaces entire line. `c$` deletes to end of line and enters Insert.

**Tasks:**
1. Extend `VimOrchestrator` to handle `EditAction(Change, ...)`:
   - Compute target range from motion
   - Delete range (same as Delete operator)
   - Switch to Insert mode
2. `cc` — select entire line content (not the newline), delete, enter Insert
3. Build and run:
   - `cw` → word deleted → typing replaces it → `Esc`
   - `cc` → line content deleted → typing replaces it → `Esc`
   - `c$` → rest of line deleted → typing replaces it → `Esc`
   - `ciw` → inner word deleted → typing replaces it → `Esc`

**Verification:** Change operator works identically to Delete + Insert mode entry.

**Dependencies:** Day 17.

---

### Day 19 — Remaining Operators (x, >, <, ~)

**Goal:** Single-key operators: delete-char, indent, dedent, toggle-case.

**Deliverable:** `x` deletes character. `>>` indents line. `<<` dedents line. `~` toggles case.

**Tasks:**
1. Extend `WordAdapter`:
   - `DeleteChar()` — `Selection.Delete()`
   - `IndentLine()` — `Selection.ParagraphFormat.LeftIndent += ???` or use `Selection.Range.ConvertToTable` trick. Simplest: `Application.CommandBars.ExecuteMso("IndentIncrease")`
   - `DedentLine()` — `Application.CommandBars.ExecuteMso("IndentDecrease")`
   - `ToggleCase()` — read char, delete, insert toggled version
2. Handle `EditAction(Indent/Dedent/ToggleCase, ...)` in `VimOrchestrator`
3. Handle `>w`, `<w` (indent/dedent with motion)
4. Build and run:
   - `x` → character under cursor deleted
   - `>>` → line indented
   - `<<` → line dedented
   - `~` → character case toggled
   - `>w` → next word indented

**Verification:** All single-key operators work.

**Dependencies:** Day 18.

---

### Day 20 — Word Motions in Word (w/b/e)

**Goal:** Word-level navigation actually moves between words in Word.

**Deliverable:** `w`/`b`/`e` correctly navigate between words in a Word document.

**Tasks:**
1. Extend `WordAdapter`:
   - `MoveToNextWord()` — `Selection.Move(wdWord, 1)`
   - `MoveToPreviousWord()` — `Selection.Move(wdWord, -1)`
   - `MoveToEndOfWord()` — `Selection.EndOf(wdWord, wdMove)`, then `Selection.MoveLeft(wdCharacter, 1)` to land ON the last char
2. Wire `MoveToAction(WordStart/WordEnd, ...)` in `VimOrchestrator`
3. Build and run:
   - Cursor at start of line → `w` → jumps to next word start
   - `b` → jumps to previous word start
   - `e` → jumps to current word end
   - `3w` → jumps 3 words forward
   - `dw` → deletes next word (combine with Day 16)

**Verification:** Word motions navigate and compose with operators correctly.

**Dependencies:** Day 14, Day 16.

---

### Day 21 — Line/Character Motions in Word (0/$/^/f/t)

**Goal:** Line boundaries and character search work in Word.

**Deliverable:** `0`/`$`/`^` navigate within line. `f{char}` finds character forward.

**Tasks:**
1. Extend `WordAdapter`:
   - `MoveToLineStart()` — `Selection.HomeKey(wdLine)`
   - `MoveToLineEnd()` — `Selection.EndOf(wdLine)`
   - `MoveToFirstNonBlank()` — `Selection.HomeKey(wdLine)` then `Selection.Move(wdWord, 1)` minus one char
   - `FindCharForward(char c)` — `Selection.Find.Forward = true`, `Selection.Find.Text = c`, `Selection.Find.Execute`
   - `FindCharBackward(char c)` — same but backward
   - `FindCharTillForward(char c)` — find, then move left 1
2. Wire in `VimOrchestrator`
3. Build and run:
   - `0` → cursor to line start
   - `$` → cursor to line end
   - `^` → cursor to first non-blank
   - `f;` → cursor finds `;` forward
   - `t;` → cursor stops just before `;`
   - `F;` → finds `;` backward

**Verification:** Line and character motions work. Compose with operators (`df;` deletes through `;`).

**Dependencies:** Day 20.

---

### Day 22 — Search (/ ? n N)

**Goal:** Incremental search, next/previous navigation.

**Deliverable:** `/word` highlights matches. `n`/`N` move between matches. `?` searches backward.

**Tasks:**
1. Extend `WordAdapter`:
   - `bool FindText(string text, bool forward, bool matchCase)` — `Selection.Find` API
   - `ClearSearchHighlights()` — clear previous Find formatting
2. Extend `VimOrchestrator` for `SearchAction`:
   - Store last search pattern
   - Call `WordAdapter.FindText(pattern, forward, matchCase)`
3. Handle `n` (repeat search forward) and `N` (repeat search backward):
   - Re-execute last Find with same direction
4. Handle `*` (search word under cursor):
   - Read word at cursor, set as search pattern, execute
5. Build and run:
   - `/test` + Enter → first "test" highlighted
   - `n` → next match
   - `N` → previous match
   - `?test` + Enter → searches backward
   - `*` on a word → searches for that word

**Verification:** Search finds text, n/N navigate, ? searches backward.

**Dependencies:** Day 13, Day 21.

---

### Day 23 — Command-Line Mode (:w :q :wq :e)

**Goal:** Save, quit, and open files via command-line mode.

**Deliverable:** `:w` saves. `:q` closes. `:wq` saves and closes. `:e filename` opens file.

**Tasks:**
1. Extend `VimOrchestrator` for `CommandLineAction`:
   - Parse command string
   - `:w` → `Application.ActiveDocument.Save()`
   - `:q` → `Application.ActiveDocument.Close(wdDoNotSaveChanges)` (or prompt if unsaved)
   - `:wq` → Save then Close
   - `:e {path}` → `Application.Documents.Open(path)`
   - `:q!` → force close without save
   - Unknown command → show error in status bar: "E492: Not a vim command: {cmd}"
2. Display command-line input in status bar while typing (before Enter)
3. Build and run:
   - Type `:w` → Enter → document saved (verify by checking saved state)
   - Type `:q` → Enter → document closes (or Word stays open if last doc)
   - Type `:wq` → saves then closes
   - Type `:e C:\path\to\file.docx` → opens that file
   - Type `:abc` → error shown in status bar

**Verification:** File operations work via command-line mode.

**Dependencies:** Day 13.

---

### Day 24 — Visual Mode in Word

**Goal:** Visual mode selects text, operators work on selection.

**Deliverable:** `v` enters visual, `h/j/k/l` extends selection, `d` deletes selection.

**Tasks:**
1. Extend `WordAdapter`:
   - `SetSelectionAnchor(int position)` — store selection start
   - `ExtendSelectionTo(int position)` — `Selection.SetRange(anchor, position)`
   - `GetSelectionText()` — `Selection.Text`
2. Extend `VimOrchestrator` for Visual mode:
   - On entering Visual mode: record anchor point (current cursor position)
   - Motion keys: extend selection from anchor to new position
   - `d` in Visual: delete selection, return to Normal
   - `y` in Visual: yank selection, return to Normal
   - `c` in Visual: delete selection, enter Insert
   - `v` → toggle back to Normal
3. `V` (Visual Line): select entire lines (snap anchor/start to line boundaries)
4. Build and run:
   - `v` → selection starts at cursor
   - `l`/`h` → selection extends character by character
   - `j`/`k` → selection extends line by line
   - `d` → selected text deleted
   - `v` → `yw` → yanks selection → `p` pastes

**Verification:** Visual mode selection and operators work correctly.

**Dependencies:** Day 16, Day 17, Day 18.

---

### Day 25 — Registers + Counts

**Goal:** Named registers, numbered registers, count prefix for commands.

**Deliverable:** `"a`/`"p` uses named register. `3dd` deletes 3 lines. `5j` moves 5 lines down.

**Tasks:**
1. Wire `RegisterManager` (from Day 9) into `VimOrchestrator`:
   - `"` + letter → set active register
   - Yank → store in active register AND unnamed register
   - Put → retrieve from active register (or unnamed if none specified)
   - `""` → explicitly select unnamed register
   - `"0`-`"9` → numbered registers (auto-populated by yanks)
2. Wire count handling (from Day 6 CommandParser) into `VimOrchestrator`:
   - `3dd` → delete 3 lines
   - `5j` → move 5 lines down
   - `2yw` → yank 2 words
   - `10G` → go to line 10
3. Build and run:
   - `"ayy` → yank line to register a
   - `"ap` → paste from register a
   - `"byw` → yank word to register b
   - `"bp` → paste from register b
   - `3dd` → 3 lines deleted
   - `5j` → cursor moves 5 lines down

**Verification:** Registers and counts work correctly.

**Dependencies:** Day 24.

---

### Day 26 — Undo/Redo (u, Ctrl+r)

**Goal:** Undo and redo via Word's native undo/redo.

**Deliverable:** `u` undoes last change. `Ctrl+r` redoes.

**Tasks:**
1. Extend `WordAdapter`:
   - `Undo()` — `Application.Undo()` (or `Selection.Undo`)
   - `Redo()` — `Application.Redo()` (or `Selection.Redo`)
   - `bool CanUndo()` — check if undo is available
   - `bool CanRedo()` — check if redo is available
2. Handle `UndoAction`/`RedoAction` in `VimOrchestrator`
3. Build and run:
   - Type text → `Esc` → `u` → text removed (undo)
   - `Ctrl+r` → text restored (redo)
   - `dd` → `u` → line restored
   - Multiple undos → multiple redos

**Verification:** Undo and redo work reliably.

**Dependencies:** Day 19.

---

### Day 27 — Dot Repeat + Paragraph Motions

**Goal:** `.` repeats last change. `{`/`}` navigate paragraphs. `%` matches brackets.

**Deliverable:** `ciw` + type + `Esc` + `.` repeats the change. `{`/`}` move by paragraph. `%` jumps to matching bracket.

**Tasks:**
1. Wire `RepeatManager` (from Day 8) into `VimOrchestrator`:
   - After each completed change action: `repeatManager.Record(action)`
   - On `.`: retrieve and re-execute last action
   - For compound actions (e.g., `ciw` + typed text): store the full sequence
2. Extend `WordAdapter`:
   - `MoveToParagraphForward()` — `Selection.Move(wdParagraph, 1)`
   - `MoveToParagraphBackward()` — `Selection.Move(wdParagraph, -1)`
   - `MoveToMatchingBracket()` — `Selection.Find.Text = "\["` or use `Selection.Range.ComputeStatistics` to find matching bracket
3. Build and run:
   - `cw` + type "hello" + `Esc` → `.` → "hello" inserted again
   - `dd` → `.` → another line deleted
   - `}` → next paragraph
   - `{` → previous paragraph
   - Cursor on `(` → `%` → jumps to `)`

**Verification:** Dot repeat and paragraph/bracket motions work.

**Dependencies:** Day 22, Day 26.

---

### Day 28 — Edge Cases + Polish

**Goal:** Handle special keys, empty documents, multiple documents, error recovery.

**Deliverable:** No crashes on edge cases. Graceful error handling.

**Tasks:**
1. Insert mode special keys:
   - `Tab` → insert tab (or spaces, configurable)
   - `Enter` → `Selection.TypeParagraph`
   - `Backspace` → `Selection.TypeBackspace`
   - These should already work since Insert mode passes through to Word, but verify
2. Empty document handling:
   - Verify all commands work when document is empty (no crash)
   - Verify `j`/`k` at document boundaries
3. Multiple document support:
   - Verify WordVim works when multiple documents are open
   - Verify state resets per-document (or shared, decide)
4. Error handling:
   - COM call failures → catch, log to status bar, don't crash Word
   - Invalid command sequence → show error in status bar (e.g., "E488: Trailing characters")
   - Hook failure → disable WordVim, show error message
5. Status bar messages:
   - Show error messages for 2 seconds then revert to mode display
   - Show "Already at newest change" for redundant `u`
6. Build and run: manual testing of edge cases

**Verification:** No crashes. Error messages are user-friendly.

**Dependencies:** Day 27.

---

## Phase C: Release (Days 29-30)

### Day 29 — Installer + CI

**Goal:** Users can install WordVim via installer. CI builds on every push.

**Deliverable:** Inno Setup installer works. GitHub Actions workflow passes.

**Tasks:**
1. Create `installer/Setup.iss` (Inno Setup script):
   - Detect .NET Framework 4.8, prompt to install if missing
   - Copy VSTO add-in DLL + manifest to Program Files
   - Register COM interop (run `regasm`)
   - Create uninstaller
   - Test installer on clean machine (or VM)
2. Create `.github/workflows/build.yml`:
   - Trigger: push to `main`, pull requests
   - Steps: checkout → setup MSBuild → restore NuGet → build → run tests
   - Upload artifact: installer .exe
3. Test CI: push a commit, verify workflow runs green

**Verification:** Installer installs add-in. CI builds and tests pass.

**Dependencies:** Day 28.

---

### Day 30 — Documentation + v1.0 Release

**Goal:** Complete documentation and cut v1.0 release.

**Deliverable:** README complete. GitHub release with installer.

**Tasks:**
1. Update `README.md`:
   - Project description and screenshot/GIF
   - Installation instructions (installer download link)
   - Supported commands table (all Vim commands implemented)
   - Known limitations
   - Contributing guidelines
   - License (MIT)
2. Create `docs/COMMANDS.md`:
   - Complete command reference
   - Mode descriptions
   - Register usage
   - Configuration options (if any)
3. Create GitHub release:
   - Tag: `v1.0.0`
   - Attach installer `.exe` as release asset
   - Write release notes

**Verification:** README is complete. Release is published.

**Dependencies:** Day 29.

---

## Risk Register

| Risk | Day | Impact | Mitigation |
|---|---|---|---|
| Win32 hook crashes Word | 12-13 | **Critical** | Hook callback must be <1ms. Post to main thread for processing. Test early. |
| Word COM threading violation | 14+ | **High** | All COM calls on main thread. Use `Application.Invoke` / `BeginInvoke`. |
| Hook conflicts with other add-ins | 13+ | **Medium** | Graceful fallback: catch hook failure, disable WordVim, show error. |
| `Selection.Find` unreliable for f/t | 21 | **Medium** | Fallback: character-by-character scan via `Range.Characters`. |
| VSTO deployment issues | 29 | **Medium** | Test installer on clean Windows VM. Document manual install as backup. |
| Word 365 updates break COM | Any | **Low** | Use only documented COM APIs. Avoid undocumented features. |

---

## Critical Path

The highest-risk milestones that determine project success:

```
Day 12 (Hook) → Day 13 (Wired) → Day 14 (Suppress + Move)
```

If Day 12 fails (hook doesn't work in Word), the entire architecture must be reconsidered. **Test Day 12 thoroughly before proceeding.**

Days 1-9 (Core) are low risk — pure logic, fully testable.
Days 15-28 (Word features) are medium risk — building on proven COM patterns.
Days 29-30 (Release) are low risk — standard tooling.
