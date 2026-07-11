# MVP: Architecture Proof

> One day. One question answered: Does the architecture work?

---

## The Question

Can a native COM Add-in built with `dotnet build` (no Visual Studio) intercept keystrokes in Word via WH_KEYBOARD, suppress them, and move the cursor via Word COM?

If yes → the project is viable.
If no → we need a different approach.

---

## What the MVP Does

**Exactly 4 things. Nothing else.**

1. **Loads in Word** — open Word, add-in loads without error
2. **Intercepts keystrokes** — WH_KEYBOARD callback fires for every key
3. **Tracks two modes** — Normal and Insert
4. **Moves cursor** — `h`/`j`/`k`/`l` move the cursor in Normal mode

### What the MVP does NOT do

- No `d`/`y`/`c` operators
- No `w`/`b`/`e` word motions
- No search
- No command-line mode (`:w`, `:q`)
- No Ribbon mode indicator
- No visual mode
- No registers
- No dot repeat
- No undo/redo
- No installer

---

## What Success Looks Like

### Manual test script (run in order):

```
1.  Open Word. Create a new blank document.
    → No error dialog. No crash. Word opens normally.

2.  Type "Hello World" in the document.
    → Text appears (Word is in Insert mode by default).

3.  Press Escape.
    → Nothing visible happens. No character appears.
    → Internally: state changes to Normal.

4.  Press 'l' five times.
    → Cursor moves right 5 characters.

5.  Press 'h' three times.
    → Cursor moves left 3 characters.

6.  Press 'j' two times.
    → Cursor moves down 2 lines (or to last line if only 1 line).

7.  Press 'k' one time.
    → Cursor moves up 1 line.

8.  Press 'i'.
    → Nothing visible happens. No character appears.
    → Internally: state changes to Insert.

9.  Type " Goodbye".
    → Text appears at cursor position.

10. Press Escape.
    → Back to Normal. No character appears.

11. Press 'l' twice.
    → Cursor moves right 2 characters.

12. Close Word. Do not save.
    → Word closes normally.
```

### If all 12 steps pass → Architecture is proven.

---

## What the MVP Contains

### Files (8 source files + 2 test files)

```
src/
├── WordVim.Core/
│   ├── VimMode.cs              # enum: Normal, Insert
│   ├── VimState.cs             # mode tracking + transitions
│   ├── VimAction.cs            # action types (Move, ChangeMode, NoOp)
│   └── KeyHandler.cs           # keystroke → action mapping
│
└── WordVim.Addin/
    ├── Connect.cs              # IDTExtensibility2 — add-in entry point
    ├── KeyboardHook.cs         # WH_KEYBOARD hook
    ├── WordAdapter.cs          # Word COM cursor movement
    └── VimOrchestrator.cs      # wires hook → engine → adapter

tests/
└── WordVim.Core.Tests/
    ├── VimStateTests.cs        # mode transition tests
    └── KeyHandlerTests.cs      # keystroke mapping tests
```

**8 source files. 2 test files. ~400 lines of code total.**

---

## What Each File Does

### `VimMode.cs`

```csharp
namespace WordVim.Core
{
    public enum VimMode
    {
        Normal,
        Insert
    }
}
```

### `VimState.cs`

Tracks current mode. One method: `Transition(VimMode target)` returns previous mode.

### `VimAction.cs`

Three types:
- `MoveAction(int direction)` — h=Left, j=Down, k=Up, l=Right
- `ChangeModeAction(VimMode mode)` — switch mode
- `NoOp()` — do nothing (let keystroke pass through)

### `KeyHandler.cs`

One method: `HandleKey(VimMode mode, int keyCode, bool shift, bool ctrl)` returns `VimAction`.

- Normal mode: h→Move(-1,0), j→Move(0,1), k→Move(0,-1), l→Move(1,0), i→ChangeMode(Insert), Esc→ChangeMode(Normal)
- Insert mode: every key → NoOp

### `Connect.cs`

Implements `IDTExtensibility2`. On startup: creates `VimOrchestrator`, installs hook. On shutdown: removes hook.

### `KeyboardHook.cs`

P/Invoke for `SetWindowsHookEx(WH_KEYBOARD=2, ...)`. Callback: extracts key code, passes to orchestrator. Returns 1 to suppress in Normal mode, or calls `CallNextHookEx` in Insert mode.

### `WordAdapter.cs`

One method per direction: `MoveLeft()`, `MoveRight()`, `MoveUp()`, `MoveDown()`. Each calls `Selection.MoveLeft/Right/Up/Down(wdCharacter, 1)`.

### `VimOrchestrator.cs`

Holds `VimState`, `KeyHandler`, `WordAdapter`. On keystroke: call handler, get action, execute action. The glue between hook and Word.

---

## Project Files

### `WordVim.Core.csproj`

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net472</TargetFramework>
    <OutputType>Library</OutputType>
    <LangVersion>latest</LangVersion>
  </PropertyGroup>
  <ItemGroup>
    <PackageReference Include="Microsoft.NETFramework.ReferenceAssemblies" Version="1.0.3" PrivateAssets="all" />
  </ItemGroup>
</Project>
```

### `WordVim.Addin.csproj`

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net472</TargetFramework>
    <OutputType>Library</OutputType>
    <LangVersion>latest</LangVersion>
    <EnableComHosting>true</EnableComHosting>
  </PropertyGroup>
  <ItemGroup>
    <PackageReference Include="Microsoft.Office.Interop.Word" Version="15.0.4420.1017" />
    <PackageReference Include="Microsoft.NETFramework.ReferenceAssemblies" Version="1.0.3" PrivateAssets="all" />
    <ProjectReference Include="..\WordVim.Core\WordVim.Core.csproj" />
    <Reference Include="extensibility">
      <HintPath>C:\Program Files (x86)\Common Files\microsoft shared\MSEnv\PublicAssemblies\extensibility.dll</HintPath>
    </Reference>
  </ItemGroup>
</Project>
```

### `WordVim.Core.Tests.csproj`

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net472</TargetFramework>
    <OutputType>Library</OutputType>
    <LangVersion>latest</LangVersion>
    <IsPackable>false</IsPackable>
  </PropertyGroup>
  <ItemGroup>
    <PackageReference Include="NUnit" Version="3.14.0" />
    <PackageReference Include="NUnit3TestAdapter" Version="4.5.0" />
    <PackageReference Include="Microsoft.NET.Test.Sdk" Version="17.8.0" />
    <PackageReference Include="Microsoft.NETFramework.ReferenceAssemblies" Version="1.0.3" PrivateAssets="all" />
    <ProjectReference Include="..\..\src\WordVim.Core\WordVim.Core.csproj" />
  </ItemGroup>
</Project>
```

---

## Build & Test Commands

```powershell
# Build everything
dotnet build WordVim.sln

# Run unit tests
dotnet test

# Register add-in (as Administrator)
.\scripts\register.ps1

# Start Word to test
Start-Process WINWORD.EXE
```

---

## What Proves Architecture Failure

If any of these happen → the approach needs rethinking:

| Failure | What it means | Alternative |
|---------|--------------|-------------|
| `dotnet build` fails to compile net472 | .NET SDK can't target Framework | Check `Microsoft.NETFramework.ReferenceAssemblies` package |
| `regasm` fails to register the DLL | COM interop broken | Check `[ComVisible]` attributes, try `/codebase` flag |
| Word doesn't load the add-in | Registry key wrong or add-in rejected | Check `HKCU\Software\Microsoft\Office\Word\Addins\`, check Event Viewer |
| WH_KEYBOARD callback never fires | Hook not installed correctly | Check `GetCurrentThreadId()` (not `Thread.ManagedThreadId`), check delegate is stored |
| WH_KEYBOARD fires but can't suppress | Return value not working | Ensure returning `1` instead of calling `CallNextHookEx` |
| Word COM call throws | Thread affinity or COM error | Ensure calls are on UI thread, check COM object is not null |
| Cursor doesn't move | Word COM API call wrong | Check `Selection` object, try `MoveLeft(wdCharacter, 1)` in VBA editor first |

---

## Day Plan

| Time block | Task | Deliverable |
|-----------|------|-------------|
| **Hour 1** | Project structure + build | `dotnet build` succeeds, 0 errors |
| **Hour 2** | VimMode + VimState + VimAction | Mode enum, state transitions |
| **Hour 3** | KeyHandler | Keystroke → action mapping |
| **Hour 4** | Unit tests | `dotnet test` passes |
| **Hour 5** | Connect + KeyboardHook | Add-in loads, hook fires |
| **Hour 6** | WordAdapter + VimOrchestrator | h/j/k/l moves cursor in Word |
| **Hour 7** | Insert mode + Esc | i → type → Esc → Normal |
| **Hour 8** | Integration testing + register script | Full manual test passes |

**8 hours. If it takes longer → scope is wrong, cut something.**

---

## After the MVP

If the MVP passes all 12 steps, the architecture is proven. Then:

| Priority | Feature | Builds on |
|----------|---------|-----------|
| 1 | `w`/`b`/`e` word motions | WordAdapter |
| 2 | `d` operator | WordAdapter (delete) |
| 3 | `y`/`p` yank/put | WordAdapter (read/write text) |
| 4 | `c` operator | d + Insert mode |
| 5 | Ribbon mode indicator | IRibbonExtensibility |
| 6 | `0`/`$` line motions | WordAdapter |
| 7 | Search (`/`/`?`) | WordAdapter (Find) |
| 8 | Command-line mode (`:w`/`:q`) | Word COM save/close |

Each of these is a separate day in the milestone plan. The MVP is Day 0 — the foundation everything else stands on.
