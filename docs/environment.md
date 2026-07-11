# Development Environment Guide

> WordVim development setup for Windows 11 with Neovim, .NET CLI, and Git.

---

## Current System State

| Component | Status | Version/Path |
|-----------|--------|-------------|
| Neovim | Installed | v0.12.4 |
| Git | Installed | 2.54.0 |
| GitHub CLI | Installed | 2.95.0 |
| Node.js | Installed | v24.17.0 |
| npm | Installed | 11.13.0 |
| .NET Framework | Installed | 4.8 |
| .NET SDK | **Not installed** | — |
| `regasm.exe` | Available | `C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe` |
| `extensibility.dll` | Available | `C:\Program Files (x86)\Common Files\microsoft shared\MSEnv\PublicAssemblies\extensibility.dll` |
| Word 365 | Installed | `C:\Program Files\Microsoft Office\Root\Office16\WINWORD.EXE` |

---

## 1. Required Software

### .NET SDK 8.0 (or 9.0)

The .NET SDK provides `dotnet build`, `dotnet test`, and the C# compiler. Required for building WordVim.

**Install:**
```powershell
winget install Microsoft.DotNet.SDK.8
```

**Verify:**
```powershell
dotnet --version    # Should print 8.0.x
dotnet --list-sdks  # Should show 8.0.x
```

**Note:** The SDK targets .NET Core/5+ by default. For WordVim's net472 target, we use the `Microsoft.NETFramework.ReferenceAssemblies` NuGet package to provide reference assemblies without installing the .NET Framework targeting pack separately.

### Git

Already installed. No action needed.

### GitHub CLI

Already installed. No action needed.

**Authenticate if not already done:**
```powershell
gh auth login
```

---

## 2. Optional Software

### Windows Terminal

Recommended for better terminal experience on Windows 11. Already included in Windows 11.

### Inno Setup

Needed only for creating the release installer (Phase C, Day 29). Not needed for development.

```powershell
winget install JRSoftware.InnoSetup
```

### Process Monitor / Process Explorer

Useful for debugging COM registration issues. Attach to Word to see loaded DLLs.

```powershell
winget install Microsoft.Sysinternals.ProcessMonitor
```

### Registry Editor

Already built into Windows (`regedit`). Used for manual registry entries during development.

---

## 3. Neovim Configuration

### Required: LSP for C#

Install `nvim-lspconfig` and the C# language server (`omnisharp` or `roslyn`).

**Option A: Omnisharp (stable)**
```powershell
# Install OmniSharp via .NET global tool
dotnet tool install -g omnisharp
```

**Option B: Roslyn LSP (newer)**
```powershell
# Install via Mason (if using lazy.nvim / packer)
# Or manually download from GitHub releases
```

**Neovim config (`init.lua` or `lspconfig.lua`):**
```lua
local lspconfig = require('lspconfig')

lspconfig.omnisharp.setup {
  cmd = { 'omnisharp' },
  -- Point to the .sln file
  root_dir = lspconfig.util.root_pattern('*.sln', '*.csproj'),
}
```

### Recommended: Neovim Plugins for .NET Development

| Plugin | Purpose |
|--------|---------|
| `nvim-lspconfig` | LSP client configuration |
| `mason.nvim` | Auto-install LSP servers, DAP adapters |
| `mason-lspconfig.nvim` | Bridge between Mason and lspconfig |
| `nvim-dap` | Debug Adapter Protocol client |
| `nvim-dap-ui` | Debug UI |
| `cs` or `omnisharp-extended` | Enhanced C# support |
| `nvim-treesitter` | Syntax highlighting for C# |
| `telescope.nvim` | Fuzzy finder for files/symbols |
| `fugitive.vim` or `lazygit` | Git integration |

### Required: C# Treesitter Parser

```vim
:TSInstall c_sharp
```

---

## 4. Build Tools

### Building

```powershell
# Build the solution
dotnet build WordVim.sln -c Debug

# Build specific project
dotnet build src/WordVim.Core/WordVim.Core.csproj -c Debug
```

### Running Tests

```powershell
# Run all tests
dotnet test

# Run specific test class
dotnet test --filter "FullyQualifiedName~VimStateTests"

# Run with verbose output
dotnet test --verbosity normal
```

### Building for Release

```powershell
dotnet build WordVim.sln -c Release
```

### Cleaning

```powershell
dotnet clean WordVim.sln
```

---

## 5. Debugging Workflow

### Debugging the Core Engine (No Word Required)

The Core engine is pure C# with no Word dependency. Debug via unit tests:

```powershell
# Run tests with debugger attached
dotnet test --filter "FullyQualifiedName~VimStateTests" --logger "console;verbosity=detailed"
```

Or use Neovim DAP (Debug Adapter Protocol):

```vim
" Launch test in debug mode
:DapContinue
" Or set breakpoint and run:
:Bp src/WordVim.Core/VimState.cs:15
```

### Debugging the Word Add-in (Requires Word)

The COM add-in runs inside Word's process. Debugging requires attaching to the Word process.

**Method 1: VS Code / Neovim DAP + vsdbg**

1. Start Word (the add-in will be loaded via registry)
2. In Neovim/VS Code, attach to the `WINWORD.EXE` process
3. Set breakpoints in your add-in code
4. Trigger the breakpoint by pressing keys in Word

**Method 2: `Console.WriteLine` / Debug output**

For quick debugging without attaching a debugger:

```csharp
System.Diagnostics.Debug.WriteLine($"Hook fired: keyCode={keyCode}");
// View in Visual Studio Output window, or DebugView tool from Sysinternals
```

**Method 3: Attach after launch**

```powershell
# Start Word
Start-Process WINWORD.EXE

# Find Word's process ID
$pid = (Get-Process WINWORD).Id

# Attach a debugger (requires vsdbg or similar)
# In Neovim with nvim-dap:
:DapContinue  --  then select "Attach to WINWORD.EXE"
```

### DebugView for Debug Output

Since we're not using Visual Studio, use Sysinternals DebugView to capture `Debug.WriteLine` output:

```powershell
winget install Microsoft.Sysinternals.DebugView
```

Run DebugView as Administrator, enable "Capture Global Win32", and all `Debug.WriteLine` calls will appear.

### Common Debug Scenarios

| Scenario | How to debug |
|----------|-------------|
| Core logic (state machine, parser) | Unit tests — `dotnet test` |
| Hook not firing | DebugView to check if callback is called. Check `nCode == HC_ACTION`. |
| Hook crashes Word | Ensure callback is < 1ms. Check delegate is stored in class field. |
| Word COM error | Check COM thread affinity. Ensure calls are on UI thread. |
| Add-in not loading | Check registry keys. Run `regasm /codebase` again. Check Event Viewer. |

---

## 6. Repository Structure

```
wordvim/
├── .git/
├── .github/
│   └── workflows/
│       └── build.yml              # CI: build + test on push
│
├── docs/
│   ├── architecture.md            # Architecture analysis
│   ├── architecture-review.md     # Critical reassessment
│   ├── milestones.md              # 30-day development plan
│   ├── research.md                # API research with sources
│   └── environment.md             # This file
│
├── src/
│   ├── WordVim.Core/              # Vim logic — NO Word dependency
│   │   ├── WordVim.Core.csproj
│   │   ├── VimMode.cs
│   │   ├── VimState.cs
│   │   ├── KeyHandler.cs
│   │   ├── CommandParser.cs
│   │   ├── VimEngine.cs
│   │   ├── Motions.cs
│   │   ├── Operators.cs
│   │   ├── PendingOperator.cs
│   │   ├── RepeatManager.cs
│   │   ├── RegisterManager.cs
│   │   └── Actions/
│   │       ├── VimAction.cs
│   │       ├── MoveAction.cs
│   │       ├── MoveToAction.cs
│   │       ├── EditAction.cs
│   │       ├── ChangeModeAction.cs
│   │       ├── SearchAction.cs
│   │       ├── CommandLineAction.cs
│   │       └── NoOp.cs
│   │
│   ├── WordVim.Hooks/             # Win32 keyboard hook
│   │   ├── WordVim.Hooks.csproj
│   │   ├── KeyboardHook.cs
│   │   └── HookManager.cs
│   │
│   ├── WordVim.Word/              # Word COM adapter
│   │   ├── WordVim.Word.csproj
│   │   ├── IWordAdapter.cs        # Interface (for testability)
│   │   ├── WordAdapter.cs         # COM implementation
│   │   ├── SelectionHelper.cs
│   │   └── DocumentHelper.cs
│   │
│   └── WordVim.Addin/             # COM Add-in entry point
│       ├── WordVim.Addin.csproj
│       ├── Connect.cs             # IDTExtensibility2 + IRibbonExtensibility
│       ├── VimOrchestrator.cs     # Wires hook → engine → Word adapter
│       ├── ModeIndicator.cs       # Ribbon + status bar updates
│       └── Ribbon.xml             # Ribbon UI definition (embedded resource)
│
├── tests/
│   ├── WordVim.Core.Tests/
│   │   ├── WordVim.Core.Tests.csproj
│   │   ├── VimStateTests.cs
│   │   ├── KeyHandlerTests.cs
│   │   ├── MotionTests.cs
│   │   ├── OperatorTests.cs
│   │   ├── CommandParserTests.cs
│   │   ├── AdvancedMotionTests.cs
│   │   ├── AdvancedOperatorTests.cs
│   │   └── EngineIntegrationTests.cs
│   │
│   └── WordVim.Integration.Tests/  # (optional) Word COM tests
│       └── WordVim.Integration.Tests.csproj
│
├── scripts/
│   ├── build.ps1                  # Build script
│   ├── test.ps1                   # Test script
│   ├── register.ps1               # Register add-in with Word
│   ├── unregister.ps1             # Unregister add-in
│   └── start-word.ps1             # Launch Word with add-in
│
├── installer/
│   ├── Setup.iss                  # Inno Setup script
│   └── register.reg               # Registry entries (for manual install)
│
├── WordVim.sln                    # Solution file
├── .editorconfig                  # Coding standards
├── .gitignore
├── LICENSE                        # MIT
└── README.md
```

---

## 7. Coding Standards

### `.editorconfig`

```ini
root = true

[*]
indent_style = space
indent_size = 4
end_of_line = lf
charset = utf-8
trim_trailing_whitespace = true
insert_final_newline = true

[*.{cs,csx}]
indent_size = 4
csharp_new_line_before_open_brace = all
csharp_new_line_before_else = true
csharp_new_line_before_catch = true
csharp_new_line_before_finally = true
csharp_indent_case_contents = true
csharp_indent_switch_labels = true
csharp_space_after_cast = false
csharp_space_after_keywords_in_control_flow_statements = true
csharp_space_between_method_declaration_parameter_list_parentheses = false
csharp_space_between_method_call_parameter_list_parentheses = false

[*.{csproj,xml,json,yml,yaml}]
indent_size = 2

[*.{ps1,psm1}]
indent_size = 4

[*.md]
trim_trailing_whitespace = false

[Makefile]
indent_style = tab
```

### C# Naming Conventions

| Element | Convention | Example |
|---------|-----------|---------|
| Namespace | PascalCase | `WordVim.Core` |
| Class | PascalCase | `VimState` |
| Interface | IPascalCase | `IWordAdapter` |
| Method | PascalCase | `ProcessKey` |
| Property | PascalCase | `CurrentMode` |
| Field (private) | _camelCase | `_currentMode` |
| Field (constant) | PascalCase | `MaxRegisterCount` |
| Parameter | camelCase | `keyCode` |
| Local variable | camelCase | `previousMode` |
| Enum | PascalCase | `VimMode` |
| Enum value | PascalCase | `VimMode.Normal` |

### File Naming

- One class per file (when practical)
- File name matches class name: `VimState.cs`, `IWordAdapter.cs`
- Test files: `{ClassName}Tests.cs` — e.g., `VimStateTests.cs`
- Actions in `Actions/` subfolder: one file per action type

### Project Naming

- `WordVim.Core` — core logic, no Word dependency
- `WordVim.Hooks` — Win32 interop
- `WordVim.Word` — Word COM adapter
- `WordVim.Addin` — entry point
- `WordVim.Core.Tests` — unit tests

### Comments

- **No comments unless asked.** Code should be self-documenting.
- XML doc comments on public interfaces (`IWordAdapter`, `VimEngine`)
- No inline comments on obvious code

---

## 8. Testing Strategy

### Unit Tests (Core Engine)

**Framework:** NUnit or xUnit (xUnit is simpler with `dotnet test`)

**Location:** `tests/WordVim.Core.Tests/`

**Coverage target:** 90%+ for `WordVim.Core`

**What to test:**
- VimState transitions (every mode pair)
- KeyHandler output (every key in every mode)
- CommandParser parsing (compound commands, counts, text objects)
- Motions (action type and parameters)
- Operators (action type and parameters)
- RepeatManager (record and replay)
- RegisterManager (store and retrieve)

**What NOT to test:**
- Word COM calls (integration tests, separate project)
- Win32 hook callbacks (integration tests, separate project)

### Integration Tests (Word COM)

**Location:** `tests/WordVim.Integration.Tests/`

**When to run:** Manually, not in CI (requires Word installed)

**What to test:**
- Add-in loads in Word without error
- Ribbon appears
- Key hook fires
- Cursor moves correctly
- Text is deleted/yanked/pasted correctly

### Running Tests

```powershell
# Run all unit tests
dotnet test

# Run with coverage
dotnet test --collect:"XPlat Code Coverage"

# Run specific test
dotnet test --filter "FullyQualifiedName~VimStateTests"

# Run tests in watch mode (re-runs on file change)
dotnet watch test
```

### Test Naming Convention

```csharp
[Test]
public void Transition_FromNormalToInsert_ReturnsNormal()
{
    // Arrange
    var state = new VimState();

    // Act
    var previous = state.Transition(VimMode.Insert);

    // Assert
    Assert.That(state.CurrentMode, Is.EqualTo(VimMode.Insert));
    Assert.That(previous, Is.EqualTo(VimMode.Normal));
}
```

Pattern: `Method_Scenario_ExpectedResult`

---

## 9. Scripts

### `scripts/build.ps1`

```powershell
#!/usr/bin/env pwsh
# Build the solution
dotnet build WordVim.sln -c Release
if ($LASTEXITCODE -ne 0) {
    Write-Error "Build failed"
    exit 1
}
Write-Host "Build succeeded" -ForegroundColor Green
```

### `scripts/test.ps1`

```powershell
#!/usr/bin/env pwsh
# Run all unit tests
dotnet test WordVim.sln --verbosity normal
if ($LASTEXITCODE -ne 0) {
    Write-Error "Tests failed"
    exit 1
}
Write-Host "All tests passed" -ForegroundColor Green
```

### `scripts/register.ps1`

```powershell
#!/usr/bin/env pwsh
# Register WordVim add-in with Word (run as Administrator)
$ErrorActionPreference = "Stop"

$regasm = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe"
$dllPath = Join-Path $PSScriptRoot "..\src\WordVim.Addin\bin\Release\net472\WordVim.Addin.dll"

# Register COM server
& $regasm /codebase $dllPath
Write-Host "COM server registered" -ForegroundColor Green

# Add Word add-in registry key
$regPath = "HKCU:\Software\Microsoft\Office\Word\Addins\WordVim.Connect"
New-Item -Path $regPath -Force | Out-Null
Set-ItemProperty -Path $regPath -Name "FriendlyName" -Value "WordVim"
Set-ItemProperty -Path $regPath -Name "Description" -Value "Vim keybindings for Microsoft Word"
Set-ItemProperty -Path $regPath -Name "LoadBehavior" -Value 3 -Type DWord
Write-Host "Add-in registered in Word" -ForegroundColor Green
Write-Host "Restart Word to load the add-in" -ForegroundColor Yellow
```

### `scripts/unregister.ps1`

```powershell
#!/usr/bin/env pwsh
# Unregister WordVim add-in (run as Administrator)
$regasm = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe"
$dllPath = Join-Path $PSScriptRoot "..\src\WordVim.Addin\bin\Release\net472\WordVim.Addin.dll"

& $regasm /u $dllPath
Remove-Item -Path "HKCU:\Software\Microsoft\Office\Word\Addins\WordVim.Connect" -ErrorAction SilentlyContinue
Write-Host "Add-in unregistered" -ForegroundColor Green
```

---

## 10. Release Workflow

### Version Numbering

Follow [SemVer](https://semver.org/): `MAJOR.MINOR.PATCH`

- `MAJOR` — breaking changes (e.g., configuration format change)
- `MINOR` — new features (e.g., new Vim commands)
- `PATCH` — bug fixes

### Release Steps

```powershell
# 1. Ensure all tests pass
.\scripts\test.ps1

# 2. Build release
.\scripts\build.ps1

# 3. Create git tag
git tag -a v1.0.0 -m "Release v1.0.0"
git push origin v1.0.0

# 4. Create GitHub release with installer
gh release create v1.0.0 `
    --title "v1.0.0" `
    --notes "Release notes here" `
    .\installer\Output\WordVimSetup.exe

# 5. Or create release from tag
gh release create v1.0.0 --generate-notes
```

### Release Checklist

- [ ] All unit tests pass (`dotnet test`)
- [ ] Build succeeds in Release mode (`dotnet build -c Release`)
- [ ] Manual testing in Word 365 on Windows 11
- [ ] Installer tested on clean machine (or VM)
- [ ] README updated with new features
- [ ] CHANGELOG updated
- [ ] Git tag created
- [ ] GitHub release published with installer attached

---

## 11. First-Time Setup Checklist

```powershell
# 1. Install .NET SDK
winget install Microsoft.DotNet.SDK.8

# 2. Verify installation
dotnet --version

# 3. Install Neovim LSP (if not already configured)
# Add nvim-lspconfig + omnisharp to your Neovim config

# 4. Install treesitter C# parser
nvim --headless "+TSInstall c_sharp" +q

# 5. Clone the repository
git clone https://github.com/spidychoipro/wordvim.git
cd wordvim

# 6. Build the solution
dotnet build

# 7. Run tests
dotnet test

# 8. Start developing
nvim .
```

---

## 12. Quick Reference

| Task | Command |
|------|---------|
| Build | `dotnet build WordVim.sln` |
| Build Release | `dotnet build WordVim.sln -c Release` |
| Test | `dotnet test` |
| Test specific | `dotnet test --filter "ClassName=VimStateTests"` |
| Clean | `dotnet clean WordVim.sln` |
| Watch test | `dotnet watch test` |
| Register add-in | `.\scripts\register.ps1` (as Admin) |
| Unregister | `.\scripts\unregister.ps1` (as Admin) |
| Create branch | `git checkout -b feature/my-feature` |
| Commit | `git commit -m "feat: add visual mode"` |
| Push | `git push origin feature/my-feature` |
| Create PR | `gh pr create` |
| View PR | `gh pr view` |
| Merge PR | `gh pr merge` |
| Create release | `gh release create v1.0.0 --generate-notes` |
