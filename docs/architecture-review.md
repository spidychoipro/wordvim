# Architecture Review: Critical Reassessment

> Challenging the previous recommendation. Honest about what was wrong.

---

## Previous Recommendation (Now Questioned)

The previous architecture proposed **VSTO (C# / .NET Framework 4.8)** as the add-in type, with Visual Studio as the development environment.

**The problem:** The user develops with Neovim + .NET CLI. Visual Studio is not in their workflow.

---

## What Was Wrong

### 1. VSTO was presented as the only option

The architecture document compared VSTO against Office.js and VBA, but **did not consider native COM Add-ins** as a separate category. This was a significant omission.

### 2. Visual Studio was assumed necessary

The architecture stated "VSTO project type requires VS" without investigating whether the same goals could be achieved with a simpler add-in type that doesn't need VS.

### 3. The comparison was incomplete

The comparison table missed the **native COM Add-in** approach entirely, which is:
- Buildable with `dotnet build` (no VS IDE)
- Capable of the same keyboard interception (WH_KEYBOARD hook)
- Capable of the same Word COM access
- Simpler than VSTO (no VSTO runtime dependency)

---

## The Four Approaches Compared

### Approach A: VSTO Add-in

| Criterion | Assessment |
|---|---|
| **What it is** | .NET Framework assembly wrapped by VSTO runtime. Provides `ThisAddIn` class with `Startup`/`Shutdown` events. |
| **Keyboard interception** | WH_KEYBOARD hook (type 2) — works |
| **Build without VS IDE** | **No.** VSTO targets (`Microsoft.VisualStudio.Tools.Office.targets`) only ship with Visual Studio or Build Tools with `OfficeBuildTools` workload. `dotnet build` cannot produce VSTO deployment artifacts because it runs MSBuild Core, not Full. |
| **Build with Build Tools** | Yes, via `msbuild.exe` from Build Tools — but requires installing VS Build Tools with Office workload |
| **Ribbon UI** | Built-in Ribbon designer in VS (or manual XML via `ThisAddIn.CustomTaskPanes`) |
| **Deployment** | ClickOnce or MSI. VSTO runtime must be installed on target machine. |
| **Dependencies** | VSTO Runtime, .NET Framework 4.x, Office PIAs |
| **Maintainability** | Good — standard C# with VS templates |
| **Long-term viability** | Maintenance mode — no new features, but still supported |

**Verdict:** Works perfectly, but requires Visual Studio Build Tools (not just .NET SDK). Not compatible with a pure Neovim + `dotnet build` workflow.

### Approach B: Native COM Add-in (C# via COM Interop)

| Criterion | Assessment |
|---|---|
| **What it is** | .NET Framework class library that implements `IDTExtensibility2` (and optionally `IRibbonExtensibility`). Registered via `regasm`. Word loads it as a COM server. |
| **Keyboard interception** | WH_KEYBOARD hook (type 2) — works identically to VSTO |
| **Build without VS IDE** | **Yes.** `dotnet build` targeting `net472` produces the DLL. No VSTO targets needed. |
| **Ribbon UI** | Via `IRibbonExtensibility` interface + Ribbon XML embedded as resource. No designer needed. |
| **Deployment** | `regasm /codebase` + manual registry key under `HKCU\Software\Microsoft\Office\Word\Addins\`. Simple script. |
| **Dependencies** | .NET Framework 4.x, Office PIAs (NuGet: `Microsoft.Office.Interop.Word`) |
| **Maintainability** | Good — plain C# class library, no magic |
| **Long-term viability** | Same as VSTO (COM add-in model is stable, just not evolving) |

**Verdict:** Achieves the same goals as VSTO but without requiring Visual Studio. Buildable with `dotnet build` + `regasm`.

### Approach C: Office.js Web Add-in

| Criterion | Assessment |
|---|---|
| **What it is** | Web application (HTML/CSS/JS) running in an iframe inside Word. |
| **Keyboard interception** | **None.** Cannot intercept keystrokes in the document body. Architectural limitation. |
| **Build without VS IDE** | Yes — `npm` + any text editor. VS deprecated for this starting VS 2026. |
| **Maintainability** | Good — standard web development |
| **Long-term viability** | Microsoft's strategic direction, but **cannot achieve the project goal** |

**Verdict:** Best developer experience, but fundamentally incapable of Vim keybinding interception.

### Approach D: Shared Add-in

| Criterion | Assessment |
|---|---|
| **Status** | **Deprecated.** Template removed from modern Visual Studio. |
| **What it was** | .NET COM add-in template (precursor to VSTO). Simpler than VSTO but required VS for project creation. |

**Verdict:** Dead. Not a viable option.

---

## Head-to-Head: VSTO vs Native COM Add-in

| Aspect | VSTO | Native COM Add-in |
|--------|------|-------------------|
| **Keyboard hook** | WH_KEYBOARD (type 2) | WH_KEYBOARD (type 2) — identical |
| **Word COM access** | Full Word Object Model | Full Word Object Model — identical |
| **Build CLI** | `msbuild` from Build Tools (not `dotnet build`) | **`dotnet build`** targeting net472 |
| **VS IDE required?** | No (Build Tools sufficient) | **No** |
| **VS Build Tools required?** | Yes (OfficeBuildTools workload) | **No** |
| **.NET SDK sufficient?** | No | **Yes** |
| **Ribbon UI** | VS Ribbon designer or manual | `IRibbonExtensibility` + XML |
| **Registration** | VSTO deployment handles it | `regasm /codebase` + registry key |
| **Runtime dependency** | VSTO Runtime must be installed | .NET Framework only (already on Windows) |
| **Project creation** | VS template or manual csproj | Manual csproj (trivial) |
| **Complexity** | Higher (VSTO runtime layers) | Lower (direct COM interop) |
| **Debugging** | VS debugger (F5) | Any .NET debugger attached to Word process |

---

## The Answer: Native COM Add-in Is the Right Choice

### Why the previous VSTO recommendation was incomplete

VSTO adds value over native COM Add-ins in three areas:
1. Ribbon designer (visual drag-and-drop)
2. ClickOnce deployment
3. Document-level customizations

**WordVim needs none of these.** It needs:
- Keyboard hook → same in both
- Word COM access → same in both
- Simple Ribbon (one label showing mode) → `IRibbonExtensibility` is sufficient
- Simple deployment → `regasm` script is sufficient

VSTO's extras are unnecessary complexity for this project.

### Why native COM Add-in is correct for WordVim

1. **Builds with `dotnet build`** — compatible with Neovim + .NET CLI workflow
2. **No Visual Studio dependency** — not even Build Tools
3. **Same keyboard interception** — WH_KEYBOARD hook works identically
4. **Same Word COM access** — identical interop assemblies
5. **Simpler** — no VSTO runtime, no VSTO targets, no VSTO deployment artifacts
6. **Proven** — [NetOffice](https://github.com/NetOfficeFw/NetOffice) (759 stars) uses this exact approach
7. **Lower barrier to contribution** — contributors don't need VS installed

### What the csproj looks like

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net472</TargetFramework>
    <OutputType>Library</OutputType>
    <LangVersion>latest</LangVersion>
    <EnableComHosting>true</EnableComHosting>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="Microsoft.Office.Interop.Word" Version="15.*" />
    <Reference Include="extensibility">
      <HintPath>$(MSBuildProgramFiles32)\Common Files\microsoft shared\MSEnv\PublicAssemblies\extensibility.dll</HintPath>
    </Reference>
  </ItemGroup>
</Project>
```

This builds with `dotnet build`. No VSTO targets, no Build Tools, no VS.

### What the class looks like

```csharp
using Extensibility;
using Microsoft.Office.Core;

[ProgId("WordVim.Connect")]
[ComVisible(true)]
[Guid("...")]
[ClassInterface(ClassInterfaceType.None)]
public class Connect : IDTExtensibility2, IRibbonExtensibility
{
    private Word.Application _application;

    public void OnConnection(object application, ext_ConnectMode connectMode,
        object addInInst, ref Array custom)
    {
        _application = (Word.Application)application;
        // Install WH_KEYBOARD hook here
    }

    public void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom)
    {
        // Uninstall hook here
    }

    public void OnAddInsUpdate(ref Array custom) { }
    public void OnStartupComplete(ref Array custom) { }
    public void OnBeginShutdown(ref Array custom) { }

    public string GetCustomUI(string ribbonID)
    {
        // Return Ribbon XML from embedded resource
    }

    // Ribbon callback: update mode label
    public string GetModeLabel(IRibbonControl control)
    {
        return _currentMode;
    }
}
```

### What deployment looks like

```bash
# Build
dotnet build -c Release

# Register
regasm /codebase bin\Release\net472\WordVim.dll

# Add Office add-in registry key (one-time)
reg add "HKCU\Software\Microsoft\Office\Word\Addins\WordVim.Connect" /v FriendlyName /t REG_SZ /d "WordVim"
reg add "HKCU\Software\Microsoft\Office\Word\Addins\WordVim.Connect" /v LoadBehavior /t REG_DWORD /d 3
```

---

## Impact on Milestones

The milestones document needs updating. The key changes:

| Day | Old (VSTO) | New (Native COM) |
|-----|-----------|------------------|
| Day 1 | Create VSTO projects | Create .NET Framework class library (`dotnet new classlib -f net472`) |
| Day 10 | VSTO `ThisAddIn.cs` | `Connect : IDTExtensibility2` class |
| Day 11 | VSTO Ribbon designer | `IRibbonExtensibility` + XML resource |
| Day 12 | Hook setup same | Hook setup same (identical) |
| Day 29 | Inno Setup for VSTO | `regasm` script + optional Inno Setup |

The Core engine (Days 1-9) is unchanged — it has no dependency on the add-in type.

---

## Real-World Precedent

| Project | Approach | Build System | VS Required? |
|---------|----------|-------------|-------------|
| [NetOffice](https://github.com/NetOfficeFw/NetOffice) | COM Add-in (IDTExtensibility2) | `build.cmd` / Cake | No |
| [NetOffice Samples](https://github.com/NetOfficeFw/Samples) | COM Add-in | Cake / AppVeyor | No |
| [Excel-DNA](https://github.com/Excel-DNA/ExcelDna) | Native .xll | `dotnet build` | No |
| [DotnetExcelComAddIn](https://github.com/HCarlb/DotnetExcelComAddIn) | COM Add-in | `build.bat` / `dotnet build` | No |
| [VimWord](https://github.com/cxw42/VimWord) | VBA (.dotm) | make | No |

---

## Remaining Risks

| Risk | Severity | Mitigation |
|------|----------|-----------|
| `regasm` is .NET Framework only | Low | .NET Framework 4.x is pre-installed on all supported Windows versions |
| `extensibility.dll` location varies | Low | Can reference via NuGet or fixed path |
| WH_KEYBOARD callback timing | Medium | Same risk as VSTO — keep callback < 1ms |
| Manual registry entries | Low | Script it. One-time setup per machine. |
| No VS debugger (F5) | Low | Attach any .NET debugger to Word process (VS Code, JetBrains, etc.) |

---

## Conclusion

**The previous recommendation of VSTO was not wrong, but it was incomplete.** VSTO works, but it unnecessarily couples the project to Visual Studio Build Tools. A native COM Add-in achieves the same goals with a simpler toolchain that is fully compatible with Neovim + `dotnet build`.

**New recommendation: Native COM Add-in (C# via COM Interop).**

This is the correct architecture for a project that values:
- CLI-first development
- Minimal toolchain dependencies
- Long-term maintainability
- Low barrier to contribution
