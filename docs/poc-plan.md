# Proof of Concept Plan

> Validate 8 assumptions. Nothing else. No architecture, no features, no Vim logic.

---

## Assumptions

| # | Assumption | Depends on |
|---|-----------|-----------|
| 1 | Native COM Add-in compiles without Visual Studio | — |
| 2 | .NET CLI builds it | — |
| 3 | `regasm` registers it | 1, 2 |
| 4 | Word loads it | 3 |
| 5 | WH_KEYBOARD can be installed | 4 |
| 6 | Hook receives events inside Word | 5 |
| 7 | Keystrokes can be suppressed | 6 |
| 8 | Word COM moves the cursor | 4 |

Assumptions 1 and 2 are the same experiment. Assumptions 5-7 are sequential. Assumption 8 is independent of 5-7.

---

## Step 1: Build Without Visual Studio

**Assumptions tested:** 1, 2

**Expected result:** `dotnet build` produces `PoC.dll` with 0 errors.

**Failure condition:** Build fails with missing reference assemblies, missing framework, or compile errors.

**Experiment:**

Create `poc/PoC.csproj`:
```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net472</TargetFramework>
    <OutputType>Library</OutputType>
    <LangVersion>latest</LangVersion>
    <EnableComHosting>true</EnableComHosting>
  </PropertyGroup>
  <ItemGroup>
    <PackageReference Include="Microsoft.NETFramework.ReferenceAssemblies" Version="1.0.3" PrivateAssets="all" />
    <Reference Include="extensibility">
      <HintPath>C:\Program Files (x86)\Common Files\microsoft shared\MSEnv\PublicAssemblies\extensibility.dll</HintPath>
    </Reference>
  </ItemGroup>
</Project>
```

Create `poc/Connect.cs`:
```csharp
using System;
using System.Runtime.InteropServices;
using Extensibility;

[ProgId("WordVimPoC.Connect")]
[ComVisible(true)]
[Guid("B1E24F3A-5C4D-4E6F-8A9B-0C1D2E3F4A5B")]
[ClassInterface(ClassInterfaceType.None)]
public class Connect : IDTExtensibility2
{
    public void OnConnection(object application, ext_ConnectMode connectMode,
        object addInInst, ref Array custom)
    {
        System.Diagnostics.Debug.WriteLine("WordVimPoC: Connected");
    }
    public void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom) { }
    public void OnAddInsUpdate(ref Array custom) { }
    public void OnStartupComplete(ref Array custom) { }
    public void OnBeginShutdown(ref Array custom) { }
}
```

**Run:**
```powershell
cd poc
dotnet build
```

**Pass:** `PoC.dll` exists in `bin/Debug/net472/`.
**Fail:** Any build error.

**Code size:** ~25 lines.

---

## Step 2: Register With regasm

**Assumption tested:** 3

**Expected result:** `regasm /codebase` exits with code 0. Registry key exists.

**Failure condition:** regasm error, DLL not found, access denied.

**Experiment:**

```powershell
# Register (run as Administrator)
$regasm = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe"
$dll = "poc\bin\Debug\net472\PoC.dll"
& $regasm /codebase $dll

# Verify registry key
Get-ItemProperty "HKCU:\Software\Microsoft\Office\Word\Addins\WordVimPoC.Connect" -ErrorAction Stop
```

**Pass:** regasm exits 0, registry key exists with FriendlyName and LoadBehavior.
**Fail:** regasm error or registry key missing.

**Note:** The Office add-in registry key (`HKCU\Software\Microsoft\Office\Word\Addins\...`) must be created manually. `regasm` only handles COM registration (CLSID), not Office add-in registration.

Create `poc/register.ps1`:
```powershell
#!/usr/bin/env pwsh
$ErrorActionPreference = "Stop"
$regasm = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe"
$dll = Join-Path $PSScriptRoot "bin\Debug\net472\PoC.dll"

& $regasm /codebase $dll

$regPath = "HKCU:\Software\Microsoft\Office\Word\Addins\WordVimPoC.Connect"
New-Item -Path $regPath -Force | Out-Null
Set-ItemProperty -Path $regPath -Name "FriendlyName" -Value "WordVim PoC"
Set-ItemProperty -Path $regPath -Name "Description" -Value "Proof of Concept"
Set-ItemProperty -Path $regPath -Name "LoadBehavior" -Value 3 -Type DWord

Write-Host "Registered. Restart Word to test." -ForegroundColor Green
```

**Code size:** ~15 lines.

---

## Step 3: Word Loads It

**Assumption tested:** 4

**Expected result:** Word opens without error. Add-in appears in File → Options → Add-ins → COM Add-ins.

**Failure condition:** Word shows error dialog. Add-in not listed. Word crashes.

**Experiment:**

1. Close Word completely (check Task Manager — no `WINWORD.EXE`).
2. Run `.\register.ps1` as Administrator.
3. Open Word.
4. Go to File → Options → Add-ins → Manage: COM Add-ins → Go.
5. Look for "WordVim PoC" in the list.

**Pass:** "WordVim PoC" appears with checkbox checked.
**Fail:** Not listed, or error on startup.

**Check Debug output:** Open DebugView (Sysinternals) before opening Word. The `Debug.WriteLine("WordVimPoC: Connected")` should appear.

**No additional code needed.**

---

## Step 4: Install WH_KEYBOARD

**Assumption tested:** 5

**Expected result:** `SetWindowsHookEx` returns a non-zero handle. No crash.

**Failure condition:** Returns `IntPtr.Zero`. Word crashes. Hook callback throws.

**Experiment:**

Add to `Connect.cs`:
```csharp
using System.Diagnostics;
using System.Runtime.InteropServices;

private delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);
private static HookProc _proc = HookCallback;
private static IntPtr _hookId = IntPtr.Zero;

[DllImport("user32.dll", SetLastError = true)]
static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hMod, uint dwThreadId);

[DllImport("user32.dll")]
static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

[DllImport("user32.dll")]
[return: MarshalAs(UnmanagedType.Bool)]
static extern bool UnhookWindowsHookEx(IntPtr hhk);

[DllImport("kernel32.dll")]
static extern uint GetCurrentThreadId();

private const int WH_KEYBOARD = 2;

static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
{
    return CallNextHookEx(IntPtr.Zero, nCode, wParam, lParam);
}

public void OnConnection(object application, ext_ConnectMode connectMode,
    object addInInst, ref Array custom)
{
    uint threadId = GetCurrentThreadId();
    _hookId = SetWindowsHookEx(WH_KEYBOARD, _proc, IntPtr.Zero, threadId);
    Debug.WriteLine($"WordVimPoC: Hook installed, handle={_hookId}");
}

public void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom)
{
    if (_hookId != IntPtr.Zero)
        UnhookWindowsHookEx(_hookId);
}
```

**Run:** Build, register, restart Word. Check DebugView for handle value.

**Pass:** Handle is non-zero (not `IntPtr.Zero`). Word doesn't crash.
**Fail:** Handle is zero. Word crashes. Callback throws exception.

**Code size:** ~40 lines (cumulative).

---

## Step 5: Hook Receives Events

**Assumption tested:** 6

**Expected result:** Callback fires on every keystroke. DebugView shows log entries.

**Failure condition:** Callback never fires. Fires for other apps but not Word.

**Experiment:**

Modify `HookCallback`:
```csharp
private static int _count = 0;

static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
{
    if (nCode == 0) // HC_ACTION
    {
        _count++;
        Debug.WriteLine($"WordVimPoC: Key #{_count}, wParam={wParam}");
    }
    return CallNextHookEx(IntPtr.Zero, nCode, wParam, lParam);
}
```

**Run:** Build, register, restart Word. Open DebugView. Type in Word.

**Pass:** DebugView shows "Key #1", "Key #2", etc. as you type.
**Fail:** No output. Output only appears when typing outside Word.

**Code size:** ~5 lines changed.

---

## Step 6: Suppress Keystrokes

**Assumption tested:** 7

**Expected result:** Returning `1` instead of `CallNextHookEx` prevents the character from appearing in Word.

**Failure condition:** Character appears despite returning 1.

**Experiment:**

Modify `HookCallback` to suppress the letter 'a' (0x41):
```csharp
static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
{
    if (nCode == 0)
    {
        int vk = Marshal.ReadInt32(lParam); // vkCode is first field
        if (vk == 0x41 && wParam == (IntPtr)0x0100) // 'a' + WM_KEYDOWN
        {
            Debug.WriteLine("WordVimPoC: Suppressing 'a'");
            return (IntPtr)1; // suppress
        }
    }
    return CallNextHookEx(IntPtr.Zero, nCode, wParam, lParam);
}
```

**Run:** Build, register, restart Word. Click in document. Press 'a'.

**Pass:** Letter 'a' does not appear. Other letters work normally. DebugView shows "Suppressing 'a'".
**Fail:** Letter 'a' still appears.

**Code size:** ~5 lines changed.

---

## Step 7: Move Cursor via Word COM

**Assumption tested:** 8

**Expected result:** `Selection.MoveRight(wdCharacter, 1)` moves cursor one character right.

**Failure condition:** COM exception. Cursor doesn't move. Wrong object referenced.

**Experiment:**

Add Office Interop reference to `PoC.csproj`:
```xml
<PackageReference Include="Microsoft.Office.Interop.Word" Version="15.0.4420.1017" />
```

Modify `Connect.cs`:
```csharp
using Word = Microsoft.Office.Interop.Word;

private Word.Application _app;

public void OnConnection(object application, ext_ConnectMode connectMode,
    object addInInst, ref Array custom)
{
    _app = (Word.Application)application;

    // Move cursor right by 1 character
    try
    {
        _app.Selection.MoveRight(Word.WdUnits.wdCharacter, 1, Word.WdMovementType.wdMove);
        Debug.WriteLine("WordVimPoC: Cursor moved right");
    }
    catch (Exception ex)
    {
        Debug.WriteLine($"WordVimPoC: COM error: {ex.Message}");
    }

    // Install hook (from Step 4)
    uint threadId = GetCurrentThreadId();
    _hookId = SetWindowsHookEx(WH_KEYBOARD, _proc, IntPtr.Zero, threadId);
    Debug.WriteLine($"WordVimPoC: Hook handle={_hookId}");
}
```

**Run:** Build, register, restart Word. Place cursor at start of a word. Check if it moved.

**Pass:** Cursor moved right by one character. DebugView shows "Cursor moved right". No COM error.
**Fail:** COM exception. Cursor didn't move. Wrong position.

**Code size:** ~15 lines changed.

---

## Summary

| Step | Assumption | Test | Pass criterion |
|------|-----------|------|----------------|
| 1 | 1, 2 | `dotnet build` | DLL exists, 0 errors |
| 2 | 3 | `regasm /codebase` + registry | Exit 0, key exists |
| 3 | 4 | Open Word | Add-in listed in COM Add-ins |
| 4 | 5 | `SetWindowsHookEx` | Non-zero handle |
| 5 | 6 | Type in Word | Callback fires |
| 6 | 7 | Press 'a' in Word | Character suppressed |
| 7 | 8 | Cursor movement | Cursor moves right |

**Total code:** ~70 lines across all steps.
**Total files:** 1 csproj, 1 cs, 1 ps1.
**Time estimate:** 2-3 hours if everything works. 30 minutes per debugging step if something fails.

---

## If Any Step Fails

| Failure | Likely cause | Fix |
|---------|-------------|-----|
| Step 1: Build fails | Missing .NET Framework targeting pack | Check `Microsoft.NETFramework.ReferenceAssemblies` package installed |
| Step 1: extensibility.dll not found | Path differs on system | Search: `Get-ChildItem -Recurse "C:\Program Files" -Filter "extensibility.dll"` |
| Step 2: regasm fails | Not running as Administrator | Right-click → Run as Administrator |
| Step 2: Registry key error | Wrong path | Check `HKCU` vs `HKLM`, checkProgId spelling |
| Step 3: Add-in not listed | LoadBehavior wrong | Set to 3 (connected + bootload). Check Event Viewer for errors. |
| Step 3: Word crashes on load | OnConnection throws | Wrap all code in try/catch. Check DebugView output. |
| Step 4: Handle is zero | Thread ID wrong | Verify `GetCurrentThreadId()` returns non-zero. Try `0` for thread ID. |
| Step 5: Callback never fires | Hook not on correct thread | Verify WH_KEYBOARD=2 (not 13). Verify thread ID is Word's UI thread. |
| Step 6: 'a' still appears | Wrong VK code or message | Check `lParam` layout. Verify `wParam` is `0x0100` (WM_KEYDOWN). |
| Step 7: COM error | Wrong application object | Verify `(Word.Application)application` cast. Try `_app.Selection.Text` first. |

---

## Cleanup After PoC

After all assumptions are validated:

```powershell
# Unregister
$regasm = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe"
& $regasm /u "poc\bin\Debug\net472\PoC.dll"
Remove-Item "HKCU:\Software\Microsoft\Office\Word\Addins\WordVimPoC.Connect" -ErrorAction SilentlyContinue
```

Then delete the `poc/` directory. The PoC code is throwaway — not part of the project.
