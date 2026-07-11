using System;
using System.IO;
using System.Runtime.InteropServices;
using Extensibility;
using Word = Microsoft.Office.Interop.Word;

[ProgId("WordVimPoC.Connect")]
[ComVisible(true)]
[Guid("6E3D0A21-B86A-4D35-9C5E-7F1A8B2D4E60")]
[ClassInterface(ClassInterfaceType.None)]
public class Connect : IDTExtensibility2
{
    private static readonly string LogPath =
        Path.Combine(Path.GetTempPath(), "wordvim-poc.log");

    delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);
    static HookProc _proc = KeyboardProc;
    static IntPtr _hook;
    static bool _insert = false;
    static Word.Application _app;

    [DllImport("user32.dll", SetLastError = true)]
    static extern IntPtr SetWindowsHookEx(int id, HookProc proc, IntPtr mod, uint tid);

    [DllImport("user32.dll")]
    static extern bool UnhookWindowsHookEx(IntPtr hook);

    [DllImport("user32.dll")]
    static extern IntPtr CallNextHookEx(IntPtr hook, int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("kernel32.dll")]
    static extern uint GetCurrentThreadId();

    static void Log(string msg)
    {
        try { File.AppendAllText(LogPath,
            $"[{DateTime.Now:HH:mm:ss.fff}] {msg}{Environment.NewLine}"); }
        catch { }
    }

    static IntPtr KeyboardProc(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= 0 && (lParam.ToInt64() & (1L << 31)) == 0)
        {
            int vk = wParam.ToInt32();

            if (_insert && vk == 0x1B)       // Escape -> Normal
            {
                _insert = false;
                try { _app.ActiveWindow.Caption = "-- NORMAL --"; } catch { }
                Log("Mode: Normal");
                return (IntPtr)1;
            }
            if (!_insert && vk == 0x49)      // 'i' -> Insert
            {
                _insert = true;
                try { _app.ActiveWindow.Caption = "-- INSERT --"; } catch { }
                Log("Mode: Insert");
                return (IntPtr)1;
            }
            if (!_insert) return (IntPtr)1;  // Normal: swallow everything else
        }
        return CallNextHookEx(IntPtr.Zero, nCode, wParam, lParam);
    }

    public void OnConnection(object application, ext_ConnectMode connectMode,
        object addInInst, ref Array custom)
    {
        try
        {
            Log("=== Experiment #6: Mode Display ===");
            _app = (Word.Application)application;
            try { _app.ActiveWindow.Caption = "-- NORMAL --"; } catch { }
            uint tid = GetCurrentThreadId();
            _hook = SetWindowsHookEx(2, _proc, IntPtr.Zero, tid);
            Log(_hook != IntPtr.Zero
                ? $"Hook installed (tid={tid})"
                : "Hook FAILED to install");
        }
        catch (Exception ex) { Log($"ERROR: {ex}"); }
    }

    public void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom)
    {
        if (_hook != IntPtr.Zero)
        {
            UnhookWindowsHookEx(_hook);
            _hook = IntPtr.Zero;
            Log("Hook removed");
        }
    }

    public void OnAddInsUpdate(ref Array custom) { }
    public void OnStartupComplete(ref Array custom) { }
    public void OnBeginShutdown(ref Array custom) { }
}
