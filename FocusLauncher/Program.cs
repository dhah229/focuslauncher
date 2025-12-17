using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        if (args.Length == 0)
            return 2;

        if (args.Length == 1 && args[0].Equals("--list", StringComparison.OrdinalIgnoreCase))
        {
            ListWindows();
            return 0;
        }

        // Args:
        //   FocusLauncher <AppID> [--title <substring>] [--process <name>]
        var raw = args[0].Trim().Trim('"');
        var appId = NormalizeAppId(raw);

        string? titleContains = null;
        string? processName = null;

        for (int i = 1; i < args.Length; i++)
        {
            if (args[i].Equals("--title", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length)
            {
                titleContains = args[++i];
                continue;
            }
            if (args[i].Equals("--process", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length)
            {
                processName = args[++i];
                continue;
            }
        }

        // 1) Prefer focusing an existing window (fast, avoids spawning a new instance)
        if (!string.IsNullOrWhiteSpace(titleContains))
        {
            if (TryFocusByTitle(titleContains!, processName))
                return 0;
        }

        // 2) Otherwise try activation (may spawn a new instance for Chrome PWAs)
        if (TryActivate(appId))
            return 0;

        // 3) Fallback shell launch
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "explorer.exe",
                Arguments = $"shell:AppsFolder\\{appId}",
                UseShellExecute = true
            });
            return 0;
        }
        catch
        {
            return 1;
        }
    }

    private static string NormalizeAppId(string input)
    {
        var s = input.Trim().Trim('"');
        const string prefix = "shell:AppsFolder\\";
        if (s.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            s = s.Substring(prefix.Length);
        return s;
    }

    private static bool TryFocusByTitle(string titleSubstring, string? processName)
    {
        bool focused = false;
        var needle = titleSubstring;

        Win32.EnumWindows((hWnd, _) =>
        {
            if (!Win32.IsWindowVisible(hWnd))
                return true;

            var title = Win32.GetWindowTitle(hWnd);
            if (string.IsNullOrWhiteSpace(title))
                return true;

            if (!title.Contains(needle, StringComparison.OrdinalIgnoreCase))
                return true;

            if (!string.IsNullOrWhiteSpace(processName))
            {
                var pid = Win32.GetWindowProcessId(hWnd);
                if (pid == 0) return true;

                try
                {
                    var p = Process.GetProcessById((int)pid);
                    if (!string.Equals(p.ProcessName, processName, StringComparison.OrdinalIgnoreCase))
                        return true;
                }
                catch
                {
                    return true;
                }
            }

            Win32.FocusWindow(hWnd);
            focused = true;
            return false; // stop
        }, IntPtr.Zero);

        return focused;
    }

    private static void ListWindows()
    {
        Console.WriteLine("Visible top-level windows (Title | Process):");
        Console.WriteLine("------------------------------------------------------------");

        Win32.EnumWindows((hWnd, _) =>
        {
            if (!Win32.IsWindowVisible(hWnd))
                return true;

            var title = Win32.GetWindowTitle(hWnd);
            if (string.IsNullOrWhiteSpace(title))
                return true;

            var pid = Win32.GetWindowProcessId(hWnd);
            var proc = "(unknown)";
            if (pid != 0)
            {
                try { proc = Process.GetProcessById((int)pid).ProcessName; } catch { }
            }

            Console.WriteLine($"{title} | {proc}");
            return true;
        }, IntPtr.Zero);
    }

    private static bool TryActivate(string appId)
    {
        try
        {
            var aam = (IApplicationActivationManager)new ApplicationActivationManager();
            uint pid;

            int hr = aam.ActivateApplication(appId, null, ActivateOptions.None, out pid);
            if (hr >= 0) return true;

            if (!appId.Contains("!"))
            {
                hr = aam.ActivateApplication(appId + "!App", null, ActivateOptions.None, out pid);
                if (hr >= 0) return true;
            }
        }
        catch { }

        return false;
    }

    [Flags]
    private enum ActivateOptions : uint
    {
        None = 0x00000000,
        DesignMode = 0x00000001,
        NoErrorUI = 0x00000002,
        NoSplashScreen = 0x00000004
    }

    [ComImport]
    [Guid("2e941141-7f97-4756-ba1d-9decde894a3d")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IApplicationActivationManager
    {
        int ActivateApplication(
            [MarshalAs(UnmanagedType.LPWStr)] string appUserModelId,
            [MarshalAs(UnmanagedType.LPWStr)] string? arguments,
            ActivateOptions options,
            out uint processId);

        int ActivateForFile(
            [MarshalAs(UnmanagedType.LPWStr)] string appUserModelId,
            IntPtr itemArray,
            [MarshalAs(UnmanagedType.LPWStr)] string? verb,
            out uint processId);

        int ActivateForProtocol(
            [MarshalAs(UnmanagedType.LPWStr)] string appUserModelId,
            IntPtr itemArray,
            out uint processId);
    }

    [ComImport]
    [Guid("45BA127D-10A8-46EA-8AB7-56EA9078943C")]
    private class ApplicationActivationManager { }

    private static class Win32
    {
        public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
        [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
        [DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr hWnd);
        [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
        [DllImport("user32.dll")] public static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowTextLength(IntPtr hWnd);

        private const int SW_RESTORE = 9;

        public static string GetWindowTitle(IntPtr hWnd)
        {
            int len = GetWindowTextLength(hWnd);
            if (len <= 0) return "";
            var sb = new StringBuilder(len + 1);
            _ = GetWindowText(hWnd, sb, sb.Capacity);
            return sb.ToString();
        }

        public static uint GetWindowProcessId(IntPtr hWnd)
        {
            _ = GetWindowThreadProcessId(hWnd, out var pid);
            return pid;
        }

        public static void FocusWindow(IntPtr hWnd)
        {
            var fg = GetForegroundWindow();
            uint fgThread = fg != IntPtr.Zero ? GetWindowThreadProcessId(fg, out _) : 0;
            uint targetThread = GetWindowThreadProcessId(hWnd, out _);

            bool attached = false;
            try
            {
                if (fgThread != 0 && fgThread != targetThread)
                    attached = AttachThreadInput(fgThread, targetThread, true);

                ShowWindowAsync(hWnd, SW_RESTORE);
                BringWindowToTop(hWnd);
                SetForegroundWindow(hWnd);
            }
            finally
            {
                if (attached)
                    AttachThreadInput(fgThread, targetThread, false);
            }
        }
    }
}
