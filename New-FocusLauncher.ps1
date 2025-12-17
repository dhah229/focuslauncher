param(
  [Parameter(Mandatory = $true)]
  [string]$AppName,  # partial match is fine, e.g. "Loop"

  [int]$Pick = 1,    # if multiple matches, pick 1st by default (1-based)

  [string]$OutDir = "$env:USERPROFILE\Documents\PowertoysLaunchers",
  [string]$BaseName
)

function Sanitize-FileName([string]$s) {
  $invalid = [IO.Path]::GetInvalidFileNameChars()
  foreach ($c in $invalid) { $s = $s.Replace($c, "_") }
  return ($s -replace '\s+', '-').Trim('-')
}

# Find matching Start menu apps (wildcard)
$matches = @(Get-StartApps | Where-Object { $_.Name -like "*$AppName*" } | Sort-Object Name)

if ($matches.Length -eq 0) {
  Write-Error "No Start Menu app matches '$AppName'. Try: Get-StartApps | Select Name,AppID"
  exit 1
}

# Show matches (helps when there are multiple)
Write-Host "Matches:" -ForegroundColor Cyan
for ($i = 0; $i -lt $matches.Length; $i++) {
  $n = $i + 1
  Write-Host ("  {0}. {1}    [{2}]" -f $n, $matches[$i].Name, $matches[$i].AppID)
}

if ($Pick -lt 1 -or $Pick -gt $matches.Length) {
  Write-Error "Pick must be between 1 and $($matches.Length)."
  exit 1
}

$chosen = $matches[$Pick - 1]
$Aumid = $chosen.AppID
$DisplayName = $chosen.Name

if (-not $BaseName) { $BaseName = Sanitize-FileName $DisplayName }

New-Item -ItemType Directory -Path $OutDir -Force | Out-Null

$ps1Path = Join-Path $OutDir "$BaseName-focus-or-launch.ps1"
$vbsPath = Join-Path $OutDir "$BaseName-focus-or-launch.vbs"

# --- PS1 template (DLL-cached for speed) ---
$ps1Template = @'
$Aumid = "{{AUMID}}"
$shellTarget = "shell:AppsFolder\$Aumid"

# Cache compiled helper DLL beside this script (huge speedup on repeated runs)
$dll = Join-Path $PSScriptRoot "AumidFocus.dll"

$code = @"
using System;
using System.Runtime.InteropServices;

public static class AumidWindowFinder {
  public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

  [DllImport("user32.dll")] static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
  [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr hWnd);
  [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr hWnd);
  [DllImport("user32.dll")] static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);

  [DllImport("shell32.dll")]
  static extern int SHGetPropertyStoreForWindow(IntPtr hwnd, ref Guid iid, out IPropertyStore propertyStore);

  [StructLayout(LayoutKind.Sequential)]
  public struct PROPERTYKEY { public Guid fmtid; public uint pid; }

  static PROPERTYKEY PKEY_AppUserModel_ID = new PROPERTYKEY {
    fmtid = new Guid("9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3"), pid = 5
  };

  [StructLayout(LayoutKind.Explicit)]
  public struct PROPVARIANT { [FieldOffset(0)] public ushort vt; [FieldOffset(8)] public IntPtr p; }

  [DllImport("ole32.dll")] static extern int PropVariantClear(ref PROPVARIANT pvar);

  [ComImport, Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
  public interface IPropertyStore {
    uint GetCount(out uint cProps);
    uint GetAt(uint iProp, out PROPERTYKEY pkey);
    uint GetValue(ref PROPERTYKEY key, out PROPVARIANT pv);
    uint SetValue(ref PROPERTYKEY key, ref PROPVARIANT pv);
    uint Commit();
  }

  static string GetAumid(IntPtr hWnd) {
    Guid iid = new Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99");
    IPropertyStore store;
    int hr = SHGetPropertyStoreForWindow(hWnd, ref iid, out store);
    if (hr != 0 || store == null) return null;

    PROPVARIANT pv;
    store.GetValue(ref PKEY_AppUserModel_ID, out pv);
    try {
      const ushort VT_LPWSTR = 31;
      if (pv.vt == VT_LPWSTR && pv.p != IntPtr.Zero) return Marshal.PtrToStringUni(pv.p);
      return null;
    } finally { PropVariantClear(ref pv); }
  }

  static string _target;
  static IntPtr _found;

  public static bool FocusByAumid(string aumid) {
    _target = aumid;
    _found = IntPtr.Zero;

    EnumWindows((hWnd, lParam) => {
      if (!IsWindowVisible(hWnd)) return true;
      try {
        if (GetAumid(hWnd) == _target) { _found = hWnd; return false; }
      } catch {}
      return true;
    }, IntPtr.Zero);

    if (_found != IntPtr.Zero) {
      ShowWindowAsync(_found, 9);
      SetForegroundWindow(_found);
      return true;
    }
    return false;
  }
}
"@

if (-not (Test-Path $dll)) {
  Add-Type -TypeDefinition $code -Language CSharp -OutputAssembly $dll -OutputType Library
}
Add-Type -Path $dll

if (-not [AumidWindowFinder]::FocusByAumid($Aumid)) {
  Start-Process explorer.exe $shellTarget
}
'@

$ps1 = $ps1Template.Replace("{{AUMID}}", $Aumid)
Set-Content -Path $ps1Path -Value $ps1 -Encoding UTF8

# --- VBS wrapper (hidden) ---
# Write as ASCII to avoid UTF-8 BOM issues in Windows Script Host
$vbs = 'CreateObject("WScript.Shell").Run """C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"" -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File ""' + $ps1Path + '""", 0, False'
Set-Content -Path $vbsPath -Value $vbs -Encoding ASCII

Write-Host ""
Write-Host "Created:" -ForegroundColor Green
Write-Host "  App:   $DisplayName"
Write-Host "  AUMID: $Aumid"
Write-Host "  PS1:   $ps1Path"
Write-Host "  VBS:   $vbsPath"
Write-Host ""
Write-Host "PowerToys -> Run program:" -ForegroundColor Yellow
Write-Host "  App:  C:\Windows\System32\wscript.exe"
Write-Host "  Args: ""$vbsPath"""
