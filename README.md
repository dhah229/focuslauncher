# FocusLauncher

A tiny Windows helper that behaves like **“focus-or-launch”**:

- If an app window is already open, **bring it to the front**
- Otherwise, **launch/activate** the app

Works well for Microsoft Store apps and “installed web apps” (Chrome PWAs) that don’t have a normal `.exe` you can point PowerToys at.

## How it works

On each run:

1) If `--title` is provided, it searches visible top-level windows and focuses the first one whose title contains that substring (optionally filtered by `--process`).  
2) If no matching window is found, it tries to activate the app by AUMID (AppID) using Windows’ activation manager.  
3) If activation fails, it falls back to launching via: `explorer.exe shell:AppsFolder\<AUMID>`.

## Usage

```powershell
FocusLauncher.exe <AppID/AUMID> [--title "<substring>"] [--process "<processName>"]
FocusLauncher.exe --list
```

Notes:
- `<AppID/AUMID>` can be either:
  - `Chrome._crx_...` / `Microsoft.Whatever...!App`
  - OR `shell:AppsFolder\Chrome._crx_...` (the tool strips the prefix automatically)
- `--process` expects the **process name without `.exe`** (example: `chrome`, not `chrome.exe`)

### List windows (for debugging)
```powershell
FocusLauncher.exe --list
```

This prints:
- Window title
- Process name  
…so you can pick a good `--title` substring and confirm the right `--process`.

## Finding the AppID (AUMID)

In PowerShell:

### Microsoft Store app example (WhatsApp)
```powershell
Get-StartApps | Where-Object { $_.Name -like "*WhatsApp*" } | Select Name, AppID
```

### Chrome installed app / PWA example (YouTube)
```powershell
Get-StartApps | Where-Object { $_.Name -like "*YouTube*" } | Select Name, AppID
```

If you don’t know the exact name, broaden it:
```powershell
Get-StartApps | Where-Object { $_.Name -like "*chrome*" } | Select Name, AppID
```

The `AppID` value is what you pass as the first argument to `FocusLauncher.exe`.

## Example: YouTube (Chrome installed app)

Test from the repo folder:

```powershell
dotnet run -- "Chrome._crx_agimnkijcamfeangaknmldooml" --title "YouTube" --process "chrome"
```

What this does:
- If a Chrome PWA window with “YouTube” in the title exists, it focuses it
- Otherwise it activates/launches the app by AUMID

## PowerToys setup (Keyboard Manager)

1) PowerToys → **Keyboard Manager** → **Remap a shortcut**
2) Action: **Run Program**
3) Program:
```text
C:\dev\focuslauncher\FocusLauncher\bin\Release\net10.0-windows\win-x64\FocusLauncher.exe
```
4) Arguments:
```text
"Chrome._crx_agimnkijcamfeangaknmldooml" --title "YouTube" --process "chrome"
```

Tip: if focusing doesn’t work, run `FocusLauncher.exe --list` and tweak your `--title` substring to match what Windows shows.

## Build / publish

### Framework-dependent (smaller EXE, requires .NET runtime on the target machine)
Use this if the machine already has the .NET Desktop Runtime installed:

```powershell
dotnet publish -c Release -r win-x64 --self-contained false /p:PublishSingleFile=true
```

### Self-contained (no .NET install required on the target machine)
Use this to run on a machine **without** .NET installed (bigger EXE):

```powershell
dotnet publish -c Release -r win-x64 --self-contained true /p:PublishSingleFile=true
```

Output will be under something like:
```text
.\bin\Release\net10.0-windows\win-x64\publish\
```

## Troubleshooting

- **It launches a new window instead of focusing**  
  Use `--title` + `--process`, and verify the window title/process using `--list`.

- **Focusing fails sometimes**  
  Windows can block foreground focus in some cases; re-trying the hotkey usually works. Also avoid mixing admin/non-admin contexts (an elevated window can be harder to focus from a non-elevated process).
