# FocusLauncher

Generate **“focus-or-launch”** shortcuts for Windows apps: press a hotkey and it will **bring the already-open app window to the front**; if it’s not running, it will **launch it**.

This was built to solve a common PowerToys Keyboard Manager annoyance:

- You can bind a hotkey to “Run program”, but…
- Microsoft Store apps and Chrome “installed web apps” (PWA-style) often **don’t have a normal `.exe` path**
- And when they *do* launch, they often **open a new window every time** (especially web apps like YouTube) instead of focusing the existing one

FocusLauncher generates the files you need so PowerToys hotkeys behave like a “real app shortcut”.

---

## Why this exists (the original problem)

The first goal was to create a PowerToys shortcut to launch a Store app (e.g., WhatsApp) without hunting for an executable path.

That worked using the app’s **AppID/AUMID** (e.g., `shell:AppsFolder\...!App`), but PowerToys was unreliable with direct `shell:` launching and `.lnk` shortcuts.

Then the “tricky one” showed up: **Chrome-installed YouTube**. Launching it always spawned a **new** window because it’s basically a Chrome app window, not a traditional single-instance desktop executable.

The solution was:

1. Identify apps by **AppUserModelID (AUMID)** (the same thing you get from `Get-StartApps`)
2. Find an already-open window with that AUMID and **focus it**
3. If no window exists, **launch** it using `explorer.exe shell:AppsFolder\<AUMID>`
4. Wrap PowerShell in a tiny **VBScript** so PowerToys runs it reliably **without a flashing terminal**
5. Cache the compiled helper code into a DLL so focusing is **fast** after the first run

---

## What it generates

For each app, FocusLauncher generates two files:

- `*-focus-or-launch.ps1`  
  Focuses the app window if it exists; otherwise launches it. Uses a cached `AumidFocus.dll` next to the script for speed.

- `*-focus-or-launch.vbs`  
  Launches the PS1 in a hidden window (no console flash). This is the file you point PowerToys at.

> **Note:** The first time you run a generated launcher, it will create `AumidFocus.dll` beside the PS1. After that, it’s much faster.

---

## Requirements

- Windows 10/11
- PowerShell 5.1+ (Windows PowerShell works great)
- Apps must appear in `Get-StartApps` (most Store apps do; Chrome-installed apps usually do too)

---

## Quick start

### 1) Download / clone
Clone the repo or download `New-FocusLauncher.ps1` into a folder, e.g.:

`C:\Users\<you>\Documents\FocusLauncher\`

### 2) Generate a launcher

From PowerShell:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\New-FocusLauncher.ps1 -AppName "WhatsApp"

