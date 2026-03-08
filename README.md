<p align="center">
  <h1 align="center">QuickLoopUp</h1>
  <p align="center">
    <strong>A native Windows dictionary popup — the macOS "Look Up" experience, reimagined for Windows.</strong>
  </p>
  <p align="center">
    <video src="assets/QuickLoopUpDemo.mp4" controls="controls" muted="muted" playsinline="playsinline" autoplay="autoplay" loop="loop" style="max-width: 600px; border-radius: 8px; box-shadow: 0 4px 12px rgba(0,0,0,0.5);"></video>
  </p>
  <p align="center">
    <a href="#features">Features</a> •
    <a href="#how-it-works">How It Works</a> •
    <a href="#building-from-source">Build</a> •
    <a href="#usage">Usage</a> •
    <a href="#license">License</a>
  </p>
</p>

---

Long-click any word in any Windows application — browsers, editors, PDFs — and get an instant dictionary popup with definitions, phonetics, synonyms & more. Phrases and unknown words fall back to Google Search seamlessly.

Built entirely in **C++** with raw **Win32 APIs**, **WebView2**, and **DWM** — no frameworks, no Electron, no overhead.

## Features

| Feature | Description |
|---------|-------------|
| **800ms Long-Click Trigger** | Hold the mouse on any word for ~0.8 seconds to look it up — fast and unobtrusive |
| **Universal Text Extraction** | Uses **UI Automation** (`IUIAutomationTextPattern`) to extract text from Chromium, native, and Electron apps |
| **Smart Clipboard Fallback** | Falls back to `Ctrl+C` simulation for legacy apps, with element-type safety checks to avoid triggering buttons |
| **Drag-Aware** | Tracks dragging state so text selections are never interrupted by false triggers |
| **Persistent HTTP** | Single `WinHTTP` session with aggressive timeouts (3s connect, 5s receive), reused across all lookups |
| **~15 MB Idle RAM** | Chromium forced into single-process mode; V8 memory reclaimed on popup dismiss |
| **DWM Rounded Corners** | Native OS-level rounded corners via `DwmSetWindowAttribute` (Windows 11+) |
| **Subtle Box Shadow** | DWM frame extension provides a clean elevation shadow separating popup from desktop |
| **Safe JSON Parsing** | Recursive-descent string extraction — handles escapes, nesting, and edge cases |
| **Multi-Monitor DPI** | Correct scaling and positioning on any monitor at any DPI |

## How It Works

```
 ┌─────────────────┐     ┌──────────────────┐     ┌───────────────────┐
 │  WH_MOUSE_LL    │────▶│  UI Automation   │────▶│  Dictionary API   │
 │  Global Hook    │     │  Text Extraction  │     │  (dictionaryapi)  │
 │  (800ms hold)   │     │  + Clipboard      │     │  + Google Search  │
 └─────────────────┘     └──────────────────┘     └───────────────────┘
                                                           │
                                                           ▼
                                                  ┌───────────────────┐
                                                  │  WebView2 Popup   │
                                                  │  (DWM corners +   │
                                                  │   shadow)         │
                                                  └───────────────────┘
```

1. **Global Hook** — `WH_MOUSE_LL` + `WH_KEYBOARD_LL` detect 800ms long-press and dismiss events across all apps.
2. **Text Extraction** — Background thread uses `ElementFromPoint` → `TextPattern::GetSelection` or `RangeFromPoint` → `ExpandToEnclosingUnit(TextUnit_Word)`. Falls back to clipboard with element-type safety.
3. **Dictionary Fetch** — Persistent `WinHTTP` session queries `api.dictionaryapi.dev`. On 404 or failure, redirects to Google Search.
4. **Rendering** — Stripped-down WebView2 renders dynamically-generated HTML/CSS. V8 pre-warmed on startup; DOM memory reclaimed on dismiss.
5. **DWM** — `DwmSetWindowAttribute(DWMWCP_ROUND)` for rounded corners, `DwmExtendFrameIntoClientArea` for shadow.

## Installation

### Quick Install

Run the installer:
**`QuickLoopUp_Setup.exe`**

The installer will:
1. Ask you to choose an install directory (default: `C:\Program Files\QuickLoopUp`)
2. Copy `QuickLoopUp.exe` (a single self-contained executable)
3. Optionally create a **Startup folder shortcut** so it launches automatically on boot
4. Optionally create a **Desktop shortcut**

### Uninstall

Run the **Uninstall QuickLoopUp** shortcut from the Start Menu, or uninstall it directly from the Windows **Apps & Features** settings menu.

## Usage

1. **Launch** — Run `QuickLoopUp.exe` or restart your PC (it auto-starts). No tray icon — it runs silently.
2. **Look up a word** — In any app, hold the left mouse button on a word for ~0.8s.
3. **Look up a phrase** — Highlight a phrase and long-click, or simply **drag to select and hold** for ~0.8s at the end of your drag. Multi-word selections open Google Search.
4. **Dismiss** — Click outside the popup, press any key, or press `Esc`.
5. **Exit** — End the process via Task Manager.

## Performance

| Metric | Value |
|--------|-------|
| Trigger delay | 800ms long-press |
| API timeout | 3s connect / 5s receive |
| Clipboard fallback | 30–40ms sleep |
| Idle RAM | < 15 MB |
| HTTP session | Persistent (reused) |
| Binary size | ~1 MB (statically linked, DLL embedded) |

## Building from Source

<details>
<summary>Click to expand build instructions</summary>

### Prerequisites

- **g++** — MinGW-w64 toolchain
- **windres** — Resource compiler (included with MinGW-w64)
- **WebView2 headers** — [Microsoft.Web.WebView2 NuGet package](https://www.nuget.org/packages/Microsoft.Web.WebView2)
- **WebView2Loader.dll** — Must be present to compile into the resources.
- **Inno Setup 6** — To compile the installer `installer.iss`

### Quick Build

```powershell
.\build.ps1
```

### Manual Build

```bash
# 1. Compile resources (embeds WebView2Loader.dll)
windres resources.rc -o resources.o

# 2. Compile application
g++ main.cpp resources.o -o QuickLoopUp.exe -O3 -mwindows -static \
    -D_WIN32_WINNT=0x0A00 -DNTDDI_VERSION=0x0A000000 \
    -I<path-to-webview2-include> \
    -luiautomationcore -lshlwapi -luser32 -lgdi32 \
    -lole32 -loleaut32 -luuid -lwinhttp -lshcore -ldwmapi

# 3. Create Installer (Optional)
iscc installer.iss
```

</details>

## Tech Stack

- **Language** — C++17
- **UI** — Win32 API + WebView2 (Microsoft Edge)
- **Text Extraction** — UI Automation COM
- **Networking** — WinHTTP
- **Window Effects** — DWM API (rounded corners + shadow)
- **Compiler** — MinGW-w64 (g++)

## License

[MIT](LICENSE) © Charan
