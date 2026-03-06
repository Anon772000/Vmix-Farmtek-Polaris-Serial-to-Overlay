# FarmTek Timer Router

Low-latency FarmTek-to-vMix timer router for Windows.

It reads FarmTek serial output, detects which timer rule should handle the frame, updates the matching vMix title over the vMix TCP API, and can trigger overlays automatically.

The current app supports:

- multiple timer rules
- routing by decimal places
- quote-based running/stopped detection
- optional flashing when a stopped time is held
- raw serial monitoring in the web UI
- a standalone Windows EXE build

## Requirements

- Windows
- vMix running on the same machine or reachable over the network
- FarmTek timer/controller connected by serial
- .NET 8 SDK if you want to build from source

## Quick Start

### Run the EXE

1. Start [FarmtekTimerRouter.exe](./FarmtekTimerRouter.exe).
2. Open the local web UI if it does not open automatically:

```text
http://localhost:8090/
```

3. In the web UI:
   - choose the COM port
   - confirm baud rate
   - click `Refresh vMix`
   - set one rule for each timer format
   - turn on `Quote means running` if your FarmTek output uses a trailing `"` while the timer is live
- turn on `Flash when stopped` if you want the held result to blink
   - `Flash interval` is how long the title stays visible between flashes
   - `Flash duration` is how long the title stays hidden during each flash
4. Click `Save settings`.

### Run from source

```powershell
.\timer.ps1
```

Useful options:

```powershell
.\timer.ps1 -ValidateOnly
.\timer.ps1 -NoBrowser
.\timer.ps1 -UiPort 8098
```

## How The Quote Mode Works

If your controller sends:

- `8.32"` while running
- `8.32` when stopped

then enable `Quote means running` on that rule.

In that mode:

- the trailing `"` is used as the running/stopped signal
- the `"` is removed before sending the number to vMix
- auto overlay can trigger when the timer starts running
- the overlay re-arms itself when the quote disappears
- optional stopped flashing uses vMix alpha `255 -> 0 -> 255`
- the old threshold-based rearm values are not needed for that rule

## Files

- [Program.cs](./Program.cs) - backend, serial reader, vMix TCP client, API
- [web/index.html](./web/index.html) - UI markup
- [web/app.js](./web/app.js) - UI logic
- [web/styles.css](./web/styles.css) - UI styling
- [timer.ps1](./timer.ps1) - source launcher
- [publish-exe.ps1](./publish-exe.ps1) - self-contained EXE publish script
- [timer.settings.example.json](./timer.settings.example.json) - example config for GitHub

## Building The EXE

```powershell
.\publish-exe.ps1
```

That produces:

- `.publish\win-x64\FarmtekTimerRouter.exe`
- a copied root EXE at `.\FarmtekTimerRouter.exe`

For GitHub, it is usually better to upload the EXE to a GitHub Release instead of committing the binary into the source repo.


## First-Time Config On Another Machine

After cloning:

1. Run the EXE or `.\timer.ps1`.
2. The app will create a fresh `timer.settings.json` if one does not exist.
3. Configure COM port, vMix target inputs, and rule options in the web UI.

## Troubleshooting

- If the UI says `Not found`, make sure you opened the correct port.
- If vMix is not updating, click `Refresh vMix` and confirm the target rule resolves to the correct input.
- If serial does not connect, check whether another app is already using the COM port.
- Use the `Serial Monitor` panel to inspect the raw FarmTek output when changing controller modes.
