# NetMon Tray

A compact Windows tray utility written in plain C for GCC/MinGW.
It monitors a list of computers in a local network, shows their online/offline state from the tray, can send Wake-on-LAN packets, start an RDP session, and can request remote shutdown or reboot with credentials stored in an INI file.

## Features

- Windows 10+ only
- Native WinAPI tray application, no external GUI frameworks
- Starts directly to the system tray
- Monitors computers from `netmon.ini`
- Per-host monitoring interval
- Right-click or left-click the tray icon to open a menu with all configured computers
- Each computer has a submenu with:
  - a non-clickable IP line
  - `On`
  - `Off`
  - `Reboot`
  - `RDP`
- Availability of `On`, `Off`, `Reboot`, and `RDP` is controlled per host from the INI file
- `On` asks for confirmation and sends a Wake-on-LAN magic packet
- `Off` asks for confirmation, connects to `\\HOST\IPC$` with the configured credentials and runs `shutdown.exe /m \\HOST /s /t 0 /f`
- `RDP` starts `mstsc.exe /v:HOST`
- Every manual command shows a result dialog and a tray balloon: success or error
- Tray balloon notifications for online/offline state changes
- Optional always-visible compact topmost window with all configured computers
- In that window, online computers are black, offline computers are silver, names are sorted by IP, and a divider is drawn whenever the 3rd IPv4 octet changes
- The online list window auto-sizes to its content, uses a yellow sticky-note style, is semi-transparent, has a thin black border, and can be dragged by its body
- Right-click on the online list window also opens the same context menu as the tray icon
- The online list window position is remembered in `netmon.cache.ini`
- Very small project footprint: a single C source file plus config/build files

## Build

Use MinGW-w64 GCC on Windows.

```bat
build_gcc.bat
```

The build script generates `netmon.exe`.

### Expected toolchain

- `gcc`
- standard MinGW-w64 Windows import libraries

Linked system libraries:

- `ws2_32`
- `iphlpapi`
- `mpr`
- `shell32`

## Configuration

The application expects `netmon.ini` in the same folder as the executable.

Example:

```ini
[general]
always_show_online=0
broadcast=255.255.255.255
icmp_timeout_ms=700
tcp_timeout_ms=350

[pc1]
name=Office-PC
address=192.168.1.10
username=.\Administrator
password=YourPasswordHere
interval_sec=5
mac=AA-BB-CC-DD-EE-FF
allow_wake=1
allow_shutdown=1
allow_reboot=1
allow_rdp=1

[pc2]
name=Render-Node
address=192.168.1.20
username=DOMAIN\AdminUser
password=YourPasswordHere
interval_sec=10
allow_wake=0
allow_shutdown=1
allow_reboot=1
allow_rdp=1
; mac is optional but strongly recommended for Wake-on-LAN
```

### INI fields

#### `[general]`

- `always_show_online` — `1` or `0`
- `broadcast` — broadcast address used for WoL, default `255.255.255.255`
- `icmp_timeout_ms` — ICMP timeout in milliseconds
- `tcp_timeout_ms` — TCP fallback timeout in milliseconds

#### Per-computer section

- `name` — display name
- `address` — IPv4 address
- `username` — account used for remote shutdown
- `password` — password for that account
- `interval_sec` — monitoring period in seconds
- `mac` — optional MAC address for Wake-on-LAN, format `AA-BB-CC-DD-EE-FF` or `AA:BB:CC:DD:EE:FF`
- `allow_wake` — `1` or `0`, shows/hides the `On` action
- `allow_shutdown` — `1` or `0`, shows/hides the `Off` action
- `allow_reboot` — `1` or `0`, shows/hides the `Reboot` action
- `allow_rdp` — `1` or `0`, shows/hides the `RDP` action

## How monitoring works

For each host the application:

1. tries an ICMP echo request
2. if ICMP fails, tries a quick TCP connect to ports `445` and `135`

If any check succeeds, the computer is treated as online.

## Wake-on-LAN notes

Wake-on-LAN requires the target MAC address.

- Best option: set `mac=` explicitly in `netmon.ini`
- If `mac=` is missing and the machine is currently online, the application tries to learn the MAC via ARP and stores it in `netmon.cache.ini`
- A machine that has never been online during app runtime may not be wakeable until its MAC is known

## Remote shutdown notes

Remote shutdown depends on Windows permissions and network policy.

Requirements usually include:

- the target account must have rights to shut the machine down remotely
- firewall rules must allow remote administration / RPC
- the app may need to run elevated on the controlling machine
- the target host must be reachable by `shutdown.exe`

When shutdown or reboot fails, the application shows the Windows error code and message to make troubleshooting easier.

## UI behavior

- On startup the app hides in the tray
- Left-click or right-click the tray icon to open the main menu
- Left-click the tray icon to show or hide the server list window
- Right-click on the online list window to open the same action menu
- `Show online list` toggles a compact topmost yellow window that lists every configured computer by name
- In that window, active machines are black and inactive machines are silver
- Selecting `On`, `Off`, or `Reboot` asks for confirmation first
- After `On`, `Off`, `Reboot`, or `RDP`, the app always shows a result message

## Security note

Passwords are stored in plain text in `netmon.ini` because that is part of the requested design.
Restrict access to the application folder accordingly.

## Project structure

```text
netmon_tray/
├─ src/
│  └─ main.c
├─ build_gcc.bat
├─ netmon.ini
├─ netmon.ini.example
└─ README.md
```

## Remarks

This project intentionally stays small and avoids external DLL dependencies beyond standard Windows system libraries.
