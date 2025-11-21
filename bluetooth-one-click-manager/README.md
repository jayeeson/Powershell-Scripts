# Bluetooth Tools

PowerShell helpers that make it easy to manage Bluetooth peripherals on Windows 11 using the official Bluetooth Command Line Tools. The main entry point is `bt.ps1`, a friendly wrapper around `btcom.exe`, `btpair.exe`, and `btdiscovery.exe`.

## Requirements
- Windows 11 with Bluetooth enabled
- [Bluetooth Command Line Tools 1.2.0.56](https://bluetoothinstaller.com/bluetooth-command-line-tools/BluetoothCLTools-1.2.0.56.exe) extracted somewhere on your `PATH`
- PowerShell 7+ (pwsh.exe) with execution policy that allows running local scripts

## Quick Start
1. Copy this repository to your machine.
2. Launch `pwsh` in the folder and run `./bt.ps1 -AddDevice` to scan and save a device into `btconfig.json`.
3. Use the shortcuts the script generates, or the .bat files, or call the script directly for day-to-day operations.

## Commands
- `./bt.ps1 -List` — show devices from `btconfig.json` and current connection status.
- `./bt.ps1 -Connect "Friendly"` — connect by friendly name, partial name, or MAC address.
- `./bt.ps1 -Disconnect "Friendly"` — force a clean disconnect.
- `./bt.ps1 -Forget "Friendly"` — remove the device pairing (manual-connect devices).
- `./bt.ps1 -Reset "Friendly"` — forget, rescan, and reconnect devices in the auto-connect group.
- `./bt.ps1 -GenerateShortcuts` — rebuild `.bat` launchers and optional desktop `.lnk` files.
- `.bt.ps1 -Edit` — modify the configuration file manually (advanced).

## Custom Desktop Shortcut Icons
- Place custom `.ico` files alongside the scripts (or reference absolute paths).
- Add icon names under the `icons` property in `btconfig.json` (for example `icons.reset`, `icons.add`).
- Shortcuts fall back to system icons when no custom value is set.

## Example `btconfig.json`
```json
{
	"auto_connect_allowed": [
		{
			"friendly_name": "Xbox Wireless Controller BLACK",
			"name": "Xbox Wireless Controller",
			"mac": "AA:BB:CC:00:11:22",
			"service_classes": [],
			"icons": {
				"reset": "bluetooth-xbox-black.ico"
			},
			"shortcuts": ["reset"]
		},
		{
			"friendly_name": "Xbox Wireless Controller WHITE",
			"name": "Xbox Wireless Controller",
			"mac": "11:22:33:44:55:66",
			"service_classes": [],
			"icons": {"reset": "bluetooth-xbox-white.ico"},
			"shortcuts": ["reset"]
		}
	],
	"manual_connect": [
		{
			"friendly_name": "My Pixel Buds Pro",
			"name": "My Pixel Buds Pro",
			"mac": "77:88:99:AA:BB:CC",
			"service_classes": [
				"110B",
				"111e"
			],
			"icons": {
				"add": "bluetooth-buds-add.ico",
				"remove": "bluetooth-buds-remove.ico"
			},
			"shortcuts": ["connect", "forget"]
		},
		{
			"friendly_name": "My Pixel 9",
			"name": "My Pixel 9",
			"mac": "DD:EE:FF:00:11:22",
			"service_classes": [],
			"icons": {},
			"shortcuts": ["connect", "forget"]
		}
	]
}
```

