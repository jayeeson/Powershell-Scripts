# Startup Program Closed Window

## The Problem

Some programs are difficult to launch in the background on boot - you want them running but without any windows open cluttering your desktop.
Run this script at computer startup using C omputer Management, or an AutoHotKey script, etc.

## Configuration

Edit the `$programs` array at the top of the script to specify which programs to launch:

```powershell
$programs = @(
    @{
        ExePath = "C:\Path\To\Program.exe"
        ProcessName = "ProcessName"
    }
)
```

You can also adjust the timing parameters:
- `$defaultWaitTime`: Delay in milliseconds between window close attempts (default: 300ms)
- `$defaultMaxAttempts`: Maximum number of attempts to close the window (default: 200)
- `$batchSize`: Number of programs to process concurrently (default: 3)

## Finding Process Names

To find the process name for a running application, run the program and keep the window open, then run this in powershell:

```powershell
Get-Process | Where-Object { $_.MainWindowTitle -ne "" }
```
