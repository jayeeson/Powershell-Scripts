# Configuration: Define your target programs here
$programs = @(
    @{
        ExePath = "C:\Program Files\Corsair\Corsair iCUE5 Software\iCUE.exe"
        ProcessName = "iCUE"
    }
    @{
        ExePath = "C:\Program Files (x86)\Pushbullet\pushbullet.exe"
        ProcessName = "pushbullet_client"
    }
)


$defaultWaitTime = 300 # ms
$defaultMaxAttempts = 200
$batchSize = 3


#######################################################



$jobs = @()

for ($i = 0; $i -lt $programs.Count; $i += $batchSize) {
    $batch = $programs[$i..[Math]::Min($i + $batchSize - 1, $programs.Count - 1)]
    
    # Start jobs for current batch
    foreach ($program in $batch) {
        Write-Host $program -ForegroundColor DarkMagenta
        $jobs += Start-Job -ScriptBlock {
            param($ExePath, $ProcessName, $WaitTime, $MaxAttempts)
            
            Function Start-And-Close-Application {
                param (
                    [string]$ExePath,
                    [string]$ProcessName,
                    [int]$WaitTime = $defaultWaitTime,
                    [int]$MaxAttempts = $defaultMaxAttempts,
                    [string]$WindowStyle = "Minimized"
                )

                $appAlreadyOpen = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue 
                if ($appAlreadyOpen) {
                    Write-Host "Process already running: $ProcessName" -ForegroundColor DarkYellow
                    Return
                }

                Start-Process -FilePath $ExePath -WindowStyle Minimized

                For ($i = 0; $i -lt $MaxAttempts; $i++) {
                    Start-Sleep -Milliseconds $WaitTime
                    $appProcess = Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -ne "" -and $_.ProcessName -eq $ProcessName}
                    if ($appProcess -and $appProcess.CloseMainWindow()) {
                        Write-Host "Successfully closed $ProcessName" -ForegroundColor Green
                        Return
                    }
                }
                Write-Host "Could not close process: $ProcessName" -ForegroundColor Red
            }
            
            Start-And-Close-Application -ExePath $ExePath -ProcessName $ProcessName `
                -WaitTime $WaitTime -MaxAttempts $MaxAttempts -WindowStyle $WindowStyle
                
        } -ArgumentList $program.ExePath, $program.ProcessName, `
            $(if ($program.WaitTime) { $program.WaitTime } else { $defaultWaitTime }), `
            $(if ($program.MaxAttempts) { $program.MaxAttempts } else { $defaultMaxAttempts }), `
            $(if ($program.WindowStyle) { $program.WindowStyle } else { "Minimized" })
    }
    
    $jobs | Wait-Job | Out-Null
    $jobs | Receive-Job
    $jobs | Remove-Job
    $jobs = @()
}

exit 0