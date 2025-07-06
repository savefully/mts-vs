$task = Get-ScheduledTask -TaskName 'GenesysController' -ErrorAction SilentlyContinue
$taskFixed = Get-ScheduledTask -TaskName "GenesysController - $env:USERNAME" -ErrorAction SilentlyContinue
if ($task) {
	Unregister-ScheduledTask -TaskName 'GenesysController' -Confirm:$false
}
if ($taskFixed) {
	Unregister-ScheduledTask -TaskName "GenesysController - $env:USERNAME" -Confirm:$false
}
Remove-Item "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup\hiddenGenesysControllerLauncher.lnk" -ErrorAction SilentlyContinue
$workerProcess = Get-CimInstance Win32_Process | Where-Object { $_.CommandLine -like "*genesysControllerWorker*" }
if ($workerProcess) {
	Stop-Process -Id $workerProcess.ProcessId -Force
}

$task = Get-ScheduledTask -TaskName 'GenesysController' -ErrorAction SilentlyContinue
Write-Host $task
$taskFixed = Get-ScheduledTask -TaskName "GenesysController - $env:USERNAME" -ErrorAction SilentlyContinue
Write-Host $taskFixed
$shortcut = Get-Item "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup\hiddenGenesysControllerLauncher.lnk" -ErrorAction SilentlyContinue
Write-Host $shortcut
$workerProcess = Get-CimInstance Win32_Process | Where-Object { $_.CommandLine -like "*genesysControllerWorker*" } 
Write-Host $workerProcess

if (-not($task -or $taskFixed -or $shortcut -or $workerProcess)) {
	Write-Host '=== GenesysController successfully disabled ==='
}
