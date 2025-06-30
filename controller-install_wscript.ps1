$path = "$env:USERPROFILE\Documents\genesysController"
$vbsScript = @'
Dim arg
arg = WScript.Arguments(0)
Set shell = CreateObject("WScript.Shell")
scriptPath = shell.ExpandEnvironmentStrings("%USERPROFILE%\Documents\genesysController\genesysController.ps1")
shell.Run "powershell.exe -WindowStyle Hidden -NoProfile -File """ & scriptPath & """ -Tag genesysControllerWorker -Context """ & arg & """", 0, False
'@
$psScript = @'
# fixed task name. now with username
param(
    [string]$Context = 'Installation'
)
$MinutesInterval = 1;
$TaskName = "GenesysController - $env:USERNAME";
$path = "$env:USERPROFILE\Documents\genesysController";
function HandleStartupShortcut {
    Write-Host 'Check shortcut in Startup...';
    $StartupShortcutPath = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup\hiddenGenesysControllerLauncher.lnk";
    if (Test-Path $StartupShortcutPath) {
        return;
    }
    $ShortcutPath = "$path\hiddenGenesysControllerLauncher.lnk"
    if (-not (Test-Path $ShortcutPath)) {
        $vbsPath = "$path\hiddenGenesysControllerLauncher.vbs"
        $WshShell = New-Object -ComObject WScript.Shell
        $shortcut = $WshShell.CreateShortcut($ShortcutPath)
        $shortcut.TargetPath = "wscript.exe"
        $shortcut.Arguments = "`"$vbsPath`" `"Startup`""
        $shortcut.WindowStyle = 0;
        $shortcut.Description = "hiddenGenesysControllerLauncher.vbs"
        $shortcut.WorkingDirectory = $path
        $shortcut.Save()
    }
    Copy-Item -Path $ShortcutPath -Destination $StartupShortcutPath
    Write-Host "Shortcut placed in Startup"
}
function HandleTask {    
    Write-Host 'Check task...';
    $Task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue;
    $TaskInfo = $Task | Get-ScheduledTaskInfo;
    if ( (-not $Task) -or ($Task.State -eq 'Disabled') -or ($Context -eq 'Startup')) {
        if ($Task) {
            Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false;
        }
        $Principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME;
        $Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -MultipleInstances Parallel;
        $Action = New-ScheduledTaskAction -Execute "wscript.exe" -Argument "`"$env:USERPROFILE\Documents\genesysController\hiddenGenesysControllerLauncher.vbs`" `"Task`"";
        $TriggerTime = (Get-Date).addSeconds(5);
        $Trigger = New-ScheduledTaskTrigger -Once -At $TriggerTime -RepetitionInterval (New-TimeSpan -Minutes $MinutesInterval);
        Write-Host "Register task: once at $TriggerTime and every $MinutesInterval minute..."
        Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger -Principal $Principal -Settings $Settings;
        Write-Host "Task $TaskName registered."
    }
}

if ($Context -ne 'Startup') {
    HandleStartupShortcut
}

if ($Context -ne 'Task') {
    HandleTask
}

if ($Context -eq 'Installation') {
    pause;
    exit;
}

$SameScriptProcesses = Get-CimInstance Win32_Process | Where-Object { $_.CommandLine -like "*genesysControllerWorker*" };
if ($SameScriptProcesses.length) { # $SameScriptProcess.length exist if >= 2
    exit;    
}
$iteration = 1;
$ctpPath = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Genesys Telecommunications Laboratories\Workspace Desktop Edition\Call_To_Phone.appref-ms";
while ($true) {
    $processes = Get-Process -Name "InteractionWorkspace" -ErrorAction SilentlyContinue;
    $isCTPrunning = $false;
    foreach ($process in $processes) {
        if ($process.path -like "*files*") {
            Stop-Process -Id $process.id -Force; 
        }
        if ($process.path -like "*apps*") {
            $isCTPrunning = $true;
        }
    }
    if (-not $isCTPrunning) {
        if (Test-Path $ctpPath) {
            Start-Process -FilePath $ctpPath;
        }
    }
    if ( ($iteration % 12) -eq 0) { # script check task and startup state every ~1 minute
       HandleStartupShortcut
       HandleTask

    }    
    $iteration++;
    Start-Sleep -Seconds 5;
}
'@

New-Item -Path $path -ItemType Directory -Force | Out-Null
Set-Content -Path "$path\hiddenGenesysControllerLauncher.vbs" -Value $vbsScript
Set-Content -Path "$path\genesysController.ps1" -Value $psScript

powershell.exe -File "$path\genesysController.ps1"
