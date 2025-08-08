[console]::InputEncoding = [System.Text.Encoding]::UTF8
[console]::OutputEncoding = [System.Text.Encoding]::UTF8

$downloadDomain = $null
$iconFile = 'app_yellow.ico';
$archivePath = 'C:\Genesys SIP Phone.zip'
$archivePathCC = 'C:\Genesys_SIP_Phone.zip'
$gspPath = 'C:\Users\Public\Downloads\Genesys SIP Phone'
$shortcutPath = 'C:\Users\Public\Desktop\Genesys SIP Phone.lnk'
$machineKeysPath = 'C:\ProgramData\Microsoft\Crypto\RSA\MachineKeys'
$NetFx3Path = "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.5"
$kvpncguiPath = 'C:\Program Files (x86)\Kerio\VPN Client\kvpncgui.exe'
$cryptoKeysPath = 'C:\ProgramData\Application Data\Microsoft\Crypto\Keys'
$citrixSelfServicePath = 'C:\Program Files (x86)\Citrix\ICA Client\SelfServicePlugin\SelfService.exe'
$customRunBatCode = @"
@echo off
if exist "GenesysSIPPhone.exe" (
    start "Genesys SIP Phone" "GenesysSIPPhone.exe" -config genesys
) else (
    start "Genesys SIP Phone" "$gspPath\GenesysSIPPhone.exe" -config genesys
)
exit   
"@
$separator = '----------------------------------------------------------------------------------------------------'
$menuString = @"
$separator
  0 - Exit
0.1 - Precheck (NetFx3 state; Kerio and Citrix versions, paths)

Genesys SIP Phone installation:
  1 - Basic installation: 1.1...1.5
1.1 - Expand archive
1.2 - Remove archive
1.3 - Create shortcuts
1.4 - Set firewall rules for GenesysSIPPhone.exe
1.5 - Set custom RUN.bat (auto close and absolute path)
1.6 - Set 6-sign number
ComponentActivator fixes:
1.7 - $machineKeysPath 
1.8 - $cryptoKeysPath

.NET 3.5 installation:
2.1 - Policy + gpupdate + restart wuauserv
2.2 - Install from WU (progress is not live, 10-15 minutes)
- - - run in a separate terminal to see the progress 
- - - "dism /online /enable-feature /featurename:NetFx3"

Tools:
8.1 - Remove paths - input paths separated by ";" (not "; ")
8.2 - Start process as admin - input path

Sundry:
9.1 - Disable firewall
9.2 - Add kerio .104 connection (reconnect 1st connection in kerio before it)
$separator
Input:
"@
function Translit {
    [CmdletBinding()]
    param([string]$InputString)
    $map = @(
        @('а', 'a'), @('б', 'b'), @('в', 'v'), @('г', 'g'), @('д', 'd'),
        @('е', 'e'), @('ё', 'yo'), @('ж', 'zh'), @('з', 'z'), @('и', 'i'),
        @('й', 'y'), @('к', 'k'), @('л', 'l'), @('м', 'm'), @('н', 'n'),
        @('о', 'o'), @('п', 'p'), @('р', 'r'), @('с', 's'), @('т', 't'),
        @('у', 'u'), @('ф', 'f'), @('х', 'kh'), @('ц', 'ts'), @('ч', 'ch'),
        @('ш', 'sh'), @('щ', 'shch'), @('ъ', "'"), @('ы', 'y'), @('ь', "'"),
        @('э', 'e'), @('ю', 'yu'), @('я', 'ya'),
        
        @('А', 'A'), @('Б', 'B'), @('В', 'V'), @('Г', 'G'), @('Д', 'D'),
        @('Е', 'E'), @('Ё', 'Yo'), @('Ж', 'Zh'), @('З', 'Z'), @('И', 'I'),
        @('Й', 'Y'), @('К', 'K'), @('Л', 'L'), @('М', 'M'), @('Н', 'N'),
        @('О', 'O'), @('П', 'P'), @('Р', 'R'), @('С', 'S'), @('Т', 'T'),
        @('У', 'U'), @('Ф', 'F'), @('Х', 'Kh'), @('Ц', 'Ts'), @('Ч', 'Ch'),
        @('Ш', 'Sh'), @('Щ', 'Shch'), @('Ъ', "'"), @('Ы', 'Y'), @('Ь', "'"),
        @('Э', 'E'), @('Ю', 'Yu'), @('Я', 'Ya')
    )
    $result = $InputString
    foreach ($pair in $map) {
        $russianChar = $pair[0]
        $latinChar = $pair[1]
        $result = $result -replace $russianChar, $latinChar
    }
    return $result
}
function RemovePaths {
    Write-Host '[DELETION] Input paths separated by ";":'
    $paths = Read-Host ":"
    $paths = $paths -split ';'
    foreach ($path in $paths) {
        Remove-Item -Path $path -Force -Confirm:$false -ErrorAction Stop
    }
}
function StartProcessAsAdmin {
    Write-Host "Input file path to start:"
    $FilePath = Read-Host ':'
    Start-Process -FilePath $FilePath -Verb RunAs
}
function FindUserFiles {
    param (
    $Regex,
    $Paths = 'Desktop;Documents;Downloads;Downloads\Telegram Desktop'
    )
    $results = @()
    $Paths = $Paths -split ';'
    $users = Get-ChildItem -Path 'C:\Users' -Directory | Where-Object {
        Test-Path "$($_.FullName)\Downloads"
    }
    foreach ($user in $users) {
        foreach ($path in $Paths) {
            $absolutePath = "$($user.FullName)\$path"
            Get-ChildItem -Path $absolutePath -ErrorAction SilentlyContinue | Where-Object {
                $_.Name -match $Regex
            } | ForEach-Object {
                $results += $_.FullName
            }
        }
    }
    if ($results.Count -eq 0) {
        $results = 'none'
    } else {
        $results = $results -join "`n"
    }
    Write-Host $separator
    Write-Host "Regex: $Regex"
    # Write-Host "`nPaths: $Paths"
    Write-Host "Found:"
    WHT $results
}
function WH {
    param( $item )
    Write-Host $separator
    Write-Host $item
}
function WHT {
    param( $item )
    # Write-Host $separator
    $result = Translit $item
    Write-Host $result
}
function WHR {
    param($a, $b)
    if($?) { WH $a } 
    else { WH $b }
}
function RunCmd {
    param([string]$Command)
    $output = & cmd /c "chcp 65001>nul & $Command 2>&1"
    $output | ForEach-Object { 
        Translit $_ | Write-Host 
    }
    return $LASTEXITCODE
}
function specifyDownloadDomain {
    if ($global:downloadDomain) { return }
    WH 'Input domain: soft.?.ru'
    $global:downloadDomain = Read-Host ":"
}
function ChoosePathByRegex {
    param ($regex)
    $filenames = Get-ChildItem -Path "C:\" -Filter $regex | Select-Object -ExpandProperty Name
    if ($filenames.Count -eq 0) {
        throw "There are no `"$regex`" installers on C:\"
    }
    $menu = "0 - Exit installation`n"
    $n = 1
    foreach ($name in $filenames) {
        $menu += [string]$n + " - " + $name + "`n"
        $n++
    }
    WH "$menu"
    WH "Input:"
    $choice = Read-Host ":"
    if ($choice -eq '0') {
        return 'exit'
    }
    $choice = [int]$choice
    return $filenames[$choice-1]
}
function InstallMSI {
    param ($msiPath)
    $process = Start-Process -FilePath "msiexec.exe" -ArgumentList "/i", $msiPath, "/quiet" -Wait -PassThru
    if ($process.ExitCode -eq 0) {
        WH 'Installation: success.'
    } else {
        WH "Installation: error. Exit code: $($process.ExitCode)"
    }
}
function HandleAutoclosingRunBat {
    Set-Content -Path "$gspPath\RUN.bat" -Value $customRunBatCode
    WH 'RUN.bat is processed: auto close after launch and absolute path to .exe'
}
function CreateShortcut {
    $WshShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut("$gspPath\Genesys SIP Phone.lnk")
    $Shortcut.TargetPath = "$gspPath\GenesysSIPPhone.exe"
    $Shortcut.Arguments = "-config genesys"
    $Shortcut.IconLocation = "$gspPath\$iconFile"
    $Shortcut.WorkingDirectory = [System.IO.Path]::GetDirectoryName("$gspPath\RUN.bat")
    $Shortcut.WindowStyle = 1
    $Shortcut.Description = "Genesys SIP Phone"
    $Shortcut.Save();
    WH "Shortcut created: $gspPath"
}
function Set6signNumber {
    Write-Host "6-sign number: "
    $sixSignNumber = Read-Host ":"
    [xml]$phoneConfig = Get-Content -Path "$gspPath\Config\genesys_phoneConfig.xml" -ErrorAction Stop
    $phoneConfig.configuration['sip-endpoint'].user.setAttribute('name', $sixSignNumber);
    $phoneConfig.Save("$gspPath\Config\genesys_phoneConfig.xml")
    WH "Number $sixSignNumber set in genesys_phoneConfig.xml"
}
function SetUpdatePolicy {
    $ServicingPath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Servicing';
    if ( -not (Test-Path $ServicingPath) ) {
        New-Item -Path $ServicingPath
    }
    $props = Get-ItemProperty -Path $ServicingPath
    if ($props.PSObject.Properties.Name -contains "UseWindowsUpdate") {
        Remove-ItemProperty -Path $ServicingPath -Name "UseWindowsUpdate"
    }
    Set-ItemProperty -Path $ServicingPath -Name "RepairContentServerSource" -Value 2
    RunCmd "gpupdate /force"
    Stop-Service wuauserv -Force > $null
    WHR "[Success] stop wuauserv" "[Failure] stop wuauserv"
    Start-Service wuauserv > $null
    WHR "[Success] start wuauserv" "[Failure] start wuauserv"

}
function SetFirewallRules {
    $programPath = "$gspPath\GenesysSIPPhone.exe"
    $inboundRuleName = "Allow $programPath - Inbound"
    $outboundRuleName = "Allow $programPath - Outbound"

    $programRules = Get-NetFirewallApplicationFilter -Program $programPath -ErrorAction SilentlyContinue | Get-NetFirewallRule
    $inboundRuleExists = Where-Object { ($programRule.DisplayName -eq $inboundRuleName) -and ($programRule.Direction -eq 'Inbound') }
    $outboundRuleExists = Where-Object { ($programRule.DisplayName -eq $outboundRuleName) -and ($programRule.Direction -eq 'Outbound') }

    if (-not $inboundRuleExists) {
        New-NetFirewallRule `
            -DisplayName $inboundRuleName `
            -Direction Inbound `
            -Program $programPath `
            -Action Allow `
            -Profile Any `
            -Enabled True
        WH "Inbound rule created."
    } else {
        WH "Inbound rule has already been created before."
    }

    if (-not $outboundRuleExists) {
        New-NetFirewallRule `
            -DisplayName $outboundRuleName `
            -Direction Outbound `
            -Program $programPath `
            -Action Allow `
            -Profile Any `
            -Enabled True
        WH "Outbound rule created."
    } else {
        WH "Outbound rule has already been created before."
    }
}
function SetFullControlForAll {
    param($folderPath)
    $acl = Get-Acl $folderPath
    $accessRuleRu = New-Object System.Security.AccessControl.FileSystemAccessRule("Все", "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")
    $accessRuleEng = New-Object System.Security.AccessControl.FileSystemAccessRule("Everyone", "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")
    try { $acl.SetAccessRule($accessRuleRu) }
    catch { $acl.SetAccessRule($accessRuleEng) }
    Set-Acl -Path $folderPath -AclObject $acl -ErrorAction Stop
    WH "Full control for group 'All': $folderPath" 
}
function AddKerioConnection {
    param($server)
    $usercfgPath = "$env:APPDATA\Kerio\VpnClient\user.cfg"
    $usercfg = Get-Content -Path $usercfgPath -ErrorAction Stop
    [xml]$usercfgXML = Get-Content -Path $usercfgPath -ErrorAction Stop
    $connection = $usercfgXML.config.connections.connection;
    if ($connection[0] -ne $null) {
      $connection = $connection[0]
    }
    $username = $connection.username
    $password = $connection.password
    $connection104 = @"
<connection type="user">
  <description>$server</description>
  <server>$server</server>
  <username>$username</username>
  <password>$password</password>
  <savePassword>1</savePassword>
  <persistent>0</persistent>
</connection>

"@
    $injectedcfg = $usercfg -replace '(?=</connections>)', $connection104
    Set-Content -Path $usercfgPath -value $injectedcfg
    WH "Added connection: $server"
}
function CheckNetFx3 {
    if (Test-Path $NetFx3Path) {
        $install = Get-ItemProperty -Path $NetFx3Path -Name Install
        if ($install.Install -eq 1) {
            Write-Host "`nNetFx3: enabled"
        } else {
            Write-Host "`nNetFx3: disabled"
        }
    } else {
        Write-Host "`nNetFx3 registry key is not found"
    }
}
function WriteHostIfPathExist {
    param($path)
    if (Test-Path $path) {
        Write-Host "Found: $path"
    }
}
function Precheck {
    CheckNetFx3
    

    # citrix version 
    $version = "Not Found"
    if (Test-Path $citrixSelfServicePath) {
        $version = (Get-Item $citrixSelfServicePath).VersionInfo.ProductVersion
    }
    Write-Host "Citrix file version: $version"
    

    # kerio version
    $version = "Not Found"
    if (Test-Path $kvpncguiPath) {
        $version = (Get-Item $kvpncguiPath).VersionInfo.ProductVersion
    }
    Write-Host "Kerio file version: $version"


    FindUserFiles "citrix"
    FindUserFiles "kerio"
    FindUserFiles "genesys"
    FindUserFiles "run.bat"

    Write-Host $separator
    WriteHostIfPathExist "C:\Genesys SIP Phone"
    WriteHostIfPathExist "C:\Genesys SIP Phone.zip"
    WriteHostIfPathExist "C:\Genesys_SIP_Phone.zip"
}
function CheckIfGenesysAlreadyInstalledByThisScript {
    if (Test-Path $gspPath) {
        WH "Already exists: $gspPath"
    }
    if (Test-Path $shortcutPath) {
        WH "Already exists: $shortcutPath"
    }
}
function HandleExpanding {
    $path = $archivePathCC
    if (-not (Test-Path $archivePathCC) ) {
        if (Test-Path $archivePath) {
            $path = $archivePath
        } else {
            throw "Archive is not found."
        }
    }
    Expand-Archive -Path $path -DestinationPath "C:\Users\Public\Downloads" -Force -ErrorAction Stop
    WH "Archive expanded: C:\Users\Public\Downloads"
}
function RemoveArchive {
    if (Test-Path $archivePath) {
        Remove-Item -Path $archivePath -ErrorAction Stop
    }
    if (Test-Path $archivePathCC) {
        Remove-Item -Path $archivePathCC -ErrorAction Stop
    }
}
function Case {
    $ErrorActionPreference = 'Stop'
    Write-Host $menuString
    $action = Read-Host ":"
    if ($action -eq "0") {
        return 'exit';
    } elseif ($action -eq "0.1") {
        Precheck
    } elseif ($action -eq "1") {
        HandleExpanding
        RemoveArchive
        WH "Archive removed."
        CreateShortcut
        Copy-Item -Path "$gspPath\Genesys SIP Phone.lnk" -Destination $shortcutPath
        WH "Shortcut created: C:\Users\Public\Desktop"
        SetFirewallRules
        HandleAutoclosingRunBat
    } elseif ($action -eq "1.1") {
        HandleExpanding
    } elseif ($action -eq "1.2") {
        RemoveArchive
        WH "Archive removed."
    } elseif ($action -eq "1.3") {
        CreateShortcut
        Copy-Item -Path "$gspPath\Genesys SIP Phone.lnk" -Destination $shortcutPath
        WH "Shortcut created: C:\Users\Public\Desktop"
    } elseif ($action -eq "1.4") {
        SetFirewallRules
    } elseif ($action -eq "1.5") {
        HandleAutoclosingRunBat
    } elseif ($action -eq "1.6") {
        Set6signNumber
    } elseif ($action -eq "1.7") {
        SetFullControlForAll $machineKeysPath
    } elseif ($action -eq "1.8") {
        SetFullControlForAll $cryptoKeysPath
    } elseif ($action -eq "2.1") {
        SetUpdatePolicy
    } elseif ($action -eq "2.2") {
        RunCmd "dism.exe /online /enable-feature /featurename:NetFX3"
        CheckNetFx3
    } elseif ($action -eq "8.1") {
        RemovePaths
    } elseif ($action -eq "8.2") {
        StartProcessAsAdmin
    } elseif ($action -eq "9.1") {
        Set-NetFirewallProfile -Profile Domain,Public,Private -Enabled False
        WH "Firewall is disabled."
    } elseif ($action -eq "9.2") {
        AddKerioConnection '194.0.162.104'
    # } elseif ($action -eq "13") {
        # specifyDownloadDomain
        # Invoke-WebRequest -Uri ('https://soft.' + $downloadDomain + '.ru/Genesys_SIP_Phone.zip') -OutFile "C:\Genesys_SIP_Phone.zip"
        # WHR "[Success] download archive" "[Failure] download archive"
    # } elseif ($action -eq "16") {
        # $filename = ChoosePathByRegex "*kerio*.msi"
        # if ($filename -eq 'exit') { return }
        # InstallMSI "C:\$filename"
    # } elseif ($action -eq "17") {
        # $filename = ChoosePathByRegex "*citrix*.exe"
        # if ($filename -eq 'exit') {
        #     return $null
        # }
        # $process = Start-Process -FilePath "C:\$filename" -ArgumentList "/silent /noreboot /forceinstall /AutoUpdateCheck=disabled /EnableCEIP=false" -Wait -NoNewWindow -PassThru
        # if ($process.ExitCode -eq 0) {
        #     WH 'Installation: success.'
        # } else {
        #     WH "Installation: error. Exit code: $($process.ExitCode)"
        # }
    }
}
function Attempt {
    param([ScriptBlock]$Callback)
    $result = $null;
    try {
        $result = $Callback.Invoke() 
    }
    catch { Write-Error $_ }
    return $result;
}
Write-Host "`n==== MTC BC setup-helper ====`n[russian letters are given in transliteration]"
Precheck
CheckIfGenesysAlreadyInstalledByThisScript
$caseResult = $null
while ($caseResult -ne 'exit') {
    $caseResult = Attempt -Callback ${function:Case}
}
