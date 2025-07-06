[console]::InputEncoding = [System.Text.Encoding]::UTF8
[console]::OutputEncoding = [System.Text.Encoding]::UTF8

$iconFile = 'app_yellow.ico';
$archivePath = 'C:\Genesys SIP Phone.zip'
$gspPath = 'C:\Users\Public\Downloads\Genesys SIP Phone'
$shortcutPath = 'C:\Users\Public\Desktop\Genesys SIP Phone.lnk'
$machineKeysPath = 'C:\ProgramData\Microsoft\Crypto\RSA\MachineKeys'
$NetFx3Path = "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.5"
$kvpncguiPath = 'C:\Program Files (x86)\Kerio\VPN Client\kvpncgui.exe'
$cryptoKeysPath = 'C:\ProgramData\Application Data\Microsoft\Crypto\Keys'
$citrixSelfServicePath = 'C:\Program Files (x86)\Citrix\ICA Client\SelfServicePlugin\SelfService.exe'
$customRunBatCode = @'
@echo off
start "Genesys SIP Phone" "C:\Users\Public\Downloads\Genesys SIP Phone\GenesysSIPPhone.exe" -config genesys
exit   
'@
$separator = '----------------------------------------------------------------------------------------------------'
$menuString = @"
$separator
0 - Exit

1 - Precheck (NetFx3 state; Kerio and Citrix versions)
2 - Expand archive: $archivePath > $gspPath
3 - Remove archive: $archivePath
4 - Create shortcut: $shortcutPath
5 - Set 6-sign number

ComponentActivator fixes:
6 - $machineKeysPath 
7 - $cryptoKeysPath

NetFx3 installation:
8 - Policy + gpupdate + restart wuauserv
9 - Install from WU (5.9%-pause is ok)

Tools:
10 - Set firewall rules for GenesysSIPPhone.exe
11 - Add kerio .104 connection (reconnect 1st connection in kerio before it)
12 - Set custom RUN.bat (auto close)
$separator
Input:
"@
function WH {
    param( $item )
    Write-Host $separator
    Write-Host $item
    Write-Host $separator
}
function HandleAutoclosingRunBat {
    Rename-Item -Path "$gspPath\RUN.bat" -NewName "ORIGINAL_RUN.bat"
    Set-Content -Path "$gspPath\RUN.bat" -Value $customRunBatCode
    WH 'RUN.bat processed.'
}
function CreateShortcut {
    $WshShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut($shortcutPath)
    $Shortcut.TargetPath = "$gspPath\RUN.bat"
    $Shortcut.IconLocation = "$gspPath\$iconFile"
    $Shortcut.WorkingDirectory = [System.IO.Path]::GetDirectoryName("$gspPath\RUN.bat")
    $Shortcut.WindowStyle = 1
    $Shortcut.Description = "Link RUN.bat"
    $Shortcut.Save();
    WH "Shortcut created: C:\Users\Public\Desktop"
}
function Set6signNumber {
    Write-Host "6-sign number: "
    $sixSignNumber = Read-Host "-"
    [xml]$phoneConfig = Get-Content -Path "$gspPath\Config\genesys_phoneConfig.xml"
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
    gpupdate /force
    net stop wuauserv
    net start wuauserv
}
function SetFirewallRules {
    $programPath = "$gspPath\GenesysSIPPhone.exe"
    $inboundRuleName = "Allow $programPath - Inbound"
    $outboundRuleName = "Allow $programPath - Outbound"

    $programRules = Get-NetFirewallApplicationFilter -Program $programPath -ErrorAction SilentlyContinue | Get-NetFirewallRule
    $inboundRuleExists = Where-Object { ($programRule.DisplayName -eq $inboundRuleName) -and ($programRule.Direction -eq 'Inbound') }
    $outboundRuleExists = Where-Object { ($programRule.DisplayName -eq $outboundRuleName) -and ($programRule.Direction -eq 'Outbound') }

    if (-not $inboundRuleExists) {
        WH "Creating inbound rule..."
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
        WH "Creating outbound rule..."
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
    param(
        $folderPath
    )
    $acl = Get-Acl $folderPath
    $accessRuleRu = New-Object System.Security.AccessControl.FileSystemAccessRule("Все", "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")
    $accessRuleEng = New-Object System.Security.AccessControl.FileSystemAccessRule("Everyone", "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")
    try { $acl.SetAccessRule($accessRuleRu) }
    catch { $acl.SetAccessRule($accessRuleEng) }
    Set-Acl -Path $folderPath -AclObject $acl
    WH "Full control for group 'All': $folderPath"
}
function AddKerioConnection {
    param (
        $server
    )
    $usercfgPath = "$env:APPDATA\Kerio\VpnClient\user.cfg"
    $usercfg = Get-Content -Path $usercfgPath
    [xml]$usercfgXML = Get-Content -Path $usercfgPath
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

function Precheck {
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

    $version = "Not Found"
    if (Test-Path $citrixSelfServicePath) {
        $version = (Get-Item $citrixSelfServicePath).VersionInfo.ProductVersion
    }
    Write-Host "Citrix file version: $version"

    $version = "Not Found"
    if (Test-Path $kvpncguiPath) {
        $version = (Get-Item $kvpncguiPath).VersionInfo.ProductVersion
    }
    Write-Host "Kerio file version: $version"
}
function CheckIfGenesysAlreadyInstalledByThisScript {
    if (Test-Path $gspPath) {
        WH "ALREADY EXISTS: $gspPath"
    }
    if (Test-Path $shortcutPath) {
        WH "ALREADY EXISTS: $shortcutPath"
    }
}
function Case {    
    Write-Host $menuString
    $action = Read-Host "-"
    if ($action -eq "0") {
        return 'exit';
    } elseif ($action -eq "1") {
        Precheck
    } elseif ($action -eq "2") {
        if (-not (Test-Path $archivePath) ) {
            throw "$archivePath is not found."
        }
        Expand-Archive -Path $archivePath -DestinationPath "C:\Users\Public\Downloads" -Force
        WH "Archive expanded: C:\Users\Public\Downloads"
    } elseif ($action -eq "3") {
        Remove-Item -Path $archivePath
        WH "$archivePath removed."
    } elseif ($action -eq "4") {
       CreateShortcut 
    } elseif ($action -eq "5") {
        Set6signNumber
    } elseif ($action -eq "6") {
        SetFullControlForAll $machineKeysPath
    } elseif ($action -eq "7") {
        SetFullControlForAll $cryptoKeysPath
    } elseif ($action -eq "8") {
        SetUpdatePolicy
    } elseif ($action -eq "9") {
        dism.exe /online /enable-feature /featurename:NetFX3
    } elseif ($action -eq "10") {
        SetFirewallRules
    } elseif ($action -eq "11") {
        AddKerioConnection '194.0.162.104'
    } elseif ($action -eq "12") {
        HandleAutoclosingRunBat
    }
}
function Attempt {
    param([ScriptBlock]$Callback)
    $result = $null;
    try { $result = $Callback.Invoke() }
    catch { Write-Error $_ }
    return $result;
}
Write-Host "`n==== MTC BC software script-helper ===="
Precheck
CheckIfGenesysAlreadyInstalledByThisScript
$caseResult = $null
while ($caseResult -ne 'exit') {
    $caseResult = Attempt -Callback ${function:Case}
}
