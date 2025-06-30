[console]::InputEncoding = [System.Text.Encoding]::UTF8
[console]::OutputEncoding = [System.Text.Encoding]::UTF8

$iconFile = "app_yellow.ico";
$archivePath = "$PSScriptRoot\Genesys SIP Phone.zip"
$kvpncguiPath = "C:\Program Files (x86)\Kerio\VPN Client\kvpncgui.exe"
$gspPath = "C:\Users\Public\Downloads\Genesys SIP Phone"
$menuString = @'

0 - Stop
1 - Kerio VPN Version
2 - Expand archive
3 - Create shortcut
4 - Set 6-sign number
5 - Remove archive
6 - Remove this script file
7 - Check NetFx3
8 - Policy + gpupdate + restart wuauserv
9 - Dism install NetFx3 from WU (long 5.9%-pause is ok)
10 - Set firewall rule for GenesysSIPPhone.exe
11 - Add kerio .104 connection (reconnect 1st connection in kerio before it)
'@

Write-Host $menuString
while ($true) {
    Write-Host "Input:"
    $action = Read-Host "-"
    if ($action -eq "0") {
        break;
    } elseif ($action -eq "1") {
        $kerioVersion = "Not Installed"
        if (Test-Path $kvpncguiPath) {
            $kerioVersion = (Get-Item $kvpncguiPath).VersionInfo.ProductVersion
        }
        Write-Host "`nKerio VPN Version: $kerioVersion"
    } elseif ($action -eq "2") {
        if (-not (Test-Path $archivePath) ) {
            throw "$archivePath is not found."
        }
        Expand-Archive -Path $archivePath -DestinationPath "C:\Users\Public\Downloads" -Force
        Write-Host "`nArchive expanded: C:\Users\Public\Downloads"
    } elseif ($action -eq "3") {
        $WshShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut("C:\Users\Public\Desktop\Genesys SIP Phone.lnk")
        $Shortcut.TargetPath = "$gspPath\RUN.bat"
        $Shortcut.IconLocation = "$gspPath\$iconFile"
        $Shortcut.WorkingDirectory = [System.IO.Path]::GetDirectoryName("$gspPath\RUN.bat")
        $Shortcut.WindowStyle = 1
        $Shortcut.Description = "Link RUN.bat"
        $Shortcut.Save();
        Write-Host "`nShortcut created: C:\Users\Public\Desktop"
    } elseif ($action -eq "4") {
        Write-Host "`n6-sign number: "
        $sixSignNumber = Read-Host "-"
        [xml]$phoneConfig = Get-Content -Path "$gspPath\Config\genesys_phoneConfig.xml"
        $phoneConfig.configuration['sip-endpoint'].user.setAttribute('name', $sixSignNumber);
        $phoneConfig.Save("$gspPath\Config\genesys_phoneConfig.xml")
        Write-Host "`nNumber $sixSignNumber set in genesys_phoneConfig.xml"
    } elseif ($action -eq "5") {
        Remove-Item -Path $archivePath
        Write-Host "`n$archivePath removed."
    } elseif ($action -eq "6") {
        $scriptPath = $MyInvocation.MyCommand.Path
        Start-Sleep -Seconds 1
        Remove-Item -Path $path -Force
        Write-Host "`n$scriptPath removed."

    } elseif ($action -eq "7") {
        $netfx3Path = "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.5"
        if (Test-Path $netfx3Path) {
            $install = Get-ItemProperty -Path $netfx3Path -Name Install
            if ($install.Install -eq 1) {
                Write-Host "[+] NetFx3 is enabled"
            } else {
                Write-Host "[-] NetFx3 is disabled"
            }
        } else {
            Write-Host "[x] NetFx3 registry key is not found"
        }
    } elseif ($action -eq "8") {
         $props = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Servicing"
        if ($props.PSObject.Properties.Name -contains "UseWindowsUpdate") {
            Remove-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Servicing" -Name "UseWindowsUpdate"
        }
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Servicing" -Name "RepairContentServerSource" -Value 2
        gpupdate /force
        net stop wuauserv
        net start wuauserv
    } elseif ($action -eq "9") {
        dism.exe /online /enable-feature /featurename:NetFX3
    } elseif ($action -eq "10") {
        $programPath = "$gspPath\GenesysSIPPhone.exe"
        $inboundRuleName = "Allow $programPath - Inbound"
        $outboundRuleName = "Allow $programPath - Outbound"

        $programRules = Get-NetFirewallApplicationFilter -Program $programPath -ErrorAction SilentlyContinue | Get-NetFirewallRule
        $inboundRuleExists = Where-Object { ($programRule.DisplayName -eq $inboundRuleName) -and ($programRule.Direction -eq 'Inbound') }
        $outboundRuleExists = Where-Object { ($programRule.DisplayName -eq $outboundRuleName) -and ($programRule.Direction -eq 'Outbound') }

        if (-not $inboundRuleExists) {
            Write-Host "Creating inbound rule..."
            New-NetFirewallRule `
                -DisplayName $inboundRuleName `
                -Direction Inbound `
                -Program $programPath `
                -Action Allow `
                -Profile Any `
                -Enabled True
            Write-Host "Inbound rule created."
        } else {
            Write-Host "Inbound rule has already been created before."
        }

        if (-not $outboundRuleExists) {
            Write-Host "Creating outbound rule..."
            New-NetFirewallRule `
                -DisplayName $outboundRuleName `
                -Direction Outbound `
                -Program $programPath `
                -Action Allow `
                -Profile Any `
                -Enabled True
            Write-Host "Outbound rule created."
        } else {
            Write-Host "Outbound rule has already been created before."
        }
    } elseif ($action -eq "11") {
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
  <description>194.0.162.104</description>
  <server>194.0.162.104</server>
  <username>$username</username>
  <password>$password</password>
  <savePassword>1</savePassword>
  <persistent>0</persistent>
</connection>

"@
        $injectedcfg = $usercfg -replace '(?=</connections>)', $connection104
        Set-Content -Path $usercfgPath -value $injectedcfg
    }
}
