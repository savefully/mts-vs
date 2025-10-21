@echo off
chcp 65001 >nul

set public=C:\Users\Public
set c_path=C:\Genesys SIP Phone 
set run_path=%public%\Desktop\RUN.bat
set _7zip=C:\Program Files\7-Zip\7z.exe
set winrar=C:\Program Files\WinRAR\WinRAR.exe
set lnk_path=%public%\Desktop\Genesys SIP Phone.lnk
set exe_path=%public%\Downloads\Genesys SIP Phone\GenesysSIPPhone.exe
set policies_servicing_path=HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Servicing
set archive_path=undefined

net session 1>nul 2>&1
if %errorlevel% neq 0 (
	echo [Error] not admin
	pause
	exit
) 

:loop
	echo.
	echo.
	echo ==== Genesys SIP Phone installation assistant ====
	echo.
	call :checks
	echo.
	echo Enter number of option:
	echo 0 - Exit
	echo 1 - Install Genesys SIP Phone (2-5 options)
	echo 2 - Expand archive
	echo 2.1 -- 7-Zip
	echo 2.2 -- WinRAR
	echo 2.3 -- Powershell
	echo 3 - Remove archive
	echo 4 - Add firewall rules for GenesysSIPPhone.exe
	echo 5 - Set RUN.bat on public desktop
	echo === .NET 3.5 ===
	echo 6 - Policy + gpupdate + restart wuauserv
	echo 7 - Install (dism)
	echo.
	set option=undefined
	set /p option=Number:
	if %option%==0 (
		goto :end
	)
	if %option%==1 (
		call :expand_archive
		call :remove_archive
		call :firewall_rules
		call :runbat
		goto :loop
	)
	if %option%==2 (
		call :expand_archive
		goto :loop
	)
	if %option%==2.1 (
		call :_7zip_
		goto :loop
	)
	if %option%==2.2 (
		call :winrar_
		goto :loop
	)
	if %option%==2.3 (
		call :powershell_expanding
		goto :loop
	)
	if %option%==3 (
		call :remove_archive
		goto :loop
	)
	if %option%==4 (
		call :firewall_rules
		goto :loop
	)
	if %option%==5 (
		call :runbat
		goto :loop
	)
	if %option%==6 (
		call :policy
		goto :loop
	)
	if %option%==7 (
		dism.exe /online /enable-feature /featurename:NetFX3
		goto :loop
	)
	echo No option: %option%.
	goto :loop

:checks
	reg query "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.5" /v Install 1>nul 2>&1
	if %errorlevel% neq 0 (
		echo [x] .NET 3.5 - not installed
	) else (
		reg query "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.5" /v Install | findstr "0x1" 1>nul 2>&1
		if %errorlevel% neq 0 (
			echo [-] .NET 3.5 - installed, but disabled
		) else (
			echo [+] .NET 3.5 - installed, enabled
		)
	)

	if exist C:\Genesys_SIP_Phone.zip (
		set archive_path=C:\Genesys_SIP_Phone.zip
	)
	if exist "C:\Genesys SIP Phone.zip" (
		set archive_path=C:\Genesys SIP Phone.zip
	)
	echo Selected source archive: %archive_path%
	if exist "%c_path%" (echo Found: %c_path%)
	if exist "%exe_path%" (echo Found: %exe_path%)
	if exist "%run_path%" (echo Found: %run_path%)
	if exist "%lnk_path%" (echo Found: %lnk_path%)
	goto :eof

:expand_archive
	if "%archive_path%"==undefined (
		pause
		color 0c
		echo C:\Genesys_SIP_Phone.zip and C:\Genesys SIP Phone.zip are not found. Can't expand.
		pause
		color 0F
		goto :loop
	)
	if exist "%_7zip%" (goto :_7zip_)
	if exist "%winrar%" (goto :winrar_)
	where powershell >nul
	if %errorlevel% equ 0 (
		goto :powershell_expanding
	) else (
		color 0c
		echo WinRAR, 7-zip and powershell are not found. No way to expand archive.
		pause
		color 0F
		goto :eof
	)
	:_7zip_
		echo Selected archiver: 7-Zip
		call :safely ^"%_7zip%^" x ^"%archive_path%^" -o^"%public%\Downloads\^" -y
		echo Expanded
		goto :eof
	:winrar_
		echo Selected archiver: WinRAR
		call :safely ^"%winrar%^" x -y ^"%archive_path%^" ^"%public%\Downloads\^"
		echo Expanded
		goto :eof
	:powershell_expanding
		echo Selected archiver: Powershell's Expand-Archive method
		call :safely powershell -ExecutionPolicy Bypass -Command ^"Expand-Archive -Path '%archive_path%' -DestinationPath '%public%\Downloads' -Force -ErrorAction Stop^"
		echo Expanded
		goto :eof	

:remove_archive
	if "%archive_path%"==undefined (
		pause
		color 0c
		echo C:\Genesys_SIP_Phone.zip and C:\Genesys SIP Phone.zip are not found. Can't remove.
		pause
		color 0F
		goto :loop
	)
	del /q "%archive_path%"
	goto :eof

:firewall_rules
	echo Check inbound firewall rule
	netsh advfirewall firewall show rule name="Allow In - %exe_path%" 1>nul 2>&1
	if %errorlevel% neq 0 (
		call :safely netsh advfirewall firewall add rule name=^"Allow In - %exe_path%^" dir=in program=^"%exe_path%^" action=allow
		echo Added: inbound firewall rule
	)
	echo Check outbound firewall rule
	netsh advfirewall firewall show rule name="Allow Out - %exe_path%" 1>nul 2>&1
	if %errorlevel% neq 0 (
		call :safely netsh advfirewall firewall add rule name=^"Allow Out - %exe_path%^" dir=out program=^"%exe_path%^" action=allow
		echo Added: outbound firewall rule
	)
	goto :eof

:runbat
	(
		echo @echo off
		echo cd ^"C:\Users\Public\Downloads\Genesys SIP Phone^"
		echo start "GSP" GenesysSIPPhone.exe -config genesys
	) > "%run_path%"
	echo Created: RUN.bat
	goto :eof

:policy
	call :safely reg add ^"%policies_servicing_path%^" /v RepairContentServerSource /t REG_DWORD /d 2 /f
	reg delete "%policies_servicing_path%" /v UseWindowsUpdate /f 1>nul 2>&1
	gpupdate /force
	net stop wuauserv
	net start wuauserv
	goto :eof
	
:safely
	%* 1>nul 2>&1
	if %errorlevel% neq 0 (
		color 0c
	    echo [Error] Command: %*
	    echo Code: %errorlevel%
	    pause
	    color 0F
	    goto :loop
	)
	goto :eof
:end
pause
