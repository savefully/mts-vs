@echo off
chcp 866 >nul

set "download_domain=none"
set "archive_path=­¥ ­ ©¤¥­"
set "public=C:\Users\Public"
set "c_path=C:\Genesys SIP Phone" 
set "run_path=%public%\Desktop\RUN.bat"
set "_7zip=C:\Program Files\7-Zip\7z.exe"
set "winrar=C:\Program Files\WinRAR\WinRAR.exe"
set "lnk_path=%public%\Desktop\Genesys SIP Phone.lnk"
set "exe_path=%public%\Downloads\Genesys SIP Phone\GenesysSIPPhone.exe"
set "policies_servicing_path=HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Servicing"

net session 1>nul 2>&1
if %errorlevel% neq 0 (
	echo ¥â ¯à ¢  ¤¬¨­¨áâà â®à 
	pause
	exit
)

if "%1"=="" (
	echo.
	echo [ˆ­ä®] „®¬¥­ áª ç¨¢ ­¨ï ­¥ ¯¥à¥¤ ­
) else (
	echo „®¬¥­ áª ç¨¢ ­¨ï: %~1
	set "download_domain=%~1"
)

:loop
	echo.
	echo.
	echo ++++++++ €áá¨áâ¥­â ãáâ ­®¢ª¨ Genesys SIP Phone ++++++++
	echo.
	call :checks
	echo.
	echo ‚¢¥¤¨â¥ ­®¬¥à ¯ã­ªâ :
	echo 0 - ‚ëå®¤
	echo 1 - “áâ ­®¢¨âì Genesys SIP Phone (2-6 ¯ã­ªâë)
	echo 2 = ‘ª ç âì  àå¨¢
	echo 3 = ˆ§¢«¥çì  àå¨¢
	echo 3.1 -- 7-Zip
	echo 3.2 -- WinRAR
	echo 3.3 -- Powershell
	echo 4 - “¤ «¨âì  àå¨¢
	echo 5 - ‚­¥áâ¨ ¨áª«îç¥­¨ï ¡à ­¤¬ ãíà  ¤«ï GenesysSIPPhone.exe
	echo 6 - ‡ ¯¨á âì RUN.bat ­  ®¡é¨© à ¡®ç¨© áâ®«
	echo #### .NET 3.5 ####
	echo 7 - Ž¡­®¢«¥­¨¥ ¯®«¨â¨ª¨ ¨ ¯¥à¥§ ¯ãáª wuauserv
	echo 8 - “áâ ­®¢ª  ç¥à¥§ DISM
	echo.
	set "option=ãáâ®"
	set /p "option=—¨á«®:"
	if %option%==0 (
		goto :end
	)
	if %option%==1 (
		call :download_archive && call :expand_archive && call :remove_archive && call :firewall_rules && call :runbat
		goto :loop
	)
	if %option%==2 (
		call :download_archive
		goto :loop
	)
	if %option%==3 (
		call :expand_archive
		goto :loop
	)
	if %option%==3.1 (
		call :_7zip_
		goto :loop
	)
	if %option%==3.2 (
		call :winrar_
		goto :loop
	)
	if %option%==3.3 (
		call :powershell_expanding
		goto :loop
	)
	if %option%==4 (
		call :remove_archive
		goto :loop
	)
	if %option%==5 (
		call :firewall_rules
		goto :loop
	)
	if %option%==6 (
		call :runbat
		goto :loop
	)
	if %option%==7 (
		call :policy
		goto :loop
	)
	if %option%==8 (
		dism.exe /online /enable-feature /featurename:NetFX3
		goto :loop
	)
	echo.
	echo ¥â ¯ã­ªâ : %option%.
	goto :loop

:checks
	reg query "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.5" /v Install 1>nul 2>&1
	if %errorlevel% neq 0 (
		echo [x] .NET 3.5 - ­¥ ãáâ ­®¢«¥­
	) else (
		reg query "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.5" /v Install | findstr "0x1" 1>nul 2>&1
		if %errorlevel% neq 0 (
			echo [-] .NET 3.5 - ãáâ ­®¢«¥­, ­® ®âª«îç¥­
		) else (
			echo [+] .NET 3.5 - ãáâ ­®¢«¥­, ¢ª«îç¥­
		)
	)
	set "archive_path=­¥ ­ ©¤¥­"
	if exist C:\Genesys_SIP_Phone.zip (
		set "archive_path=C:\Genesys_SIP_Phone.zip"
	)
	if exist "C:\Genesys SIP Phone.zip" (
		set "archive_path=C:\Genesys SIP Phone.zip"
	)
	echo €àå¨¢ ¤«ï ¨§¢«¥ç­¨ï: %archive_path%
	if exist "%c_path%" (echo Ž¡­ àã¦¥­: %c_path%)
	if exist "%exe_path%" (echo Ž¡­ àã¦¥­: %exe_path%)
	if exist "%run_path%" (echo Ž¡­ àã¦¥­: %run_path%)
	if exist "%lnk_path%" (echo Ž¡­ àã¦¥­: %lnk_path%)
	goto :eof

:download_archive
	echo.
	if "%download_domain%"=="none" (
		set /p "download_domain=‚¢¥¤¨â¥ ¤®¬¥­, ª®â®àë© ¤®«¦¥­ áâ®ïâì ¢¬¥áâ® * ^> soft.*.ru:"
	)
	call :safely curl -o ^"C:\Genesys_SIP_Phone.zip^" ^"https://soft.%download_domain%.ru/Genesys_SIP_Phone.zip^" || exit /b 1
	set "archive_path=C:\Genesys_SIP_Phone.zip"
	echo €àå¨¢ áª ç ­
	goto :eof

:expand_archive
	echo.
	if "%archive_path%"=="­¥ ­ ©¤¥­" (
		color 0c
		echo C:\Genesys_SIP_Phone.zip ¨ C:\Genesys SIP Phone.zip ­¥ ­ ©¤¥­ë. ¥â  àå¨¢  ¤«ï ¨§¢«¥ç¥­¨ï.
		pause
		color 0F
		exit /b 1
	)
	if exist "%_7zip%" (
		call :_7zip_
		if %errorlevel% equ 0 (goto :eof)
	)
	if exist "%winrar%" (
		call :winrar_
		if %errorlevel% equ 0 (goto :eof)
	)
	where powershell >nul
	if %errorlevel% equ 0 (
		call :powershell_expanding
		if %errorlevel% equ 0 (goto :eof)
	) else (
		color 0c
		echo WinRAR, 7-zip ¨ powershell ­¥ ­ ©¤¥­ë. ¥â ¢®§¬®¦­®áâ¨ ¨§¢«¥çì  àå¨¢.
		pause
		color 0F
		exit /b 1
	)
	:_7zip_
		echo.
		echo ‚ë¡à ­: 7-Zip
		call :safely ^"%_7zip%^" x ^"%archive_path%^" -o^"%public%\Downloads\^" -y
		if %errorlevel% equ 0 (
			echo €àå¨¢ ¨§¢«¥ç¥­
			goto :eof
		)
		echo ¥ ã¤ «®áì ¨§¢«¥çì  àå¨¢ ¨§-§  ®è¨¡ª¨
		exit /b 1
	:winrar_
		echo.
		echo ‚ë¡à ­: WinRAR
		call :safely ^"%winrar%^" x -y ^"%archive_path%^" ^"%public%\Downloads\^"
		if %errorlevel% equ 0 (
			echo €àå¨¢ ¨§¢«¥ç¥­
			goto :eof
		)
		echo ¥ ã¤ «®áì ¨§¢«¥çì  àå¨¢ ¨§-§  ®è¨¡ª¨
		exit /b 1
	:powershell_expanding
		echo.
		echo ‚ë¡à ­: Powershell's Expand-Archive
		call :safely powershell -ExecutionPolicy Bypass -Command ^"Expand-Archive -Path '%archive_path%' -DestinationPath '%public%\Downloads' -Force -ErrorAction Stop^"
		if %errorlevel% equ 0 (
			echo €àå¨¢ ¨§¢«¥ç¥­
			goto :eof
		)
		echo ¥ ã¤ «®áì ¨§¢«¥çì  àå¨¢ ¨§-§  ®è¨¡ª¨
		exit /b 1

:remove_archive
	echo.
	if "%archive_path%"=="­¥ ­ ©¤¥­" (
		color 0c
		echo [ˆ­ä®] C:\Genesys_SIP_Phone.zip ¨ C:\Genesys SIP Phone.zip ­¥ ­ ©¤¥­ë. ¥â  àå¨¢  ¤«ï ã¤ «¥­¨ï.
		pause
		color 0F
	) else (
		del /q "%archive_path%"
		echo €àå¨¢ ã¤ «¥­
	)
	goto :eof

:firewall_rules
	echo.
	echo à®¢¥àª  ­ «¨ç¨ï ¯à ¢¨«  ¢å®¤ïé¨å á®¥¤¨­¥­¨©...
	netsh advfirewall firewall show rule name="Allow In - %exe_path%" 1>nul 2>&1
	if %errorlevel% neq 0 (
		netsh advfirewall firewall add rule name=^"Allow In - %exe_path%^" dir=in program=^"%exe_path%^" action=allow 1>nul 2>&1
		if %errorlevel% neq 0 (
			echo „®¡ ¢«¥­® ¯à ¢¨«® ¢å®¤ïé¨å á®¥¤¨­¥­¨©
		) else (
			echo [à¥¤ã¯à¥¦¤¥­¨¥] ¥ ã¤ «®áì ¤®¡ ¢¨âì ¯à ¢¨«® ¢å®¤ïé¨å á®¥¤¨­¥­¨©
		)
	)
	echo à®¢¥àª  ­ «¨ç¨ï ¯à ¢¨«  ¨áå®¤ïé¨å á®¥¤¨­¥­¨©...
	netsh advfirewall firewall show rule name="Allow Out - %exe_path%" 1>nul 2>&1
	if %errorlevel% neq 0 (
		netsh advfirewall firewall add rule name=^"Allow Out - %exe_path%^" dir=out program=^"%exe_path%^" action=allow 1>nul 2>&1
		if %errorlevel% neq 0 (
			echo „®¡ ¢«¥­® ¯à ¢¨«® ¨áå®¤ïé¨å á®¥¤¨­¥­¨©
		) else (
			echo [à¥¤ã¯à¥¦¤¥­¨¥] ¥ ã¤ «®áì ¤®¡ ¢¨âì ¯à ¢¨«® ¨áå®¤ïé¨å á®¥¤¨­¥­¨©
		)
	)
	goto :eof

:runbat
	echo.
	(
		echo @echo off
		echo cd ^"C:\Users\Public\Downloads\Genesys SIP Phone^"
		echo start "GSP" GenesysSIPPhone.exe -config genesys
	) > "%run_path%"
	echo RUN.bat § ¯¨á ­
	goto :eof

:policy
	echo.
	call :safely reg add ^"%policies_servicing_path%^" /v RepairContentServerSource /t REG_DWORD /d 2 /f || exit /b 1
	reg delete "%policies_servicing_path%" /v UseWindowsUpdate /f 1>nul 2>&1
	gpupdate /force
	net stop wuauserv
	net start wuauserv
	goto :eof
	
:safely
	echo.
	%* 1>nul 2>&1
	if %errorlevel% neq 0 (
		color 0c
	    echo [Žè¨¡ª ] Š®¬ ­¤ : %*
	    echo Š®¤: %errorlevel%
	    pause
	    color 0F
	    exit /b 1
	)
	goto :eof
:end

pause
