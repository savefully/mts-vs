@echo off
chcp 866 >nul

set public=C:\Users\Public
set c_path=C:\Genesys SIP Phone 
set run_path=%public%\Desktop\RUN.bat
set _7zip=C:\Program Files\7-Zip\7z.exe
set winrar=C:\Program Files\WinRAR\WinRAR.exe
set lnk_path=%public%\Desktop\Genesys SIP Phone.lnk
set exe_path=%public%\Downloads\Genesys SIP Phone\GenesysSIPPhone.exe
set policies_servicing_path=HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Servicing
set archive_path=�� ������

net session 1>nul 2>&1
if %errorlevel% neq 0 (
	echo ��� �ࠢ �����������
	pause
	exit
) 

:loop
	echo.
	echo.
	echo ==== ����⥭� ��⠭���� Genesys SIP Phone ====
	echo.
	call :checks
	echo.
	echo ������ ����� �㭪�:
	echo 0 - ��室
	echo 1 - ��⠭����� Genesys SIP Phone (2-5 �㭪��)
	echo 2 - ������� ��娢
	echo 2.1 -- 7-Zip
	echo 2.2 -- WinRAR
	echo 2.3 -- Powershell
	echo 3 - ������� ��娢
	echo 4 - ����� �᪫�祭�� �࠭������ ��� GenesysSIPPhone.exe
	echo 5 - ������� RUN.bat �� ��騩 ࠡ�稩 �⮫
	echo === .NET 3.5 ===
	echo 6 - ���������� ����⨪� � ��१���� wuauserv
	echo 7 - ��⠭���� �१ DISM
	echo.
	set option=undefined
	set /p option=��᫮:
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
	echo ��� �㭪�: %option%.
	goto :loop

:checks
	reg query "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.5" /v Install 1>nul 2>&1
	if %errorlevel% neq 0 (
		echo [x] .NET 3.5 - �� ��⠭�����
	) else (
		reg query "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.5" /v Install | findstr "0x1" 1>nul 2>&1
		if %errorlevel% neq 0 (
			echo [-] .NET 3.5 - ��⠭�����, �� �⪫�祭
		) else (
			echo [+] .NET 3.5 - ��⠭�����, ����祭
		)
	)

	if exist C:\Genesys_SIP_Phone.zip (
		set archive_path=C:\Genesys_SIP_Phone.zip
	)
	if exist "C:\Genesys SIP Phone.zip" (
		set archive_path=C:\Genesys SIP Phone.zip
	)
	echo ��娢 ��� �����筨�: %archive_path%
	if exist "%c_path%" (echo �����㦥�: %c_path%)
	if exist "%exe_path%" (echo �����㦥�: %exe_path%)
	if exist "%run_path%" (echo �����㦥�: %run_path%)
	if exist "%lnk_path%" (echo �����㦥�: %lnk_path%)
	goto :eof

:expand_archive
	if "%archive_path%"=="�� ������" (
		pause
		color 0c
		echo C:\Genesys_SIP_Phone.zip � C:\Genesys SIP Phone.zip �� �������. ��� ��娢� ��� �����祭��.
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
		echo WinRAR, 7-zip � powershell �� �������. ��� ���������� ������� ��娢.
		pause
		color 0F
		goto :eof
	)
	:_7zip_
		echo ��࠭: 7-Zip
		call :safely ^"%_7zip%^" x ^"%archive_path%^" -o^"%public%\Downloads\^" -y
		echo ��娢 �����祭
		goto :eof
	:winrar_
		echo ��࠭: WinRAR
		call :safely ^"%winrar%^" x -y ^"%archive_path%^" ^"%public%\Downloads\^"
		echo ��娢 �����祭
		goto :eof
	:powershell_expanding
		echo ��࠭: Powershell's Expand-Archive method
		call :safely powershell -ExecutionPolicy Bypass -Command ^"Expand-Archive -Path '%archive_path%' -DestinationPath '%public%\Downloads' -Force -ErrorAction Stop^"
		echo ��娢 �����祭
		goto :eof	

:remove_archive
	if "%archive_path%"==undefined (
		pause
		color 0c
		echo C:\Genesys_SIP_Phone.zip � C:\Genesys SIP Phone.zip �� �������. ��� ��娢� ��� 㤠�����.
		pause
		color 0F
		goto :loop
	)
	del /q "%archive_path%"
	goto :eof

:firewall_rules
	echo �஢�ઠ ������ �ࠢ��� �室��� ᮥ�������...
	netsh advfirewall firewall show rule name="Allow In - %exe_path%" 1>nul 2>&1
	if %errorlevel% neq 0 (
		call :safely netsh advfirewall firewall add rule name=^"Allow In - %exe_path%^" dir=in program=^"%exe_path%^" action=allow
		echo ��������� �ࠢ��� �室��� ᮥ�������
	)
	echo �஢�ઠ ������ �ࠢ��� ��室��� ᮥ�������...
	netsh advfirewall firewall show rule name="Allow Out - %exe_path%" 1>nul 2>&1
	if %errorlevel% neq 0 (
		call :safely netsh advfirewall firewall add rule name=^"Allow Out - %exe_path%^" dir=out program=^"%exe_path%^" action=allow
		echo ��������� �ࠢ��� ��室��� ᮥ�������
	)
	goto :eof

:runbat
	(
		echo @echo off
		echo cd ^"C:\Users\Public\Downloads\Genesys SIP Phone^"
		echo start "GSP" GenesysSIPPhone.exe -config genesys
	) > "%run_path%"
	echo RUN.bat ����ᠭ
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
	    echo [�訡��] �������: %*
	    echo ���: %errorlevel%
	    pause
	    color 0F
	    goto :loop
	)
	goto :eof
:end
pause