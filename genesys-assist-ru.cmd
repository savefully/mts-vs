@echo off
chcp 866 >nul

set "download_domain=none"
set "archive_path=не найден"
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
	echo Нет прав администратора
	pause
	exit
)

if "%1"=="" (
	echo.
	echo [Инфо] Домен скачивания не передан
) else (
	echo Домен скачивания: %~1
	set "download_domain=%~1"
)
pause
:loop
	echo.
	echo.
	echo ++++++++ Ассистент установки Genesys SIP Phone ++++++++
	echo.
	call :checks
	echo.
	echo Введите номер пункта:
	echo 0 - Выход
	echo 1 - Установить Genesys SIP Phone (2-6 пункты)
	echo 2 = Скачать архив
	echo 3 = Извлечь архив
	echo 3.1 -- 7-Zip
	echo 3.2 -- WinRAR
	echo 3.3 -- Powershell
	echo 4 - Удалить архив
	echo 5 - Внести исключения брандмауэра для GenesysSIPPhone.exe
	echo 6 - Записать RUN.bat на общий рабочий стол
	echo #### .NET 3.5 ####
	echo 7 - Обновление политики и перезапуск wuauserv
	echo 8 - Установка через DISM
	echo.
	set "option=Пусто"
	set /p "option=Число:"
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
	echo Нет пункта: %option%.
	goto :loop

:checks
	reg query "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.5" /v Install 1>nul 2>&1
	if %errorlevel% neq 0 (
		echo [x] .NET 3.5 - не установлен
	) else (
		reg query "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.5" /v Install | findstr "0x1" 1>nul 2>&1
		if %errorlevel% neq 0 (
			echo [-] .NET 3.5 - установлен, но отключен
		) else (
			echo [+] .NET 3.5 - установлен, включен
		)
	)
	set "archive_path=не найден"
	if exist C:\Genesys_SIP_Phone.zip (
		set "archive_path=C:\Genesys_SIP_Phone.zip"
	)
	if exist "C:\Genesys SIP Phone.zip" (
		set "archive_path=C:\Genesys SIP Phone.zip"
	)
	echo Архив для извлечния: %archive_path%
	if exist "%c_path%" (echo Обнаружен: %c_path%)
	if exist "%exe_path%" (echo Обнаружен: %exe_path%)
	if exist "%run_path%" (echo Обнаружен: %run_path%)
	if exist "%lnk_path%" (echo Обнаружен: %lnk_path%)
	goto :eof

:download_archive
	echo.
	if "%download_domain%"=="none" (
		set /p "download_domain=Введите домен, который должен стоять вместо * ^> soft.*.ru:"
	)
	call :safely curl -o ^"C:\Genesys_SIP_Phone.zip^" ^"https://soft.%download_domain%.ru/Genesys_SIP_Phone.zip^" || exit /b 1
	set "archive_path=C:\Genesys_SIP_Phone.zip"
	echo Архив скачан
	goto :eof

:expand_archive
	echo.
	if "%archive_path%"=="не найден" (
		color 0c
		echo C:\Genesys_SIP_Phone.zip и C:\Genesys SIP Phone.zip не найдены. Нет архива для извлечения.
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
		echo WinRAR, 7-zip и powershell не найдены. Нет возможности извлечь архив.
		pause
		color 0F
		exit /b 1
	)
	:_7zip_
		echo.
		echo Выбран: 7-Zip
		call :safely ^"%_7zip%^" x ^"%archive_path%^" -o^"%public%\Downloads\^" -y
		if %errorlevel% equ 0 (
			echo Архив извлечен
			goto :eof
		)
		echo Не удалось извлечь архив из-за ошибки
		exit /b 1
	:winrar_
		echo.
		echo Выбран: WinRAR
		call :safely ^"%winrar%^" x -y ^"%archive_path%^" ^"%public%\Downloads\^"
		if %errorlevel% equ 0 (
			echo Архив извлечен
			goto :eof
		)
		echo Не удалось извлечь архив из-за ошибки
		exit /b 1
	:powershell_expanding
		echo.
		echo Выбран: Powershell's Expand-Archive
		call :safely powershell -ExecutionPolicy Bypass -Command ^"Expand-Archive -Path '%archive_path%' -DestinationPath '%public%\Downloads' -Force -ErrorAction Stop^"
		if %errorlevel% equ 0 (
			echo Архив извлечен
			goto :eof
		)
		echo Не удалось извлечь архив из-за ошибки
		exit /b 1

:remove_archive
	echo.
	if "%archive_path%"=="не найден" (
		color 0c
		echo [Инфо] C:\Genesys_SIP_Phone.zip и C:\Genesys SIP Phone.zip не найдены. Нет архива для удаления.
		pause
		color 0F
	) else (
		del /q "%archive_path%"
		echo Архив удален
	)
	goto :eof

:firewall_rules
	echo.
	echo Проверка наличия правила входящих соединений...
	netsh advfirewall firewall show rule name="Allow In - %exe_path%" 1>nul 2>&1
	if %errorlevel% neq 0 (
		netsh advfirewall firewall add rule name=^"Allow In - %exe_path%^" dir=in program=^"%exe_path%^" action=allow 1>nul 2>&1
		if %errorlevel% neq 0 (
			echo Добавлено правило входящих соединений
		) else (
			echo [Предупреждение] Не удалось добавить правило входящих соединений
		)
	)
	echo Проверка наличия правила исходящих соединений...
	netsh advfirewall firewall show rule name="Allow Out - %exe_path%" 1>nul 2>&1
	if %errorlevel% neq 0 (
		netsh advfirewall firewall add rule name=^"Allow Out - %exe_path%^" dir=out program=^"%exe_path%^" action=allow 1>nul 2>&1
		if %errorlevel% neq 0 (
			echo Добавлено правило исходящих соединений
		) else (
			echo [Предупреждение] Не удалось добавить правило исходящих соединений
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
	echo RUN.bat записан
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
	    echo [Ошибка] Команда: %*
	    echo Код: %errorlevel%
	    pause
	    color 0F
	    exit /b 1
	)
	goto :eof
:end
pause