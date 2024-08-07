@echo off

REM Verificar se o processo já está em execução
tasklist /FI "IMAGENAME eq python.exe" 2>NUL | find /I /N "bot.py">NUL
if "%ERRORLEVEL%"=="0" (
    echo Já existe um processo em execução.
    pause
) else (
    echo Iniciando o processo...
    start python C:\Meus\Projetos\Python\PortalTransparenciaAposentados-Pencionistas\PTAP\PTAP\bot.py
)

