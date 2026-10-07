@echo off
setlocal EnableExtensions
title Consolidador Excel

rem Vai para a pasta do .bat (pushd tambem funciona com caminhos de rede \\servidor\...)
pushd "%~dp0"

set "SCRIPT=Executor_Consolizador_Excel_ATUALIZADO.py"
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"

if not exist "%SCRIPT%" (
    echo [ERRO] Arquivo "%SCRIPT%" nao encontrado em:
    echo        %CD%
    goto :fim_erro
)

rem Localiza um Python funcional: primeiro o launcher "py", depois "python"
set "PY="
py -3 -c "import sys" >nul 2>&1
if not errorlevel 1 set "PY=py -3"
if not defined PY (
    python -c "import sys" >nul 2>&1
    if not errorlevel 1 set "PY=python"
)
if not defined PY (
    echo [ERRO] Nenhum Python encontrado.
    echo        Instale o Python 3.13 em https://www.python.org/downloads/windows/
    echo        marcando a opcao "Add python.exe to PATH".
    goto :fim_erro
)

for /f "delims=" %%i in ('%PY% -c "import sys; print(sys.executable)"') do set "PYEXE=%%i"
echo Python em uso: %PYEXE%
echo.

rem Verifica as dependencias do script
%PY% -c "import pandas, openpyxl, xlrd" >nul 2>&1
if errorlevel 1 (
    echo [AVISO] Faltam dependencias: pandas, openpyxl e/ou xlrd.
    choice /c SN /m "Instalar agora neste Python"
    if errorlevel 2 goto :fim_erro
    %PY% -m pip install pandas openpyxl "xlrd>=2.0.1"
    if errorlevel 1 (
        echo [ERRO] Falha na instalacao das dependencias.
        goto :fim_erro
    )
    echo.
)

%PY% "%SCRIPT%"
set "RC=%ERRORLEVEL%"
echo.

if not "%RC%"=="0" (
    echo [ERRO] O script terminou com codigo %RC%.
    goto :fim_erro
)

popd
pause
exit /b 0

:fim_erro
popd
pause
exit /b 1
