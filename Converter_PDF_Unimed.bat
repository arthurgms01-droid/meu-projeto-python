@echo off
setlocal
chcp 65001 >nul
title Conversor PDF Unimed para Excel

rem Pasta onde este .bat esta (o .py deve ficar na mesma pasta)
cd /d "%~dp0"
set "SCRIPT=%~dp0pdf_unimed_para_excel.py"

if not exist "%SCRIPT%" (
    echo ERRO: arquivo pdf_unimed_para_excel.py nao encontrado em:
    echo %~dp0
    echo Coloque o .bat e o .py na mesma pasta.
    pause
    exit /b 1
)

rem Localiza o Python: primeiro o launcher "py", depois "python"
set "PY="
where py >nul 2>&1 && set "PY=py -3"
if not defined PY (
    where python >nul 2>&1 && set "PY=python"
)
if not defined PY (
    echo ERRO: Python nao encontrado.
    echo Instale em https://www.python.org/downloads/ e marque "Add Python to PATH".
    pause
    exit /b 1
)

rem Instala as bibliotecas na primeira execucao, se faltarem
%PY% -c "import pdfplumber, openpyxl" >nul 2>&1
if errorlevel 1 (
    echo Instalando bibliotecas necessarias: pdfplumber e openpyxl...
    %PY% -m pip install --user pdfplumber openpyxl
    if errorlevel 1 (
        echo ERRO: nao foi possivel instalar as bibliotecas.
        pause
        exit /b 1
    )
)

rem Abre o programa. PDFs ou pastas arrastados sobre o .bat ja entram na lista.
%PY% "%SCRIPT%" %*
if errorlevel 1 (
    echo.
    echo O programa terminou com erro. Veja a mensagem acima.
    pause
)

endlocal
