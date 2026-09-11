@echo off
title Build - Conversor XLSX para XLS
color 0A
echo.
echo  ============================================
echo   Build: Conversor XLSX ^> XLS
echo  ============================================
echo.

:: Verifica se Python esta instalado
python --version >nul 2>&1
if errorlevel 1 (
    echo  [ERRO] Python nao encontrado.
    echo  Baixe em: https://www.python.org/downloads/
    echo  Marque a opcao "Add Python to PATH" na instalacao.
    pause
    exit /b 1
)

echo  [1/3] Instalando dependencias...
pip install pyinstaller pandas openpyxl xlwt --quiet
if errorlevel 1 (
    echo  [ERRO] Falha ao instalar dependencias.
    pause
    exit /b 1
)

echo  [2/3] Gerando executavel...
pyinstaller --onefile --windowed --name "Conversor_XLSX_XLS" conversor_app.py
if errorlevel 1 (
    echo  [ERRO] Falha no build.
    pause
    exit /b 1
)

echo  [3/3] Concluido!
echo.
echo  Executavel gerado em: dist\Conversor_XLSX_XLS.exe
echo.
pause
