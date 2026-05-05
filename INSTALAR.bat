@echo off
echo ============================================
echo  TAF-EB -- Sistema de Avaliacao Fisica EB
echo ============================================
echo.
echo Verificando Node.js...
node --version >nul 2>&1
if errorlevel 1 (
    echo ERRO: Node.js nao encontrado.
    echo Baixe em: https://nodejs.org
    pause
    exit /b 1
)
echo Instalando dependencias...
npm install
echo.
echo Instalando Electron globalmente...
npm install -g electron
echo.
echo ============================================
echo  Instalacao concluida! Execute EXECUTAR.bat
echo ============================================
pause
