@echo off
chcp 65001 >nul
title Gerar RPA Download DCTF.exe
cd /d "%~dp0"
echo.
echo  ==============================================
echo   RPA Download DCTF - gerando o executavel
echo  ==============================================
echo.

set "PY="
where py >nul 2>nul && set "PY=py -3"
if not defined PY (
  where python >nul 2>nul && set "PY=python"
)
if not defined PY (
  echo  Python nao encontrado.
  echo  Instale em https://www.python.org/downloads/ marcando "Add python.exe to PATH".
  pause
  exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
  echo  Criando ambiente isolado...
  %PY% -m venv .venv || goto erro
)
call ".venv\Scripts\activate.bat" || goto erro

python -c "import tkinter" >nul 2>nul
if errorlevel 1 (
  echo.
  echo  O seu Python foi instalado sem o componente Tcl/Tk, que desenha a janela do robo.
  echo  Rode o instalador do Python de novo, escolha "Modify" e marque "tcl/tk and IDLE".
  echo  Depois apague a pasta .venv e rode este arquivo outra vez.
  pause
  exit /b 1
)

echo  Instalando dependencias (pode levar alguns minutos na primeira vez)...
python -m pip install --upgrade pip >nul
python -m pip install -r requirements.txt pyinstaller==6.22.3 || goto erro

echo  Empacotando...
python -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name "RPA Download DCTF" ^
  --icon "assets\icone.ico" ^
  --add-data "assets;assets" ^
  --collect-all playwright ^
  rpa_download_dctf.py || goto erro

echo.
echo  Pronto! O programa esta em:
echo  %cd%\dist\RPA Download DCTF.exe
echo.
explorer dist
pause
exit /b 0

:erro
echo.
echo  Algo deu errado. Tire um print desta janela e envie para analise.
pause
exit /b 1
