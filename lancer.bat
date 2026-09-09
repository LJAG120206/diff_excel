@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

where node >nul 2>&1
if errorlevel 1 (
  echo Node.js est introuvable.
  echo Installez-le depuis https://nodejs.org puis relancez ce script.
  pause
  exit /b 1
)

if not exist "node_modules\" (
  echo Installation des dependances...
  call npm install
  if errorlevel 1 (
    echo Echec de l'installation.
    pause
    exit /b 1
  )
  echo.
)

echo Lancement de Diff Excel...
echo Fermez cette fenetre pour arreter le serveur.
echo.
call npm run dev -- --open
if errorlevel 1 pause
