@echo off
setlocal EnableExtensions EnableDelayedExpansion
title Diff Excel

rem Garde la fenetre ouverte au double-clic (Explorer lance cmd /c).
if /I not "%~1"=="GO" (
  cmd /k "%~f0" GO
  exit /b 0
)

cd /d "%~dp0"
echo ----------------------------------------
echo Diff Excel
echo Dossier : %cd%
echo ----------------------------------------
echo.

call :resolve_node
if errorlevel 1 goto INSTALL_NODE
goto HAVE_NODE

:INSTALL_NODE
echo Node.js est introuvable. Telechargement de la LTS portable, sans droits admin...
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install-node-portable.ps1" "%~dp0node-portable"
if errorlevel 1 goto INSTALL_FAIL
call :resolve_node
if errorlevel 1 goto INSTALL_FAIL
goto HAVE_NODE

:INSTALL_FAIL
echo.
echo Echec du telechargement automatique.
echo Sans admin : telechargez le zip Windows x64 LTS sur https://nodejs.org
echo Dezippez-le dans node-portable a cote de lancer.bat, puis relancez.
echo Ou definissez NODE_PORTABLE vers le dossier qui contient node.exe.
echo.
goto END

:HAVE_NODE
echo Node.js utilise :
where node
node -v
echo.

if exist "node_modules\" goto START_SERVER
echo Installation des dependances...
call npm.cmd install
if errorlevel 1 (
  echo Echec de l'installation des dependances.
  goto END
)
echo.

:START_SERVER
netstat -ano | findstr ":5173" | findstr "LISTENING" >nul
if errorlevel 1 goto RUN_VITE
echo Le serveur tourne deja. Ouverture du navigateur...
start "" "http://127.0.0.1:5173/"
echo.
echo Application : http://127.0.0.1:5173/
echo Fermez l'autre fenetre Diff Excel pour arreter le serveur.
echo.
goto END

:RUN_VITE
echo Lancement de Diff Excel...
echo Navigateur : http://127.0.0.1:5173/
echo Fermez cette fenetre pour arreter le serveur.
echo.
start "" /min cmd /c "timeout /t 4 /nobreak >nul & start http://127.0.0.1:5173/"
call npm.cmd start
if errorlevel 1 echo Le serveur n'a pas demarre. Voir le message ci-dessus.
echo.
goto END

:resolve_node
if defined NODE_PORTABLE (
  if exist "%NODE_PORTABLE%\node.exe" (
    set "PATH=%NODE_PORTABLE%;!PATH!"
    echo Node portable : %NODE_PORTABLE%
    exit /b 0
  )
  echo NODE_PORTABLE est defini mais node.exe est introuvable dans :
  echo   %NODE_PORTABLE%
  echo.
)
if exist "%~dp0node-portable\node.exe" (
  set "PATH=%~dp0node-portable;!PATH!"
  echo Node portable : %~dp0node-portable
  exit /b 0
)
if exist "%~dp0node-portable\" (
  for /d %%D in ("%~dp0node-portable\*") do (
    if exist "%%D\node.exe" (
      set "PATH=%%D;!PATH!"
      echo Node portable : %%D
      exit /b 0
    )
  )
)
where node >nul 2>&1
if errorlevel 1 exit /b 1
exit /b 0

:END
echo.
echo ----------------------------------------
echo Fenetre volontairement laissee ouverte.
echo Tapez exit ou fermez la fenetre.
echo ----------------------------------------