@echo off
chcp 65001 >nul
set "F=%~dp0거지키우기_k.html"
if not exist "%F%" (
  echo 같은 폴더에 거지키우기_k.html 파일이 없어요.
  pause
  exit /b
)
set "U=file:///%F:\=/%#mini"
set "OPT=--window-size=262,460 --window-position=1500,560"
reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe" >nul 2>&1 && goto chrome
reg query "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe" >nul 2>&1 && goto chrome
start "" msedge --app="%U%" %OPT%
exit /b
:chrome
start "" chrome --app="%U%" %OPT%
exit /b
