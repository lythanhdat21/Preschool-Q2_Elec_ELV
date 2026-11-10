@echo off
chcp 65001 >nul
rem Chay website tren may nay (va cho may khac cung mang LAN xem).
rem Giu cua so nay mo trong luc xem. Dong cua so = tat website.
cd /d "%~dp0.."
start "" "http://localhost:8000/Script/"
where python >nul 2>nul && (python -m http.server 8000) || (py -m http.server 8000)
pause
