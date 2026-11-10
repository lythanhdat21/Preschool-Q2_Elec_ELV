@echo off
rem Giu cua so nay mo: moi lan luu Ton kho.xlsx thi data.js (du lieu website) tu cap nhat.
title Ton kho - tu cap nhat website
cd /d "%~dp0"
where py >nul 2>nul && (py -3 cap_nhat_du_lieu.py --watch) || (python cap_nhat_du_lieu.py --watch)
pause
