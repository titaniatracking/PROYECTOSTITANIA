@echo off
title Sistema de odometros
cd /d "%~dp0"

start "Odometros recepcion" /min python web_mantenimiento_odometros.py
start "Odometros distancia" /min python web_mantenimiento_distancia.py
start "Menu odometros" /min python menu_odometros.py

timeout /t 2 /nobreak >nul
start "" http://127.0.0.1:8062
exit

