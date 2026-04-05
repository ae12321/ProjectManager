@echo off
cd /d "%~dp0"
powershell -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~n0.ps1"

