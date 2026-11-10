@echo off
title Enregistrements Controle d entree
cd /d "%~dp0"
rem Ouvre l application via un petit serveur local (http://localhost), sans passer par internet.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0outils\serveur-local.ps1"
if errorlevel 1 pause
