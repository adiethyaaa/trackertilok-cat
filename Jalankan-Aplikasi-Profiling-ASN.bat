@echo off
title Profiling ASN - Penjadwalan Ujian
cd /d "%~dp0"
echo ========================================================
echo   Menjalankan Aplikasi Profiling ASN...
echo ========================================================
start "" http://localhost:3000
node server.js
pause

