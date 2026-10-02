@echo off
echo === 1. ¿QUIÉN SOY? (Usuario y permisos) ===
whoami
echo.
echo === 2. ¿DÓNDE ESTÁ EJECUTANDO N8N? (Directorio de trabajo) ===
echo %cd%
echo.
echo === 3. ¿DÓNDE ESTÁ GUARDADO ESTE .BAT? ===
echo %~dp0
echo.
echo === 4. ¿QUÉ ARCHIVOS VEO EN LA CARPETA DEL .BAT? ===
dir "%~dp0"