@echo off
SET SCRIPT=%~dp0src\xlsx2json.py


python %SCRIPT% %1 %2

pause