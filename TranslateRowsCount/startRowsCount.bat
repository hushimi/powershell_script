@echo off
SET CUDIR=%~dp0
pushd %~dp0
color 0a
cls

powershell -ExecutionPolicy Unrestricted -File ".\TranslateRowsCount.ps1"

popd
pause