@echo off
SET CUDIR=%~dp0
pushd %~dp0
color 0a
cls

powershell -ExecutionPolicy Unrestricted -File ".\mouseMove.ps1"

popd
pause