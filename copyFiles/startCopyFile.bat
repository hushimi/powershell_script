@echo off
SET CUDIR=%~dp0
pushd %~dp0
color 0a
cls

echo ���j���[
echo ------------------------------
echo 1.�����o
echo 2.�|��˗�
echo 3.�|�󕶏͎擾
echo 4.Cancel
echo ------------------------------
set /p topmenu="���j���[��I�����Ă������� : "

if /i %topmenu%==1 (
  powershell -WindowStyle Hidden -ExecutionPolicy Unrestricted -File ".\submitProduct.ps1"
) else if /i %topmenu%==2 (
  powershell -WindowStyle Hidden -ExecutionPolicy Unrestricted -File ".\requestTranslate.ps1"
) else if /i %topmenu%==3 (
  powershell -WindowStyle Hidden -ExecutionPolicy Unrestricted -File ".\getTranslateFile.ps1"
) else (
  popd
  goto END
)

popd