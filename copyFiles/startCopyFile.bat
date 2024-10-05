@echo off
SET CUDIR=%~dp0
pushd %~dp0
color 0a
cls

echo メニュー
echo ------------------------------
echo 1.動画提出
echo 2.翻訳依頼
echo 3.翻訳文章取得
echo 4.Cancel
echo ------------------------------
set /p topmenu="メニューを選択してください : "

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