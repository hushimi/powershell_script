# モジュール取り込み
Import-Module ((Split-path $PSScriptRoot) + "\modules\utility.psm1") -Force;
Add-Type -AssemblyName System.Windows.Forms;

<#
----------------------------------------------------
Main処理
- スクリプト実行dirにmediaフォルダがあれば削除
- 画像抽出するOfficeファイルを選択
- ファイルをスクリプト実行dirにコピー
- ファイルをzipに変換
- 7zipでzipをtemp dirに展開
- 展開したdirから各拡張子に基づいたパスのmediaフォルダを抽出
- 抽出したmediaフォルダを選択したファイルと同階層に設置
----------------------------------------------------
#>

# startExtract.batにて設定した変数を$PSScriptRootに設定
[string]$scriptPath = (cmd /c "echo %CUDIR%");

# media フォルダの削除
if (Test-Path ($PSScriptRoot + "\media\")) {
    Write-Host "Deleting old media folder..." -ForegroundColor Yellow;
    Remove-Item -Recurse -Force ($PSScriptRoot + "\media\");
    cls;
}

# 対象ファイルの選択
[hashtable]$ini = Get-IniFile ($PSScriptRoot + ".\config.ini")
[string]$basePath = $ini["START_PATH"]["WORKSPACE"];

$tgtFilePath = Get-FilePath $basePath "Office Files|*.docx;*.xlsx;*.pptx" $false;
if ($tgtFilePath -eq "") { exit; }
[string]$tgtFileName = [System.IO.Path]::GetFileNameWithoutExtension($tgtFilePath);
[string]$tgtExtension = [System.IO.Path]::GetExtension($tgtFilePath);
[string]$tgtParentPath = Split-Path $tgtFilePath -Parent;
Write-Host $tgtParentPath -ForegroundColor Green;

# ファイルをZIPとしてコピー・展開
Write-Host "選択したファイルをコピー中..." -ForegroundColor Yellow;
Copy-Item $tgtFilePath ($PSScriptRoot + "\" + $tgtFileName + ".zip");
Write-Host "コピー完了" -ForegroundColor Cyan;
New-Item -ItemType Directory ($PSScriptRoot + "\temp\");
Expand-Archive -Path ($PSScriptRoot + "\" + $tgtFileName + ".zip") -DestinationPath ($PSScriptRoot + "\temp\");

# 画像フォルダ抜き出し
[string]$extractPath = "";
switch ($tgtExtension) {
    ".pptx" {$extractPath = "ppt\media"}
    ".xlsx" {$extractPath = "xl\media"}
    ".docx" {$extractPath = "word\media"}
}
Write-Host "画像抜き出し中..." -ForegroundColor Yellow;
Copy-Item -Recurse ($PSScriptRoot + "\temp\" + $extractPath) -Destination $PSScriptRoot;
Write-Host "抜き出し完了" -ForegroundColor Cyan;

# 後始末
Remove-Item -Recurse ($PSScriptRoot + "\temp\"), ($PSScriptRoot + "\" + $tgtFileName + ".zip");
cls;
Move-Item ($scriptPath + "media") ($tgtParentPath + "\" + $tgtFileName + "(画像)");
Start-Process ($tgtParentPath);
