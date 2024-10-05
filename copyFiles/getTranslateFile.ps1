# モジュール取り込み
Import-Module ((Split-path $PSScriptRoot) + "\modules\pathFinder.psm1") -Force;
Import-Module ((Split-path $PSScriptRoot) + "\modules\utility.psm1") -Force;
Add-Type -AssemblyName System.Windows.Forms;

<#
----------------------------------------------------
.SYNOPOSIS
    翻訳依頼フォルダからコピー対象のExcelパスを取得

.PARAMETER targetName
    コピー先フォルダの名前
.PARAMETER sourcePath
    翻訳依頼フォルダのパス
.PARAMETER lang
    提出対象の言語
----------------------------------------------------
#>
function getSourcePath {
    Param(
        [string]$targetName,
        [string]$sourcePath,
        [string]$lang
    )

    # 翻訳文取得対象のフォルダ名設定
    [string]$targetLang = "";
    switch ($lang) {
        "English" {$targetLang = "英語"}
        "Korea" {$targetLang = "韓国語"}
        "China" {$targetLang = "簡体字"}
        "Taiwan" {$targetLang = "繁体字"}
    }

    # コピー元フォルダ名から管理番号とコロン削除
    [string]$movieTitle = ([regex]::Replace($targetName, "^J[0-9]{4}_", "")) -replace "[ ： |：]";
    
    # 動画タイトルを含むフォルダからフルパスを取得
    [string]$translateFolderPath = "";
    [string]$translateFilePath = "";
    [string]$baseName = "";
    Get-ChildItem -Path $sourcePath -Directory | ForEach-Object {
        $baseName = ($_.BaseName -replace "[ ： |：|:]");
        if($baseName -like "*" + $movieTitle + "*") {
            $translateFolderPath = ($_.FullName + "\" + $targetLang + "\");
        }
    }

    # コピー元フォルダから翻訳Excelファイルのパスを取得
    if (-not (Test-Path $translateFolderPath)) { return ""; }
    [int]$xlsxFileCount = ((Get-ChildItem -Path $translateFolderPath -Filter "*.xlsx" | Measure-Object).Count);
    if($xlsxFileCount -gt 1) {
        $translateFilePath = Get-FilePath $translateFolderPath "Excel Worksheets|*.xlsx" $false;

    } else {
        $translateFilePath = (Get-ChildItem -Path $translateFolderPath -Filter "*.xlsx" | Select-Object -First 1).FullName;
    }

    return $translateFilePath;
}

<#
----------------------------------------------------
Main処理
- フォームでコピー元、コピー先のパス・対象言語を選択
- MP4ファイル・翻訳文章のコピーを実行
----------------------------------------------------
#>
# フォルダ選択時の起点
[hashtable]$ini = Get-IniFile ($PSScriptRoot + ".\config.ini")
[string]$dest = $ini["START_PATH"]["WORKSPACE"];
[string]$source = $ini["START_PATH"]["TRANSLATE"];

# フォルダ選択フォーム表示(pathFinder.psm1)
# Property : menu, destination path, source path
[hashtable]$scriptConf = showSubmitForm @("English", "Korea", "China", "Taiwan") $dest $source;

# 入力値判定
[bool]$isContinue = $true;
foreach ($key in $scriptConf.Keys) {
    $confVal = $scriptConf[$key];
    if ($confVal -eq "") { $isContinue = $false; }
}
if (-not $isContinue) {
    [System.Windows.Forms.MessageBox]::Show(
        "パスと対象言語のいずれかが入力されていません。", 
        "submitProducts",
        "OK",
        "Error"
    );
    exit;
}

# ファイルコピー処理
[string]$destFullPath = "";
Get-ChildItem -Path $scriptConf.dest -Directory | ForEach-Object {
    
    # 管理番号のあるフォルダがコピー操作対象
    if($_.Name -match "^J[0-9]{4}_") {
        # コピーする翻訳文章のパスを取得
        [string]$copyItemPath = getSourcePath $_.Name $scriptConf.source $scriptConf.lang;

        # コピー先のパス設定
        $destFullPath = ($_.FullName + "\翻訳\" + $scriptConf.lang + "\");

        # 一番高い翻訳ファイルのリビジョン取得
        [int]$highestExcelRevision = 0;
        [int]$cuExcelRevision = 0;
        Get-ChildItem -Path ($destFullPath) -Filter "*.xlsx" | ForEach-Object {
            # 一番高いリビジョンの判定
            $cuExcelRevision = [int]($_.BaseName -replace "^.*_r");
            if($cuExcelRevision -gt $highestExcelRevision) {
                $highestExcelRevision = $cuExcelRevision;
            }
        }

        # ファイルコピー実行
        if(-not ($copyItemPath -eq "") -and (Test-Path $destFullPath)) {
            # コピー元ファイル一つ目のコピー
            [string]$newFileName = Split-Path $copyItemPath -Leaf;
            $newFileName = $newFileName -replace "【.*】", "【ATC】";
            $newFileName = $newFileName -replace "_r\d{1,2}", "";
            $newFileName = ($newFileName -replace "\.xlsx", "") + "_" `
                + ($scriptConf.lang.Substring(0,1)) + "_r" `
                + ($highestExcelRevision + 1)`
            ;
            Copy-Item $copyItemPath ($destFullPath + "\" + $newFileName + ".xlsx");

            # コピー元ファイル二つ目のコピー
            $newFileName = $newFileName -replace "_r[1-9]{1,2}", "";
            $newFileName += "_r" + ($highestExcelRevision + 2);
            Copy-Item $copyItemPath ($destFullPath + "\" + $newFileName + ".xlsx");

        } else {
            [System.Windows.Forms.MessageBox]::Show(
                ($_.Name) + "のコピーに失敗しました。`r`n`r`nコピー元またはコピー先にフォルダが存在しない可能性があります。", 
                "submitProducts",
                "OK",
                "Error"
            );
        }
    }
}

# コピー元とコピー先を開く
Start-Process $scriptConf.dest;
Start-Process $scriptConf.source;


