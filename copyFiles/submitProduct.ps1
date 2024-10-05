# モジュール取り込み
Import-Module ((Split-path $PSScriptRoot) + "\modules\pathFinder.psm1") -Force;
Import-Module ((Split-Path $PSScriptRoot) + "\modules\utility.psm1") -Force;
Add-Type -AssemblyName System.Windows.Forms;

<#
----------------------------------------------------
.SYNOPOSIS
    コピー先フォルダのパスを取得

.PARAMETER srcFolderName
    コピー元フォルダの名前
.PARAMETER destParentPath
    提出先フォルダが格納されている親フォルダのパス
.PARAMETER lang
    提出対象の言語
----------------------------------------------------
#>
function getDestinationPath {
    Param(
        [string]$srcFolderName,
        [string]$destParentPath,
        [string]$lang
    )

    # 言語ごとの正規表現設定
    [string]$regExp = "";
    switch ($lang) {
        "Japanese" {$regExp = "^J"}
        "English" {$regExp = "^E"}
        "Korea" {$regExp = "^K"}
        "China" {$regExp = "^C"}
        "Taiwan" {$regExp = "^T"}
    }

    # コピー元フォルダ名から管理番号とコロン削除
    [string]$movieTitle = ([regex]::Replace($srcFolderName, "^J[0-9]{4}_", "")) -replace "[ ： |：]";
    
    # 正規表現にマッチ And 動画タイトルを含むフォルダのフルパス取得
    # 変換に種類がある記号注意
    [string]$destFullPath = "";
    [string]$baseName = "";
    Get-ChildItem -Path $destParentPath -Directory | ForEach-Object {
        $baseName = ($_.BaseName -replace "[ ： |：|:]");
        if(($baseName -like "*" + $movieTitle + "*") -and ($_.BaseName -match $regExp)) {
            $destFullPath = $_.FullName;
        }
    }

    return $destFullPath;
}

<#
----------------------------------------------------
.SYNOPOSIS
    コピーするMP4ファイルと翻訳文章ファイルのパスを取得

.PARAMETER srcFolderPath
    コピー元フォルダのフルパス
.PARAMETER lang
    提出対象の言語
----------------------------------------------------
#>
function getCopyFiles {
    Param(
        [string]$srcFolderPath,
        [string]$lang
    )

    [hashtable]$returnHash = @{
        mp4 = "";
        translateFile = "";
    }

    # 言語ごとの正規表現設定
    [string]$regExp = "";
    [string]$translateStartPath = "";
    switch ($lang) {
        "Japanese" {$regExp = "^J"; $translateStartPath = "\翻訳"}
        "English" {$regExp = "^E"; $translateStartPath = "\翻訳\English\"}
        "Korea" {$regExp = "^K"; $translateStartPath = "\翻訳\Korea\"}
        "China" {$regExp = "^C"; $translateStartPath = "\翻訳\China\"}
        "Taiwan" {$regExp = "^T"; $translateStartPath = "\翻訳\Taiwan\"}
    }

    [int]$highestMp4Revision = 0;
    [int]$cuMp4Revision = 0;
    # 一番リビジョンの高いMP4ファイルフルパスを取得
    Get-ChildItem -Path $srcFolderPath -File -Filter "*.mp4" | ForEach-Object {
        if(($_.BaseName -match $regExp + ".*_r[0-9]{1,2}")) {
            # 一番高いリビジョンの判定
            $cuMp4Revision = [int]($_.BaseName -replace ($regExp + ".*_r"));
            if($cuMp4Revision -gt $highestMp4Revision) {
                $returnHash.mp4 = $_.FullName;
                $highestMp4Revision = $cuMp4Revision;
            }
        }
    }

    # 一番リビジョンの高い翻訳ファイル(Excel)のフルパスを取得
    [int]$highestExcelRevision = 0;
    [int]$cuExcelRevision = 0;
    if (-not (Test-Path ($srcFolderPath + $translateStartPath))) { return $returnHash; }
    Get-ChildItem -Path ($srcFolderPath + $translateStartPath) -Filter "*.xlsx" | ForEach-Object {
        # 日本語の場合
        if(($lang -eq "Japanese") -and ($_.BaseName -match "^【翻訳】")) {
            $returnHash.translateFile = $_.FullName;

        # 多言語の場合
        } else {
            # 一番高いリビジョンの判定
            $cuExcelRevision = [int]($_.BaseName -replace "^.*_r");
            if($cuExcelRevision -gt $highestExcelRevision) {
                $returnHash.translateFile = $_.FullName;
                $highestExcelRevision = $cuExcelRevision;
            }
        }
    }

    return $returnHash;
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
[string]$source = $ini["START_PATH"]["WORKSPACE"];
[string]$dest = $ini["START_PATH"]["DELIVERY"];

# フォルダ選択フォーム表示(pathFinder.psm1)
# Property : menu, destination path, source path
[hashtable]$scriptConf = showSubmitForm @("Japanese", "English", "Korea", "China", "Taiwan") $dest $source;


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

# ファイルコピー
[string]$destFullPath = "";
Get-ChildItem -Path $scriptConf.source -Directory | ForEach-Object {
    
    # 管理番号のあるフォルダがコピー操作対象
    if($_.Name -match "^J[0-9]{4}_") {
        # コピー先のパス取得
        [string]$destFullPath = getDestinationPath $_.Name $scriptConf.dest $scriptConf.lang;

        # Property : mp4,translateFile
        [hashtable]$copyItems = getCopyFiles $_.FullName $scriptConf.lang;

        if(-not ($destFullPath -eq "") -and (Test-Path $destFullPath)) {
            Copy-Item $copyItems.mp4 $destFullPath;
            Copy-Item $copyItems.translateFile $destFullPath;
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


