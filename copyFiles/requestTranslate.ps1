# モジュール取り込み
Import-Module ((Split-path $PSScriptRoot) + "\modules\pathFinder.psm1") -Force;
Import-Module ((Split-path $PSScriptRoot) + "\modules\utility.psm1") -Force;
Add-Type -AssemblyName System.Windows.Forms;

<#
----------------------------------------------------
.SYNOPOSIS
    コピー先フォルダのパスを取得

.PARAMETER srcFolderName
    コピー元フォルダの名前
.PARAMETER destParentPath
    提出先フォルダが格納されている親フォルダのパス
----------------------------------------------------
#>
function getDestinationPath {
    Param(
        [string]$srcFolderName,
        [string]$destParentPath
    )

    # コピー元フォルダ名から管理番号とコロン削除
    [string]$movieTitle = ([regex]::Replace($srcFolderName, "^J[0-9]{4}_", "")) -replace "[ ： |：]";
    
    # 正規表現にマッチ And 動画タイトルを含むフォルダのフルパス取得
    # 変換に種類がある記号注意
    [string]$destFullPath = "";
    [string]$baseName = "";
    Get-ChildItem -Path $destParentPath -Directory | ForEach-Object {
        $baseName = ($_.BaseName -replace "[ ： |：|:]");
        if(($baseName -like "*" + $movieTitle + "*")) {
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
        movieName = "";
        translateFile = "";
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
                $returnHash.movieName = $_.BaseName;
                $highestMp4Revision = $cuMp4Revision;
            }
        }
    }

    # 翻訳ファイル(Excel)のフルパスを取得
    [int]$highestExcelRevision = 0;
    [int]$cuExcelRevision = 0;
    if (-not (Test-Path ($srcFolderPath + "\翻訳\"))) { return $returnHash; }
    Get-ChildItem -Path ($srcFolderPath + "\翻訳\") -Filter "*.xlsx" | ForEach-Object {
        if(($_.BaseName -match "^【翻訳】")) {
            $returnHash.translateFile = $_.FullName;
        }
    }

    return $returnHash;
}

<#
----------------------------------------------------
.SYNOPOSIS
    各言語フォルダが空の場合、翻訳ファイルをコピーする

.PARAMETER projPath
    提出先フォルダのフルパス
.PARAMETER excelFileName
    コピー元フォルダのフルパス
----------------------------------------------------
#>
function copyToLangFile {
    Param(
        [string]$projPath,
        [string]$excelFileName
    )
    
    # 翻訳ファイル格納先をループし言語フォルダ内が空なら翻訳ファイルコピー
    [string]$copyFileName = "";
    Get-ChildItem -Path $projPath -Directory | ForEach-Object {
        if ((Get-ChildItem $_.FullName | Measure-Object).Count -eq 0) {
            $copyFileName = [regex]::Replace($excelFileName, "^【.*】", ("【" + $_.Name + "】"));
            Copy-Item ($projPath + "\" + $excelFileName) ($projPath + "\" + $_.Name + "\" + $copyFileName);
        }
    }

}

<#
----------------------------------------------------
Main処理
- フォームでコピー元、コピー先のパスを選択
- 日本語MP4ファイルショートカット作成・翻訳文章コピー
----------------------------------------------------
#>

# フォルダ選択時の起点
[hashtable]$ini = Get-IniFile ($PSScriptRoot + ".\config.ini")
[string]$source = $ini["START_PATH"]["WORKSPACE"];
[string]$dest = $ini["START_PATH"]["TRANSLATE"];

# フォルダ選択フォーム表示
# Property : source, dest
[hashtable]$scriptConf = showTranslateRequestForm $source $dest;

# 入力値判定
[bool]$isContinue = $true;
foreach ($key in $scriptConf.Keys) {
    $confVal = $scriptConf[$key];
    if ($confVal -eq "") { $isContinue = $false; }
}
if (-not $isContinue) {
    [System.Windows.Forms.MessageBox]::Show(
        "パスが入力されていません。", 
        "翻訳依頼提出",
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
        [string]$destFullPath = getDestinationPath $_.Name $scriptConf.dest;

        # Property : mp4,movieName, translateFile
        [hashtable]$copyItems = getCopyFiles $_.FullName $scriptConf.lang;

        if(-not ($destFullPath -eq "") -and (Test-Path $destFullPath)) {
            Copy-Item $copyItems.translateFile $destFullPath;
            CreateShortcut $copyItems.mp4 "" ($destFullPath + "\" + $copyItems.movieName + ".lnk");
        } else {
            [System.Windows.Forms.MessageBox]::Show(
                ($_.Name) + "のコピーに失敗しました。`r`n`r`nコピー元またはコピー先にフォルダが存在しない可能性があります。", 
                "submitProducts",
                "OK",
                "Error"
            );
        }

        copyToLangFile $destFullPath (Split-Path $copyItems.translateFile -Leaf);
    }
}

# コピー元とコピー先を開く
Start-Process $scriptConf.dest;
Start-Process $scriptConf.source;


