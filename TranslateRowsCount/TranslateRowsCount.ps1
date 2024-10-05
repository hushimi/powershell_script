# モジュール取り込み
Import-Module ((Split-path $PSScriptRoot) + "\modules\utility.psm1") -Force;
Add-Type -AssemblyName System.Windows.Forms;

<#
----------------------------------------------------
.SYNOPOSIS
    指定されたフォルダ内から翻訳依頼ファイルを探す

.PARAMETER $searchDir
    翻訳ファイルを格納しているフォルダのパス
----------------------------------------------------
#>
Function Get-TranslateFile {
    Param(
        [string]$searchDir
    )
    
    [string]$translateFilePath = "";
    Get-ChildItem -Path $searchDir | Where-Object { $_.Name -match "^【翻訳】.*\.xlsx$" } | ForEach-Object {
        if (Test-Path $_.FullName) {
            $translateFilePath = $_.FullName;
        }
    }

    return $translateFilePath;
}

<#
----------------------------------------------------
.SYNOPOSIS
    翻訳依頼用ファイルを開いて行数を数える

.PARAMETER $tgtExcelPath
    翻訳ファイルのパス
----------------------------------------------------
#>
Function CountTgtFileRows {
    Param(
        [string]$tgtExcelPath
    )
    # excel設定
    [object]$excel = New-Object -ComObject Excel.Application;
    $excel.Visible = $false;
    $excel.DisplayAlerts = $false;

    # book取得
    [object]$book = $excel.Workbooks.Open($tgtExcelPath);

    # Sheet取得
    [object]$sheet;
    $book.Sheets | ForEach-Object {
        if ($_.Name -match '.*動画Text.*') {
            $sheet = $book.Sheets.Item($_.Name);
        }
    }

    # 行数カウント(C列の文字が入っている行数を数える)
    [int]$rowsCount = 0;
    [int]$lastRow = $sheet.Cells($sheet.Rows.count, 3).End([Microsoft.Office.Interop.Excel.XlDirection]::xlUp.value__).Row;
    for ($i = 4; $i -le $lastRow; $i++) {
        if (-not ([string]::IsNullOrEmpty($sheet.Cells.Item($i, 3).value()))) {
            $rowsCount += 1;
        }
    }
    
    $book.Close();
    $excel.Quit();

    return $rowsCount.toString();
}

<#
----------------------------------------------------
Main処理
- TEL-Equipment以下の作業フォルダを選択
- 選択したフォルダ内をループして翻訳ファイルを見つける
- 各翻訳ファイルの翻訳文章行数を数えてコンソールに出力
----------------------------------------------------
#>
# 対象ファイルの選択
[hashtable]$ini = Get-IniFile ($PSScriptRoot + ".\config.ini")
[string]$basePath = $ini["START_PATH"]["WORKSPACE"];
[string]$tgtFolderPath = Get-Folder $basePath;
if ($tgtFolderPath -eq "") { exit; }

[string]$tgtFilePath = "";
[string]$rowsCount = "";
Get-ChildItem -Path $tgtFolderPath -Directory | ForEach-Object {
    if (Test-Path ($_.FullName + "\翻訳\")) {
        $tgtFilePath = Get-TranslateFile ($_.FullName + "\翻訳");
        $rowsCount = CountTgtFileRows $tgtFilePath;

        Write-Host (Split-Path $tgtFilePath -Leaf);
        Write-Host ($rowsCount.ToString() + "行") -ForegroundColor Cyan;
    }
}