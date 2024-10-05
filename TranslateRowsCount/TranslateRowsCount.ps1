# ���W���[����荞��
Import-Module ((Split-path $PSScriptRoot) + "\modules\utility.psm1") -Force;
Add-Type -AssemblyName System.Windows.Forms;

<#
----------------------------------------------------
.SYNOPOSIS
    �w�肳�ꂽ�t�H���_������|��˗��t�@�C����T��

.PARAMETER $searchDir
    �|��t�@�C�����i�[���Ă���t�H���_�̃p�X
----------------------------------------------------
#>
Function Get-TranslateFile {
    Param(
        [string]$searchDir
    )
    
    [string]$translateFilePath = "";
    Get-ChildItem -Path $searchDir | Where-Object { $_.Name -match "^�y�|��z.*\.xlsx$" } | ForEach-Object {
        if (Test-Path $_.FullName) {
            $translateFilePath = $_.FullName;
        }
    }

    return $translateFilePath;
}

<#
----------------------------------------------------
.SYNOPOSIS
    �|��˗��p�t�@�C�����J���čs���𐔂���

.PARAMETER $tgtExcelPath
    �|��t�@�C���̃p�X
----------------------------------------------------
#>
Function CountTgtFileRows {
    Param(
        [string]$tgtExcelPath
    )
    # excel�ݒ�
    [object]$excel = New-Object -ComObject Excel.Application;
    $excel.Visible = $false;
    $excel.DisplayAlerts = $false;

    # book�擾
    [object]$book = $excel.Workbooks.Open($tgtExcelPath);

    # Sheet�擾
    [object]$sheet;
    $book.Sheets | ForEach-Object {
        if ($_.Name -match '.*����Text.*') {
            $sheet = $book.Sheets.Item($_.Name);
        }
    }

    # �s���J�E���g(C��̕����������Ă���s���𐔂���)
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
Main����
- TEL-Equipment�ȉ��̍�ƃt�H���_��I��
- �I�������t�H���_�������[�v���Ė|��t�@�C����������
- �e�|��t�@�C���̖|�󕶏͍s���𐔂��ăR���\�[���ɏo��
----------------------------------------------------
#>
# �Ώۃt�@�C���̑I��
[hashtable]$ini = Get-IniFile ($PSScriptRoot + ".\config.ini")
[string]$basePath = $ini["START_PATH"]["WORKSPACE"];
[string]$tgtFolderPath = Get-Folder $basePath;
if ($tgtFolderPath -eq "") { exit; }

[string]$tgtFilePath = "";
[string]$rowsCount = "";
Get-ChildItem -Path $tgtFolderPath -Directory | ForEach-Object {
    if (Test-Path ($_.FullName + "\�|��\")) {
        $tgtFilePath = Get-TranslateFile ($_.FullName + "\�|��");
        $rowsCount = CountTgtFileRows $tgtFilePath;

        Write-Host (Split-Path $tgtFilePath -Leaf);
        Write-Host ($rowsCount.ToString() + "�s") -ForegroundColor Cyan;
    }
}