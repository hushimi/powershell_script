# ���W���[����荞��
Import-Module ((Split-path $PSScriptRoot) + "\modules\pathFinder.psm1") -Force;
Import-Module ((Split-path $PSScriptRoot) + "\modules\utility.psm1") -Force;
Add-Type -AssemblyName System.Windows.Forms;

<#
----------------------------------------------------
.SYNOPOSIS
    �|��˗��t�H���_����R�s�[�Ώۂ�Excel�p�X���擾

.PARAMETER targetName
    �R�s�[��t�H���_�̖��O
.PARAMETER sourcePath
    �|��˗��t�H���_�̃p�X
.PARAMETER lang
    ��o�Ώۂ̌���
----------------------------------------------------
#>
function getSourcePath {
    Param(
        [string]$targetName,
        [string]$sourcePath,
        [string]$lang
    )

    # �|�󕶎擾�Ώۂ̃t�H���_���ݒ�
    [string]$targetLang = "";
    switch ($lang) {
        "English" {$targetLang = "�p��"}
        "Korea" {$targetLang = "�؍���"}
        "China" {$targetLang = "�ȑ̎�"}
        "Taiwan" {$targetLang = "�ɑ̎�"}
    }

    # �R�s�[���t�H���_������Ǘ��ԍ��ƃR�����폜
    [string]$movieTitle = ([regex]::Replace($targetName, "^J[0-9]{4}_", "")) -replace "[ �F |�F]";
    
    # ����^�C�g�����܂ރt�H���_����t���p�X���擾
    [string]$translateFolderPath = "";
    [string]$translateFilePath = "";
    [string]$baseName = "";
    Get-ChildItem -Path $sourcePath -Directory | ForEach-Object {
        $baseName = ($_.BaseName -replace "[ �F |�F|:]");
        if($baseName -like "*" + $movieTitle + "*") {
            $translateFolderPath = ($_.FullName + "\" + $targetLang + "\");
        }
    }

    # �R�s�[���t�H���_����|��Excel�t�@�C���̃p�X���擾
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
Main����
- �t�H�[���ŃR�s�[���A�R�s�[��̃p�X�E�Ώی����I��
- MP4�t�@�C���E�|�󕶏͂̃R�s�[�����s
----------------------------------------------------
#>
# �t�H���_�I�����̋N�_
[hashtable]$ini = Get-IniFile ($PSScriptRoot + ".\config.ini")
[string]$dest = $ini["START_PATH"]["WORKSPACE"];
[string]$source = $ini["START_PATH"]["TRANSLATE"];

# �t�H���_�I���t�H�[���\��(pathFinder.psm1)
# Property : menu, destination path, source path
[hashtable]$scriptConf = showSubmitForm @("English", "Korea", "China", "Taiwan") $dest $source;

# ���͒l����
[bool]$isContinue = $true;
foreach ($key in $scriptConf.Keys) {
    $confVal = $scriptConf[$key];
    if ($confVal -eq "") { $isContinue = $false; }
}
if (-not $isContinue) {
    [System.Windows.Forms.MessageBox]::Show(
        "�p�X�ƑΏی���̂����ꂩ�����͂���Ă��܂���B", 
        "submitProducts",
        "OK",
        "Error"
    );
    exit;
}

# �t�@�C���R�s�[����
[string]$destFullPath = "";
Get-ChildItem -Path $scriptConf.dest -Directory | ForEach-Object {
    
    # �Ǘ��ԍ��̂���t�H���_���R�s�[����Ώ�
    if($_.Name -match "^J[0-9]{4}_") {
        # �R�s�[����|�󕶏͂̃p�X���擾
        [string]$copyItemPath = getSourcePath $_.Name $scriptConf.source $scriptConf.lang;

        # �R�s�[��̃p�X�ݒ�
        $destFullPath = ($_.FullName + "\�|��\" + $scriptConf.lang + "\");

        # ��ԍ����|��t�@�C���̃��r�W�����擾
        [int]$highestExcelRevision = 0;
        [int]$cuExcelRevision = 0;
        Get-ChildItem -Path ($destFullPath) -Filter "*.xlsx" | ForEach-Object {
            # ��ԍ������r�W�����̔���
            $cuExcelRevision = [int]($_.BaseName -replace "^.*_r");
            if($cuExcelRevision -gt $highestExcelRevision) {
                $highestExcelRevision = $cuExcelRevision;
            }
        }

        # �t�@�C���R�s�[���s
        if(-not ($copyItemPath -eq "") -and (Test-Path $destFullPath)) {
            # �R�s�[���t�@�C����ڂ̃R�s�[
            [string]$newFileName = Split-Path $copyItemPath -Leaf;
            $newFileName = $newFileName -replace "�y.*�z", "�yATC�z";
            $newFileName = $newFileName -replace "_r\d{1,2}", "";
            $newFileName = ($newFileName -replace "\.xlsx", "") + "_" `
                + ($scriptConf.lang.Substring(0,1)) + "_r" `
                + ($highestExcelRevision + 1)`
            ;
            Copy-Item $copyItemPath ($destFullPath + "\" + $newFileName + ".xlsx");

            # �R�s�[���t�@�C����ڂ̃R�s�[
            $newFileName = $newFileName -replace "_r[1-9]{1,2}", "";
            $newFileName += "_r" + ($highestExcelRevision + 2);
            Copy-Item $copyItemPath ($destFullPath + "\" + $newFileName + ".xlsx");

        } else {
            [System.Windows.Forms.MessageBox]::Show(
                ($_.Name) + "�̃R�s�[�Ɏ��s���܂����B`r`n`r`n�R�s�[���܂��̓R�s�[��Ƀt�H���_�����݂��Ȃ��\��������܂��B", 
                "submitProducts",
                "OK",
                "Error"
            );
        }
    }
}

# �R�s�[���ƃR�s�[����J��
Start-Process $scriptConf.dest;
Start-Process $scriptConf.source;


