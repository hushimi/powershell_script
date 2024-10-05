# ���W���[����荞��
Import-Module ((Split-path $PSScriptRoot) + "\modules\pathFinder.psm1") -Force;
Import-Module ((Split-Path $PSScriptRoot) + "\modules\utility.psm1") -Force;
Add-Type -AssemblyName System.Windows.Forms;

<#
----------------------------------------------------
.SYNOPOSIS
    �R�s�[��t�H���_�̃p�X���擾

.PARAMETER srcFolderName
    �R�s�[���t�H���_�̖��O
.PARAMETER destParentPath
    ��o��t�H���_���i�[����Ă���e�t�H���_�̃p�X
.PARAMETER lang
    ��o�Ώۂ̌���
----------------------------------------------------
#>
function getDestinationPath {
    Param(
        [string]$srcFolderName,
        [string]$destParentPath,
        [string]$lang
    )

    # ���ꂲ�Ƃ̐��K�\���ݒ�
    [string]$regExp = "";
    switch ($lang) {
        "Japanese" {$regExp = "^J"}
        "English" {$regExp = "^E"}
        "Korea" {$regExp = "^K"}
        "China" {$regExp = "^C"}
        "Taiwan" {$regExp = "^T"}
    }

    # �R�s�[���t�H���_������Ǘ��ԍ��ƃR�����폜
    [string]$movieTitle = ([regex]::Replace($srcFolderName, "^J[0-9]{4}_", "")) -replace "[ �F |�F]";
    
    # ���K�\���Ƀ}�b�` And ����^�C�g�����܂ރt�H���_�̃t���p�X�擾
    # �ϊ��Ɏ�ނ�����L������
    [string]$destFullPath = "";
    [string]$baseName = "";
    Get-ChildItem -Path $destParentPath -Directory | ForEach-Object {
        $baseName = ($_.BaseName -replace "[ �F |�F|:]");
        if(($baseName -like "*" + $movieTitle + "*") -and ($_.BaseName -match $regExp)) {
            $destFullPath = $_.FullName;
        }
    }

    return $destFullPath;
}

<#
----------------------------------------------------
.SYNOPOSIS
    �R�s�[����MP4�t�@�C���Ɩ|�󕶏̓t�@�C���̃p�X���擾

.PARAMETER srcFolderPath
    �R�s�[���t�H���_�̃t���p�X
.PARAMETER lang
    ��o�Ώۂ̌���
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

    # ���ꂲ�Ƃ̐��K�\���ݒ�
    [string]$regExp = "";
    [string]$translateStartPath = "";
    switch ($lang) {
        "Japanese" {$regExp = "^J"; $translateStartPath = "\�|��"}
        "English" {$regExp = "^E"; $translateStartPath = "\�|��\English\"}
        "Korea" {$regExp = "^K"; $translateStartPath = "\�|��\Korea\"}
        "China" {$regExp = "^C"; $translateStartPath = "\�|��\China\"}
        "Taiwan" {$regExp = "^T"; $translateStartPath = "\�|��\Taiwan\"}
    }

    [int]$highestMp4Revision = 0;
    [int]$cuMp4Revision = 0;
    # ��ԃ��r�W�����̍���MP4�t�@�C���t���p�X���擾
    Get-ChildItem -Path $srcFolderPath -File -Filter "*.mp4" | ForEach-Object {
        if(($_.BaseName -match $regExp + ".*_r[0-9]{1,2}")) {
            # ��ԍ������r�W�����̔���
            $cuMp4Revision = [int]($_.BaseName -replace ($regExp + ".*_r"));
            if($cuMp4Revision -gt $highestMp4Revision) {
                $returnHash.mp4 = $_.FullName;
                $highestMp4Revision = $cuMp4Revision;
            }
        }
    }

    # ��ԃ��r�W�����̍����|��t�@�C��(Excel)�̃t���p�X���擾
    [int]$highestExcelRevision = 0;
    [int]$cuExcelRevision = 0;
    if (-not (Test-Path ($srcFolderPath + $translateStartPath))) { return $returnHash; }
    Get-ChildItem -Path ($srcFolderPath + $translateStartPath) -Filter "*.xlsx" | ForEach-Object {
        # ���{��̏ꍇ
        if(($lang -eq "Japanese") -and ($_.BaseName -match "^�y�|��z")) {
            $returnHash.translateFile = $_.FullName;

        # ������̏ꍇ
        } else {
            # ��ԍ������r�W�����̔���
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
Main����
- �t�H�[���ŃR�s�[���A�R�s�[��̃p�X�E�Ώی����I��
- MP4�t�@�C���E�|�󕶏͂̃R�s�[�����s
----------------------------------------------------
#>

# �t�H���_�I�����̋N�_
[hashtable]$ini = Get-IniFile ($PSScriptRoot + ".\config.ini")
[string]$source = $ini["START_PATH"]["WORKSPACE"];
[string]$dest = $ini["START_PATH"]["DELIVERY"];

# �t�H���_�I���t�H�[���\��(pathFinder.psm1)
# Property : menu, destination path, source path
[hashtable]$scriptConf = showSubmitForm @("Japanese", "English", "Korea", "China", "Taiwan") $dest $source;


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

# �t�@�C���R�s�[
[string]$destFullPath = "";
Get-ChildItem -Path $scriptConf.source -Directory | ForEach-Object {
    
    # �Ǘ��ԍ��̂���t�H���_���R�s�[����Ώ�
    if($_.Name -match "^J[0-9]{4}_") {
        # �R�s�[��̃p�X�擾
        [string]$destFullPath = getDestinationPath $_.Name $scriptConf.dest $scriptConf.lang;

        # Property : mp4,translateFile
        [hashtable]$copyItems = getCopyFiles $_.FullName $scriptConf.lang;

        if(-not ($destFullPath -eq "") -and (Test-Path $destFullPath)) {
            Copy-Item $copyItems.mp4 $destFullPath;
            Copy-Item $copyItems.translateFile $destFullPath;
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


