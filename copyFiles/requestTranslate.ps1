# ���W���[����荞��
Import-Module ((Split-path $PSScriptRoot) + "\modules\pathFinder.psm1") -Force;
Import-Module ((Split-path $PSScriptRoot) + "\modules\utility.psm1") -Force;
Add-Type -AssemblyName System.Windows.Forms;

<#
----------------------------------------------------
.SYNOPOSIS
    �R�s�[��t�H���_�̃p�X���擾

.PARAMETER srcFolderName
    �R�s�[���t�H���_�̖��O
.PARAMETER destParentPath
    ��o��t�H���_���i�[����Ă���e�t�H���_�̃p�X
----------------------------------------------------
#>
function getDestinationPath {
    Param(
        [string]$srcFolderName,
        [string]$destParentPath
    )

    # �R�s�[���t�H���_������Ǘ��ԍ��ƃR�����폜
    [string]$movieTitle = ([regex]::Replace($srcFolderName, "^J[0-9]{4}_", "")) -replace "[ �F |�F]";
    
    # ���K�\���Ƀ}�b�` And ����^�C�g�����܂ރt�H���_�̃t���p�X�擾
    # �ϊ��Ɏ�ނ�����L������
    [string]$destFullPath = "";
    [string]$baseName = "";
    Get-ChildItem -Path $destParentPath -Directory | ForEach-Object {
        $baseName = ($_.BaseName -replace "[ �F |�F|:]");
        if(($baseName -like "*" + $movieTitle + "*")) {
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
        movieName = "";
        translateFile = "";
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
                $returnHash.movieName = $_.BaseName;
                $highestMp4Revision = $cuMp4Revision;
            }
        }
    }

    # �|��t�@�C��(Excel)�̃t���p�X���擾
    [int]$highestExcelRevision = 0;
    [int]$cuExcelRevision = 0;
    if (-not (Test-Path ($srcFolderPath + "\�|��\"))) { return $returnHash; }
    Get-ChildItem -Path ($srcFolderPath + "\�|��\") -Filter "*.xlsx" | ForEach-Object {
        if(($_.BaseName -match "^�y�|��z")) {
            $returnHash.translateFile = $_.FullName;
        }
    }

    return $returnHash;
}

<#
----------------------------------------------------
.SYNOPOSIS
    �e����t�H���_����̏ꍇ�A�|��t�@�C�����R�s�[����

.PARAMETER projPath
    ��o��t�H���_�̃t���p�X
.PARAMETER excelFileName
    �R�s�[���t�H���_�̃t���p�X
----------------------------------------------------
#>
function copyToLangFile {
    Param(
        [string]$projPath,
        [string]$excelFileName
    )
    
    # �|��t�@�C���i�[������[�v������t�H���_������Ȃ�|��t�@�C���R�s�[
    [string]$copyFileName = "";
    Get-ChildItem -Path $projPath -Directory | ForEach-Object {
        if ((Get-ChildItem $_.FullName | Measure-Object).Count -eq 0) {
            $copyFileName = [regex]::Replace($excelFileName, "^�y.*�z", ("�y" + $_.Name + "�z"));
            Copy-Item ($projPath + "\" + $excelFileName) ($projPath + "\" + $_.Name + "\" + $copyFileName);
        }
    }

}

<#
----------------------------------------------------
Main����
- �t�H�[���ŃR�s�[���A�R�s�[��̃p�X��I��
- ���{��MP4�t�@�C���V���[�g�J�b�g�쐬�E�|�󕶏̓R�s�[
----------------------------------------------------
#>

# �t�H���_�I�����̋N�_
[hashtable]$ini = Get-IniFile ($PSScriptRoot + ".\config.ini")
[string]$source = $ini["START_PATH"]["WORKSPACE"];
[string]$dest = $ini["START_PATH"]["TRANSLATE"];

# �t�H���_�I���t�H�[���\��
# Property : source, dest
[hashtable]$scriptConf = showTranslateRequestForm $source $dest;

# ���͒l����
[bool]$isContinue = $true;
foreach ($key in $scriptConf.Keys) {
    $confVal = $scriptConf[$key];
    if ($confVal -eq "") { $isContinue = $false; }
}
if (-not $isContinue) {
    [System.Windows.Forms.MessageBox]::Show(
        "�p�X�����͂���Ă��܂���B", 
        "�|��˗���o",
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
        [string]$destFullPath = getDestinationPath $_.Name $scriptConf.dest;

        # Property : mp4,movieName, translateFile
        [hashtable]$copyItems = getCopyFiles $_.FullName $scriptConf.lang;

        if(-not ($destFullPath -eq "") -and (Test-Path $destFullPath)) {
            Copy-Item $copyItems.translateFile $destFullPath;
            CreateShortcut $copyItems.mp4 "" ($destFullPath + "\" + $copyItems.movieName + ".lnk");
        } else {
            [System.Windows.Forms.MessageBox]::Show(
                ($_.Name) + "�̃R�s�[�Ɏ��s���܂����B`r`n`r`n�R�s�[���܂��̓R�s�[��Ƀt�H���_�����݂��Ȃ��\��������܂��B", 
                "submitProducts",
                "OK",
                "Error"
            );
        }

        copyToLangFile $destFullPath (Split-Path $copyItems.translateFile -Leaf);
    }
}

# �R�s�[���ƃR�s�[����J��
Start-Process $scriptConf.dest;
Start-Process $scriptConf.source;


