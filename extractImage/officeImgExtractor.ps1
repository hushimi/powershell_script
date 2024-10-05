# ���W���[����荞��
Import-Module ((Split-path $PSScriptRoot) + "\modules\utility.psm1") -Force;
Add-Type -AssemblyName System.Windows.Forms;

<#
----------------------------------------------------
Main����
- �X�N���v�g���sdir��media�t�H���_������΍폜
- �摜���o����Office�t�@�C����I��
- �t�@�C�����X�N���v�g���sdir�ɃR�s�[
- �t�@�C����zip�ɕϊ�
- 7zip��zip��temp dir�ɓW�J
- �W�J����dir����e�g���q�Ɋ�Â����p�X��media�t�H���_�𒊏o
- ���o����media�t�H���_��I�������t�@�C���Ɠ��K�w�ɐݒu
----------------------------------------------------
#>

# startExtract.bat�ɂĐݒ肵���ϐ���$PSScriptRoot�ɐݒ�
[string]$scriptPath = (cmd /c "echo %CUDIR%");

# media �t�H���_�̍폜
if (Test-Path ($PSScriptRoot + "\media\")) {
    Write-Host "Deleting old media folder..." -ForegroundColor Yellow;
    Remove-Item -Recurse -Force ($PSScriptRoot + "\media\");
    cls;
}

# �Ώۃt�@�C���̑I��
[hashtable]$ini = Get-IniFile ($PSScriptRoot + ".\config.ini")
[string]$basePath = $ini["START_PATH"]["WORKSPACE"];

$tgtFilePath = Get-FilePath $basePath "Office Files|*.docx;*.xlsx;*.pptx" $false;
if ($tgtFilePath -eq "") { exit; }
[string]$tgtFileName = [System.IO.Path]::GetFileNameWithoutExtension($tgtFilePath);
[string]$tgtExtension = [System.IO.Path]::GetExtension($tgtFilePath);
[string]$tgtParentPath = Split-Path $tgtFilePath -Parent;
Write-Host $tgtParentPath -ForegroundColor Green;

# �t�@�C����ZIP�Ƃ��ăR�s�[�E�W�J
Write-Host "�I�������t�@�C�����R�s�[��..." -ForegroundColor Yellow;
Copy-Item $tgtFilePath ($PSScriptRoot + "\" + $tgtFileName + ".zip");
Write-Host "�R�s�[����" -ForegroundColor Cyan;
New-Item -ItemType Directory ($PSScriptRoot + "\temp\");
Expand-Archive -Path ($PSScriptRoot + "\" + $tgtFileName + ".zip") -DestinationPath ($PSScriptRoot + "\temp\");

# �摜�t�H���_�����o��
[string]$extractPath = "";
switch ($tgtExtension) {
    ".pptx" {$extractPath = "ppt\media"}
    ".xlsx" {$extractPath = "xl\media"}
    ".docx" {$extractPath = "word\media"}
}
Write-Host "�摜�����o����..." -ForegroundColor Yellow;
Copy-Item -Recurse ($PSScriptRoot + "\temp\" + $extractPath) -Destination $PSScriptRoot;
Write-Host "�����o������" -ForegroundColor Cyan;

# ��n��
Remove-Item -Recurse ($PSScriptRoot + "\temp\"), ($PSScriptRoot + "\" + $tgtFileName + ".zip");
cls;
Move-Item ($scriptPath + "media") ($tgtParentPath + "\" + $tgtFileName + "(�摜)");
Start-Process ($tgtParentPath);
