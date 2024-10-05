using namespace System.Windows.Forms;
Add-Type -AssemblyName System.Windows.Forms;
Add-Type -AssemblyName System.Drawing;
[Application]::EnableVisualStyles();

<#
----------------------------------------------------
.SYNOPOSIS
    �t�H���_�[�I���_�C�A���O�\��

.INPUTS initialDirectory
    �t�H���_�I�����̋N�_�ƂȂ�p�X
.OUTPUTS
    �I�������t�H���_�̃p�X
----------------------------------------------------
#>
Function Get-Folder {
    Param(
        [string]$initialDirectory
    )
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null;

    $folderBrowser = New-Object FolderBrowserDialog -Property @{
        Description = "�t�H���_��I��"
        SelectedPath = $initialDirectory
        ShowNewFolderButton = $false
    }
    
    [string]$folderPath = "";
    if($folderBrowser.ShowDialog() -eq "OK") {
        $folderPath = $folderBrowser.SelectedPath;
    }
    return $folderPath;
}

<#
----------------------------------------------------
.SYNOPOSIS
    �t�@�C���I���_�C�A���O�\��

.INPUTS initialDirectory
    �t�H���_�I�����̋N�_�ƂȂ�p�X
.OUTPUTS
    �I�������t�@�C���̃p�X
----------------------------------------------------
#>
Function Get-FilePath {
    Param(
        [string]$initialDirectory,
        [string]$filterStr,
        [bool]$isMultiSelect
    )

    $fileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
        Title = "�t�@�C����I��"
        InitialDirectory = $initialDirectory
        Filter = $filterStr
        Multiselect = $isMultiSelect
    }

    [string]$dialogResult = $fileBrowser.ShowDialog();

    if($dialogResult -eq "OK" -and -not ($isMultiSelect)) {
        return $fileBrowser.FileName;
    } elseif($dialogResult -eq "OK" -and $isMultiSelect) {
        return $fileBrowser.FileNames;
    } else {
        return "";
    }
}


<#
----------------------------------------------------
.SYNOPOSIS
    �V���[�g�J�b�g���w�肵���p�X�ɍ쐬����

.INPUTS sourcePath
    �V���[�g�J�b�g�̌��p�X
.INPUTS shortcutArgs
    �V���[�g�J�b�g�ɐݒ肷�����
.INPUTS destinationPath
    �V���[�g�J�b�g��ݒu����ꏊ
----------------------------------------------------
#>
Function CreateShortcut {
    Param(
        [string]$sourcePath,
        [string]$shortcutArgs,
        [string]$destinationPath
    )

    $wshShell = New-Object -ComObject WScript.Shell;
    $shortCut = $wshShell.CreateShortcut($destinationPath);
    $shortCut.TargetPath = $sourcePath;
    $shortCut.Arguments = $shortcutArgs;
    $shortCut.Save();
}

function Get-IniFile {
    <#    
    .DESCRIPTION
    �w�肳�ꂽ�p�X��ini�t�@�C����ǂݎ��Akey-value�^�̘A�z�z��ɕϊ�
    
    .PARAMETER filePath
    ini�t�@�C���̃p�X
    
    .PARAMETER anonymous
    The section name to use for the anonymous section (keys that come before any section declaration).
    
    .PARAMETER comments
    Enables saving of comments to a comment section in the resulting hash table.
    The comments for each section will be stored in a section that has the same name as the section of its origin, but has the comment suffix appended.
    Comments will be keyed with the comment key prefix and a sequence number for the comment. The sequence number is reset for every section.
    
    .PARAMETER commentsSectionsSuffix
    The suffix for comment sections. The default value is an underscore ('_').

    .PARAMETER commentsKeyPrefix
    The prefix for comment keys. The default value is 'Comment'.
    
    .EXAMPLE
    !!!�g�p��!!!
    Get-IniFile /path/to/my/inifile.ini
    $ini[section][key]�Œl�ɃA�N�Z�X
    
    .NOTES
    The resulting hash table has the form [sectionName->sectionContent], 
    where sectionName is a string and sectionContent is a hash table of the form [key->value] where both are strings.
    This function is largely copied from https://stackoverflow.com/a/43697842/1031534
    A modified version with a working example is pulished at https://gist.github.com/seakintruth/05da4c3c38f72c4b796aeae9be453e8e

    .LICENSE
    Attribution-ShareAlike 4.0 International (CC BY-SA 4.0)
    #>
    
    param(
        [parameter(Mandatory = $true)] [string] $filePath,
        [string] $anonymous = 'NoSection',
        [switch] $comments,
        [string] $commentsSectionsSuffix = '_',
        [string] $commentsKeyPrefix = 'Comment'
    )

    # �w�肵��ini�t�@�C���̊e�s���C�e���[�g���A�A�z�z����쐬
    $ini = @{}
    switch -regex -file ($filePath) {
        # ini�t�@�C��Section��key�Ɋ܂߂邩����
        "^\[(.+)\]$" {
            $section = $matches[1]
            $ini[$section] = @{}
            $CommentCount = 0
            if ($comments) {
                $commentsSection = $section + $commentsSectionsSuffix
                $ini[$commentsSection] = @{}
            }
            continue
        }

        # ini�t�@�C��Comment��z��Ɋ܂߂邩����
        "^(;.*)$" {
            if ($comments) {
                if (!($section)) {
                    $section = $anonymous
                    $ini[$section] = @{}
                }
                $value = $matches[1]
                $CommentCount = $CommentCount + 1
                $name = $commentsKeyPrefix + $CommentCount
                $commentsSection = $section + $commentsSectionsSuffix
                $ini[$commentsSection][$name] = $value
            }
            continue
        }

        # key-value�^���쐬
        # section�������ꍇ�A$ini[NoSection][key]=$value�ƂȂ�
        "^(.+?)\s*=\s*(.*)$" {
            if (!($section)) {
                $section = $anonymous
                $ini[$section] = @{}
            }
            # key-value�𒊏o
            $name, $value = $matches[1..2]
            $ini[$section][$name] = $value
            continue
        }
    }
    return $ini
}


# �֐����J
Export-ModuleMember -Function Get-Folder;
Export-ModuleMember -Function Get-FilePath;
Export-ModuleMember -Function CreateShortcut;
Export-ModuleMember -Function Get-IniFile;