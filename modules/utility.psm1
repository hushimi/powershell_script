using namespace System.Windows.Forms;
Add-Type -AssemblyName System.Windows.Forms;
Add-Type -AssemblyName System.Drawing;
[Application]::EnableVisualStyles();

<#
----------------------------------------------------
.SYNOPOSIS
    フォルダー選択ダイアログ表示

.INPUTS initialDirectory
    フォルダ選択時の起点となるパス
.OUTPUTS
    選択したフォルダのパス
----------------------------------------------------
#>
Function Get-Folder {
    Param(
        [string]$initialDirectory
    )
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null;

    $folderBrowser = New-Object FolderBrowserDialog -Property @{
        Description = "フォルダを選択"
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
    ファイル選択ダイアログ表示

.INPUTS initialDirectory
    フォルダ選択時の起点となるパス
.OUTPUTS
    選択したファイルのパス
----------------------------------------------------
#>
Function Get-FilePath {
    Param(
        [string]$initialDirectory,
        [string]$filterStr,
        [bool]$isMultiSelect
    )

    $fileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
        Title = "ファイルを選択"
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
    ショートカットを指定したパスに作成する

.INPUTS sourcePath
    ショートカットの元パス
.INPUTS shortcutArgs
    ショートカットに設定する引数
.INPUTS destinationPath
    ショートカットを設置する場所
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
    指定されたパスのiniファイルを読み取り、key-value型の連想配列に変換
    
    .PARAMETER filePath
    iniファイルのパス
    
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
    !!!使用例!!!
    Get-IniFile /path/to/my/inifile.ini
    $ini[section][key]で値にアクセス
    
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

    # 指定したiniファイルの各行をイテレートし、連想配列を作成
    $ini = @{}
    switch -regex -file ($filePath) {
        # iniファイルSectionをkeyに含めるか判定
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

        # iniファイルCommentを配列に含めるか判定
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

        # key-value型を作成
        # sectionが無い場合、$ini[NoSection][key]=$valueとなる
        "^(.+?)\s*=\s*(.*)$" {
            if (!($section)) {
                $section = $anonymous
                $ini[$section] = @{}
            }
            # key-valueを抽出
            $name, $value = $matches[1..2]
            $ini[$section][$name] = $value
            continue
        }
    }
    return $ini
}


# 関数公開
Export-ModuleMember -Function Get-Folder;
Export-ModuleMember -Function Get-FilePath;
Export-ModuleMember -Function CreateShortcut;
Export-ModuleMember -Function Get-IniFile;