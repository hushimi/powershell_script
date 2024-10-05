function prompt() {
    $cmdPromptUser = $env:USERNAME;
    $date = Get-Date -Format "yyyy/MM/dd HH:mm:ss";
    $cuPath = "$($pwd)"
    Write-Host "";
    Write-Host "$($date) " -NoNewline -ForegroundColor red;
    Write-Host "$($cuPath)" -ForegroundColor Cyan;
    Write-Host "[$($cmdPromptUser)]" -NoNewline -ForegroundColor Green;


    return " > "
}

<#
.SYNOPSIS
    comment out multiline in ise
.DESCRIPTION

.PARAMETER Items
.INPUTS
.OUTPUTS
#>
function multiLineComment() {
    $file = $psise.CurrentFile;
    $text = $file.Editor.SelectedText;
    if ($text.StartsWith("<#")) {
        $comment = $text.Substring(2).TrimEnd("#>");
    }
    else {
        $comment = "<#" + $text + "#>";
    }
    $file.Editor.InsertText($comment);
}
$psise.CurrentPowerShellTab.AddOnsMenu.Submenus.Add('Toggle Comment', { multiLineComment }, 'CTRL+K');
$Host.UI.RawUI.BackgroundColor = "black"
$Host.UI.RawUI.ForegroundColor = "white"
