function prompt {
    $cmdPromptUser = $env:USERNAME;
    $date = Get-Date -Format "yyyy/MM/dd HH:mm:ss";
    $cuPath = "$($pwd)"
    Write-Host "";
    Write-Host "$($date) " -NoNewline -ForegroundColor red;
    Write-Host "$($cuPath)" -ForegroundColor Cyan;
    Write-Host "[$($cmdPromptUser)]" -NoNewline -ForegroundColor Green;

    return " > "
}

function tools {
    Set-Location ($HOME + "\Documents\03_Tool");
}

function dev {
    Set-Location ($HOME + "\Documents\");
}