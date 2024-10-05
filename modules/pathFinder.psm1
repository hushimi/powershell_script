using namespace System.Windows.Forms;
Add-Type -AssemblyName System.Windows.Forms;
Add-Type -AssemblyName System.Drawing;
[Application]::EnableVisualStyles();

# モジュール読み込み
Import-Module ((Split-Path $PSScriptRoot) + "\modules\utility.psm1") -Force;

<#
----------------------------------------------------
.SYNOPOSIS
    動画提出スクリプト起動時の設定フォームを開く

.PARAMETER menus
    リストボックスに表示する文字配列
.PARAMETER destPath
    提出先フォルダ選択時の起点となるパス
.PARAMETER sourcePath
    コピー元フォルダ選択時の起点となるパス
----------------------------------------------------
#>
function showSubmitForm {
    Param(
        [string[]]$menus, 
        [string]$destPath,
        [string]$sourcePath
    )

    $scrSize = [System.Windows.Forms.SystemInformation]::PrimaryMonitorSize;
    [int]$scrWidth = $scrSize.Width;
    [int]$scrHeight = $scrSize.Height;

    # フォーム設定
    $form = New-Object Form -Property @{
        Text = "動画提出"
        Size = New-Object Drawing.Size(($scrWidth * 0.3), ($scrHeight * 0.6))
        MaximizeBox = $false
        TopMost = $True
        StartPosition = "CenterScreen"
        Font = New-Object Drawing.Font("Meiryo UI", 9)
    }
    
    # リストボックス設定
    $listBox = New-Object ListBox -Property @{
        Size = New-Object Drawing.Size(($form.Width * 0.7), ($form.height * 0.3))
        Font = New-Object Drawing.Font("Meiryo UI", 11)
    }
    $listBoxXPos = (($form.Width - $listBox.Width) / 2) - 10;
    $listBox.Location = New-Object System.Drawing.Point($listBoxXPos, 260);
    # リストボックスに項目を追加
    foreach ($menu in $menus) { [void] $listBox.Items.Add($menu); }

    # ラベル設定
    $label = New-Object Label -Property @{
        Text = "提出言語を選択"
        Size = New-Object Drawing.Size(230, 20)
        AutoSize = $true;
        Location = New-Object Drawing.Point($listBoxXPos, ($listBox.Top - 30))
    }

    # ボタン設定
    $BtnWidth = 75;
    $okBtnXpos = $listBoxXPos + $listBox.Width - ($BtnWidth * 2) - 10;
    $OKBtn = New-Object Button -Property @{
        Text = "OK"
        Size = New-Object Drawing.Size($BtnWidth, 30)
        Location = New-Object Drawing.Point($okBtnXpos, ($listbox.Bottom + 40))
        DialogResult = [System.Windows.Forms.DialogResult]::OK
    }
    
    $cancelBtnXpos = $listBoxXPos + $listBox.Width - ($BtnWidth);
    $CancelBtn = New-Object Button -Property @{
        Text = "Cancel"
        Size = New-Object Drawing.Size($BtnWidth, 30)
        Location = New-Object Drawing.Point($cancelBtnXpos, ($listbox.Bottom + 40))
        DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    }

    $pickerBtn1 = New-Object Button -Property @{
        Text = "コピー元選択"
        Size = New-Object Drawing.Size(100, 30)
        AutoSize = $True
        Location = New-Object Drawing.Point($listBoxXPos, 10)
    }
    $pickerBtn1.Add_Click({ $textBox1.Text = Get-Folder $sourcePath; })

    $pickerBtn2 = New-Object Button -Property @{
        Text = "コピー先選択"
        Size = New-Object Drawing.Size(100, 30)
        AutoSize = $True
        Location = New-Object Drawing.Point($listBoxXPos, 115)
    }
    $pickerBtn2.Add_Click({ $textBox2.Text = Get-Folder $destPath; })

    # テキストボックス設定
    $textBox1 = New-Object TextBox -Property @{
        Width = $listBox.Width
        Location = New-Object Drawing.Point($listBoxXPos, 55)
    }
    $textBox2 = New-Object TextBox -Property @{
        Width = $listBox.Width
        Location = New-Object Drawing.Point($listBoxXPos, 155)
    }

    # フォームにアイテムを追加
    $form.Controls.AddRange(@(
        $label,
        $textBox1, $textBox2,
        $listBox, 
        $OKBtn, $CancelBtn, $pickerBtn1, $pickerBtn2
    ));

    # キーとボタンの関係
    $form.AcceptButton = $OKBtn
    $form.CancelButton= $CancelBtn

    # フォームを表示
    $result = $form.ShowDialog();

    # フォーム内容をリターン
    [hashtable]$returnHash = @{
        source = "";
        dest = "";
        lang = ""
    }
    if ($result -eq "OK") {
        $returnHash.source = $textBox1.Text;
        $returnHash.dest = $textBox2.Text;
        $returnHash.lang = $listBox.SelectedItem;
    }

    return $returnHash;
}

<#
----------------------------------------------------
.SYNOPOSIS
    翻訳依頼提出スクリプト起動時の設定フォームを開く

.PARAMETER sourcePath
    コピー元フォルダ選択時の起点となるパス
.PARAMETER destPath
    提出先フォルダ選択時の起点となるパス
----------------------------------------------------
#>
function showTranslateRequestForm {
    Param(
        [string]$sourcePath,
        [string]$destPath
    )

    $scrSize = [System.Windows.Forms.SystemInformation]::PrimaryMonitorSize;
    [int]$scrWidth = $scrSize.Width;
    [int]$scrHeight = $scrSize.Height;

    # フォーム設定
    $form = New-Object Form -Property @{
        Text = "翻訳依頼提出"
        Size = New-Object Drawing.Size(($scrWidth * 0.3), ($scrHeight * 0.4))
        MaximizeBox = $false
        TopMost = $True
        StartPosition = "CenterScreen"
        Font = New-Object Drawing.Font("Meiryo UI", 9)
    }

    # オブジェクトをフォームの中央に配置するための座標
    $maxObjectWidth = ($form.Width * 0.7);
    $centeringXPos = (($form.width - $maxObjectWidth) / 2) - 10;

    # テキストボックス設定
    $textBox1 = New-Object TextBox -Property @{
        Width = $maxObjectWidth
        Location = New-Object Drawing.Point($centeringXPos, 75)
    }
    $textBox2 = New-Object TextBox -Property @{
        Width = $maxObjectWidth
        Location = New-Object Drawing.Point($centeringXPos, 195)
    }

    # ボタン設定
    $BtnWidth = 75;
    $okBtnXpos = (($centeringXPos + $textBox2.Width) - ($BtnWidth * 2)) - 10;
    $OKBtn = New-Object Button -Property @{
        Text = "OK"
        Size = New-Object Drawing.Size($BtnWidth, 30)
        Location = New-Object Drawing.Point($okBtnXpos, ($textBox2.Bottom + 40))
        DialogResult = [System.Windows.Forms.DialogResult]::OK
    }
    
    $cancelBtnXpos = ($centeringXPos + $textBox2.Width) - ($BtnWidth);
    $CancelBtn = New-Object Button -Property @{
        Text = "Cancel"
        Size = New-Object Drawing.Size($BtnWidth, 30)
        Location = New-Object Drawing.Point($cancelBtnXpos, ($textBox2.Bottom + 40))
        DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    }

    $pickerBtn1 = New-Object Button -Property @{
        Text = "コピー元選択"
        Size = New-Object Drawing.Size(100, 30)
        AutoSize = $true
        Location = New-Object Drawing.Point($centeringXPos, 30)
    }
    $pickerBtn1.Add_Click({ $textBox1.Text = Get-Folder $sourcePath; })

    $pickerBtn2 = New-Object Button -Property @{
        Text = "コピー先選択"
        Size = New-Object Drawing.Size(100, 30)
        AutoSize = $true
        Location = New-Object Drawing.Point($centeringXPos, 155)
    }
    $pickerBtn2.Add_Click({ $textBox2.Text = Get-Folder $destPath; })

    # フォームにアイテムを追加
    $form.Controls.AddRange(@(
        $textBox1, $textBox2,
        $OKBtn, $CancelBtn, $pickerBtn1, $pickerBtn2
    ));
    $form.AcceptButton = $OKBtn;
    $form.CancelButton= $CancelBtn;

    # フォームを表示
    $result = $form.ShowDialog();

    # フォーム内容をリターン
    [hashtable]$returnHash = @{
        source = "";
        dest = "";
    }
    if ($result -eq "OK") {
        $returnHash.source = $textBox1.Text;
        $returnHash.dest = $textBox2.Text;
    }

    return $returnHash;
}


# 関数公開
Export-ModuleMember -Function showSubmitForm;
Export-ModuleMember -Function showTranslateRequestForm