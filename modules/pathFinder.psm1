using namespace System.Windows.Forms;
Add-Type -AssemblyName System.Windows.Forms;
Add-Type -AssemblyName System.Drawing;
[Application]::EnableVisualStyles();

# ���W���[���ǂݍ���
Import-Module ((Split-Path $PSScriptRoot) + "\modules\utility.psm1") -Force;

<#
----------------------------------------------------
.SYNOPOSIS
    �����o�X�N���v�g�N�����̐ݒ�t�H�[�����J��

.PARAMETER menus
    ���X�g�{�b�N�X�ɕ\�����镶���z��
.PARAMETER destPath
    ��o��t�H���_�I�����̋N�_�ƂȂ�p�X
.PARAMETER sourcePath
    �R�s�[���t�H���_�I�����̋N�_�ƂȂ�p�X
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

    # �t�H�[���ݒ�
    $form = New-Object Form -Property @{
        Text = "�����o"
        Size = New-Object Drawing.Size(($scrWidth * 0.3), ($scrHeight * 0.6))
        MaximizeBox = $false
        TopMost = $True
        StartPosition = "CenterScreen"
        Font = New-Object Drawing.Font("Meiryo UI", 9)
    }
    
    # ���X�g�{�b�N�X�ݒ�
    $listBox = New-Object ListBox -Property @{
        Size = New-Object Drawing.Size(($form.Width * 0.7), ($form.height * 0.3))
        Font = New-Object Drawing.Font("Meiryo UI", 11)
    }
    $listBoxXPos = (($form.Width - $listBox.Width) / 2) - 10;
    $listBox.Location = New-Object System.Drawing.Point($listBoxXPos, 260);
    # ���X�g�{�b�N�X�ɍ��ڂ�ǉ�
    foreach ($menu in $menus) { [void] $listBox.Items.Add($menu); }

    # ���x���ݒ�
    $label = New-Object Label -Property @{
        Text = "��o�����I��"
        Size = New-Object Drawing.Size(230, 20)
        AutoSize = $true;
        Location = New-Object Drawing.Point($listBoxXPos, ($listBox.Top - 30))
    }

    # �{�^���ݒ�
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
        Text = "�R�s�[���I��"
        Size = New-Object Drawing.Size(100, 30)
        AutoSize = $True
        Location = New-Object Drawing.Point($listBoxXPos, 10)
    }
    $pickerBtn1.Add_Click({ $textBox1.Text = Get-Folder $sourcePath; })

    $pickerBtn2 = New-Object Button -Property @{
        Text = "�R�s�[��I��"
        Size = New-Object Drawing.Size(100, 30)
        AutoSize = $True
        Location = New-Object Drawing.Point($listBoxXPos, 115)
    }
    $pickerBtn2.Add_Click({ $textBox2.Text = Get-Folder $destPath; })

    # �e�L�X�g�{�b�N�X�ݒ�
    $textBox1 = New-Object TextBox -Property @{
        Width = $listBox.Width
        Location = New-Object Drawing.Point($listBoxXPos, 55)
    }
    $textBox2 = New-Object TextBox -Property @{
        Width = $listBox.Width
        Location = New-Object Drawing.Point($listBoxXPos, 155)
    }

    # �t�H�[���ɃA�C�e����ǉ�
    $form.Controls.AddRange(@(
        $label,
        $textBox1, $textBox2,
        $listBox, 
        $OKBtn, $CancelBtn, $pickerBtn1, $pickerBtn2
    ));

    # �L�[�ƃ{�^���̊֌W
    $form.AcceptButton = $OKBtn
    $form.CancelButton= $CancelBtn

    # �t�H�[����\��
    $result = $form.ShowDialog();

    # �t�H�[�����e�����^�[��
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
    �|��˗���o�X�N���v�g�N�����̐ݒ�t�H�[�����J��

.PARAMETER sourcePath
    �R�s�[���t�H���_�I�����̋N�_�ƂȂ�p�X
.PARAMETER destPath
    ��o��t�H���_�I�����̋N�_�ƂȂ�p�X
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

    # �t�H�[���ݒ�
    $form = New-Object Form -Property @{
        Text = "�|��˗���o"
        Size = New-Object Drawing.Size(($scrWidth * 0.3), ($scrHeight * 0.4))
        MaximizeBox = $false
        TopMost = $True
        StartPosition = "CenterScreen"
        Font = New-Object Drawing.Font("Meiryo UI", 9)
    }

    # �I�u�W�F�N�g���t�H�[���̒����ɔz�u���邽�߂̍��W
    $maxObjectWidth = ($form.Width * 0.7);
    $centeringXPos = (($form.width - $maxObjectWidth) / 2) - 10;

    # �e�L�X�g�{�b�N�X�ݒ�
    $textBox1 = New-Object TextBox -Property @{
        Width = $maxObjectWidth
        Location = New-Object Drawing.Point($centeringXPos, 75)
    }
    $textBox2 = New-Object TextBox -Property @{
        Width = $maxObjectWidth
        Location = New-Object Drawing.Point($centeringXPos, 195)
    }

    # �{�^���ݒ�
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
        Text = "�R�s�[���I��"
        Size = New-Object Drawing.Size(100, 30)
        AutoSize = $true
        Location = New-Object Drawing.Point($centeringXPos, 30)
    }
    $pickerBtn1.Add_Click({ $textBox1.Text = Get-Folder $sourcePath; })

    $pickerBtn2 = New-Object Button -Property @{
        Text = "�R�s�[��I��"
        Size = New-Object Drawing.Size(100, 30)
        AutoSize = $true
        Location = New-Object Drawing.Point($centeringXPos, 155)
    }
    $pickerBtn2.Add_Click({ $textBox2.Text = Get-Folder $destPath; })

    # �t�H�[���ɃA�C�e����ǉ�
    $form.Controls.AddRange(@(
        $textBox1, $textBox2,
        $OKBtn, $CancelBtn, $pickerBtn1, $pickerBtn2
    ));
    $form.AcceptButton = $OKBtn;
    $form.CancelButton= $CancelBtn;

    # �t�H�[����\��
    $result = $form.ShowDialog();

    # �t�H�[�����e�����^�[��
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


# �֐����J
Export-ModuleMember -Function showSubmitForm;
Export-ModuleMember -Function showTranslateRequestForm