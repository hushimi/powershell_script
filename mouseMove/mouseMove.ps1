# .NET��Cursor�N���X�𗘗p���邽�߂�System.Windows.Forms�����[�h
add-type -AssemblyName System.Windows.Forms

# mouse_event API�𗘗p���邽�߂̏���
$signature=@' 
      [DllImport("user32.dll",CharSet=CharSet.Auto, CallingConvention=CallingConvention.StdCall)]
      public static extern void mouse_event(long dwFlags, long dx, long dy, long cButtons, long dwExtraInfo);
'@
$SendMouseEvent = Add-Type -memberDefinition $signature -name "Win32MouseEventNew" -namespace Win32Functions -passThru;

# �}�E�X����ݒ� -----------------------------------

# �}�E�X�ړ��EDown�EUp
$MOUSEEVENTF_MOVE = 0x00000001;
$MOUSEEVENTF_LEFT_DOWN = 0x0002;
$MOUSEEVENTF_LEFT_UP = 0x0004;

# �E�}�E�X�̈ړ��C�x���g�����p�̐U�ꕝ
$MoveMouseDistance = 2;
# �E�}�E�X�̍��W�����E�ɂ��炷�p�̐U�ꕝ
$MoveMouseDistanceX = 20;
# �����񐔖ڂ͍��ցA��񐔖ڂŉE�ւ��炷���߂̃t���O
$isMoveLeft = $true;

# ---------------------------------------------------

Write-Host Happy slacking off!! -ForegroundColor Yellow;
Write-Host Press Ctrl+C to stop -ForegroundColor DarkYellow;
$SleepSec = 3;
while ($true) {
    Start-Sleep $SleepSec;

    # Teams���A�N�e�B�u�ɂ���

    # �N���b�N����
    $SendMouseEvent::mouse_event($MOUSEEVENTF_LEFT_DOWN, 0, 0, 0, 0);
    $SendMouseEvent::mouse_event($MOUSEEVENTF_LEFT_UP, 0, 0, 0, 0);

    # ���݂̃}�E�X��X,Y���W���擾
    $x = [System.Windows.Forms.Cursor]::Position.X;
    $y = [System.Windows.Forms.Cursor]::Position.Y;

    # �}�E�X�����E�ɐU��
    $SendMouseEvent::mouse_event($MOUSEEVENTF_MOVE, -$MoveMouseDistance, 0, 0, 0);
    $SendMouseEvent::mouse_event($MOUSEEVENTF_MOVE, $MoveMouseDistance, 0, 0, 0);

    # �}�E�X�J�[�\���ʒu�ړ�
    if ($isMoveLeft) {
        $x += $MoveMouseDistanceX;
        $isMoveLeft = $false;
    }
    else {
        $x -= $MoveMouseDistanceX;
        $isMoveLeft = $true;
    }
    # �}�E�X�J�[�\���ړ�
    [System.Windows.Forms.Cursor]::Position = new-object System.Drawing.Point($x, $y);
    $x = [System.Windows.Forms.Cursor]::Position.X;
    $y = [System.Windows.Forms.Cursor]::Position.Y;
}