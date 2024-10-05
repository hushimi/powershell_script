# .NETのCursorクラスを利用するためにSystem.Windows.Formsをロード
add-type -AssemblyName System.Windows.Forms

# mouse_event APIを利用するための準備
$signature=@' 
      [DllImport("user32.dll",CharSet=CharSet.Auto, CallingConvention=CallingConvention.StdCall)]
      public static extern void mouse_event(long dwFlags, long dx, long dy, long cButtons, long dwExtraInfo);
'@
$SendMouseEvent = Add-Type -memberDefinition $signature -name "Win32MouseEventNew" -namespace Win32Functions -passThru;

# マウス動作設定 -----------------------------------

# マウス移動・Down・Up
$MOUSEEVENTF_MOVE = 0x00000001;
$MOUSEEVENTF_LEFT_DOWN = 0x0002;
$MOUSEEVENTF_LEFT_UP = 0x0004;

# ・マウスの移動イベント生成用の振れ幅
$MoveMouseDistance = 2;
# ・マウスの座標を左右にずらす用の振れ幅
$MoveMouseDistanceX = 20;
# 偶数回数目は左へ、奇数回数目で右へずらすためのフラグ
$isMoveLeft = $true;

# ---------------------------------------------------

Write-Host Happy slacking off!! -ForegroundColor Yellow;
Write-Host Press Ctrl+C to stop -ForegroundColor DarkYellow;
$SleepSec = 3;
while ($true) {
    Start-Sleep $SleepSec;

    # Teamsをアクティブにする

    # クリックする
    $SendMouseEvent::mouse_event($MOUSEEVENTF_LEFT_DOWN, 0, 0, 0, 0);
    $SendMouseEvent::mouse_event($MOUSEEVENTF_LEFT_UP, 0, 0, 0, 0);

    # 現在のマウスのX,Y座標を取得
    $x = [System.Windows.Forms.Cursor]::Position.X;
    $y = [System.Windows.Forms.Cursor]::Position.Y;

    # マウスを左右に振る
    $SendMouseEvent::mouse_event($MOUSEEVENTF_MOVE, -$MoveMouseDistance, 0, 0, 0);
    $SendMouseEvent::mouse_event($MOUSEEVENTF_MOVE, $MoveMouseDistance, 0, 0, 0);

    # マウスカーソル位置移動
    if ($isMoveLeft) {
        $x += $MoveMouseDistanceX;
        $isMoveLeft = $false;
    }
    else {
        $x -= $MoveMouseDistanceX;
        $isMoveLeft = $true;
    }
    # マウスカーソル移動
    [System.Windows.Forms.Cursor]::Position = new-object System.Drawing.Point($x, $y);
    $x = [System.Windows.Forms.Cursor]::Position.X;
    $y = [System.Windows.Forms.Cursor]::Position.Y;
}