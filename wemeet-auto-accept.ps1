# wemeet-auto-accept.ps1
# 腾讯会议远程控制请求弹窗自动接受
# 检测策略：Win32 枚举 WeMeetApp 可见小窗口(弹窗尺寸) -> 截图像素识别蓝色[接受]按钮 -> 点击
# 不依赖 UIA 文本（腾讯会议 Qt 自绘弹窗不暴露"接受"等无障碍名称）
# 通过 Startup 目录快捷方式开机自启（无窗口后台运行）

param(
    [int]$IntervalSec = 5,
    [int]$MaxMinutes = 0
)

$ErrorActionPreference = "SilentlyContinue"

Add-Type -AssemblyName System.Drawing
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class WemeetWin {
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc cb, IntPtr lParam);
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool SetCursorPos(int X, int Y);
    [DllImport("user32.dll")] public static extern void mouse_event(uint f, uint dx, uint dy, uint d, IntPtr e);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    public struct RECT { public int Left, Top, Right, Bottom; }
    public const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
    public const uint MOUSEEVENTF_LEFTUP = 0x0004;
    public static void Click(int x, int y) {
        SetCursorPos(x, y);
        mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, IntPtr.Zero);
        mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, IntPtr.Zero);
    }
}
"@

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$logFile = Join-Path $scriptDir "wemeet-auto-accept.log"

function Write-Log($msg) {
    $line = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $msg
    Add-Content -Path $logFile -Value $line -Encoding UTF8
}

function Get-WeMeetWindows {
    $list = New-Object System.Collections.ArrayList
    $procs = Get-Process | Where-Object { $_.ProcessName -like "WeMeetApp*" } | Select-Object -ExpandProperty Id
    if (-not $procs) { return $list }
    $set = New-Object 'System.Collections.Generic.HashSet[int]'
    foreach ($p in $procs) { [void]$set.Add([int]$p) }
    $cb = [WemeetWin+EnumWindowsProc]{
        param($h, $l)
        $pid2 = [uint32]0
        [WemeetWin]::GetWindowThreadProcessId($h, [ref]$pid2) | Out-Null
        if ($set.Contains([int]$pid2) -and [WemeetWin]::IsWindowVisible($h)) {
            $r = New-Object WemeetWin+RECT
            [WemeetWin]::GetWindowRect($h, [ref]$r) | Out-Null
            $w = $r.Right - $r.Left; $ht = $r.Bottom - $r.Top
            if ($w -ge 250 -and $w -le 700 -and $ht -ge 120 -and $ht -le 350 -and $r.Left -gt -100) {
                [void]$list.Add([PSCustomObject]@{ H=$h; X=$r.Left; Y=$r.Top; W=$w; Ht=$ht })
            }
        }
        return $true
    }
    [WemeetWin]::EnumWindows($cb, [IntPtr]::Zero) | Out-Null
    return $list
}

# 截取窗口区域，在下半部找蓝色[接受]按钮，返回按钮中心屏幕坐标；找不到返回 $null
function Find-AcceptButton($win) {
    # 定位策略：找下半部"最长连续蓝色实心横条"（即[接受]按钮），取其行集合的中心。
    # 不用包围盒：弹窗阴影/圆角透出的零散蓝像素会拉偏包围盒。
    $bmp = New-Object System.Drawing.Bitmap $win.W, $win.Ht
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($win.X, $win.Y, 0, 0, (New-Object System.Drawing.Size($win.W, $win.Ht)))
    $g.Dispose()

    $isBlue = {
        param($p)
        return ($p.B -ge 170 -and $p.R -le 130 -and $p.G -ge 60 -and $p.G -le 200)
    }

    $anchorRun = $null   # @{X0;X1;Y;Len}
    $runs = New-Object System.Collections.ArrayList
    $yStart = [int]($win.Ht * 0.5)
    for ($y = $yStart; $y -lt $win.Ht; $y += 2) {
        $runStart = -1; $lastBlue = -2
        for ($x = 0; $x -lt $win.W; $x += 2) {
            $blue = (& $isBlue $bmp.GetPixel($x, $y))
            if ($blue) {
                if ($runStart -lt 0) { $runStart = $x }
                $lastBlue = $x
            } elseif ($runStart -ge 0 -and ($x - $lastBlue) -gt 3) {
                # 连续段结束（允许 1 个采样间隙）
                $len = $lastBlue - $runStart
                if ($len -ge 16) { [void]$runs.Add(@{ X0=$runStart; X1=$lastBlue; Y=$y; Len=$len }) }
                $runStart = -1
            }
        }
        if ($runStart -ge 0) {
            $len = $lastBlue - $runStart
            if ($len -ge 16) { [void]$runs.Add(@{ X0=$runStart; X1=$lastBlue; Y=$y; Len=$len }) }
        }
    }
    $bmp.Dispose()
    if ($runs.Count -eq 0) { return $null }

    # 最长横条为锚点
    foreach ($r in $runs) { if (-not $anchorRun -or $r.Len -gt $anchorRun.Len) { $anchorRun = $r } }
    if ($anchorRun.Len -lt 16) { return $null }

    # 收集与锚点 x 范围重叠 >=50% 且长度 >= 锚点 60% 的行，求精确中心
    $xSum = 0; $xN = 0; $ys = New-Object System.Collections.ArrayList
    foreach ($r in $runs) {
        $ov = [Math]::Min($r.X1, $anchorRun.X1) - [Math]::Max($r.X0, $anchorRun.X0)
        if ($ov -gt 0 -and $ov -ge ($anchorRun.Len * 0.5) -and $r.Len -ge ($anchorRun.Len * 0.6)) {
            [void]$ys.Add($r.Y)
            $xSum += ($r.X0 + $r.X1) / 2; $xN++
        }
    }
    $cx = [int]($xSum / [Math]::Max(1, $xN))
    $ysArr = $ys | Sort-Object
    $cy = [int]$ysArr[[int]($ysArr.Count / 2)]

    # 点击前像素校验：中心点必须是蓝色
    $bmp2 = New-Object System.Drawing.Bitmap 1, 1
    $g2 = [System.Drawing.Graphics]::FromImage($bmp2)
    $g2.CopyFromScreen($win.X + $cx, $win.Y + $cy, 0, 0, (New-Object System.Drawing.Size(1, 1)))
    $g2.Dispose()
    $centerPx = $bmp2.GetPixel(0, 0)
    $bmp2.Dispose()
    if (-not (& $isBlue $centerPx)) { return $null }

    return @{ X = $win.X + $cx; Y = $win.Y + $cy; Cnt = $runs.Count }
}

Write-Log "监控启动 pid=$PID interval=${IntervalSec}s mode=pixel-detect"

$deadline = if ($MaxMinutes -gt 0) { (Get-Date).AddMinutes($MaxMinutes) } else { [DateTime]::MaxValue }
$attempts = @{}      # hwnd -> 已尝试点击次数
$lastAliveLog = Get-Date

while ((Get-Date) -lt $deadline) {
    try {
        $wins = Get-WeMeetWindows
        foreach ($win in $wins) {
            $h = [int64]$win.H
            $tried = 0
            if ($attempts.ContainsKey($h)) { $tried = $attempts[$h] }
            if ($tried -ge 6) { continue }

            $btn = Find-AcceptButton $win
            if ($btn) {
                Write-Log ("发现疑似远程控制弹窗 hwnd={0} rect=({1},{2}) {3}x{4} 蓝色像素={5}" -f $h, $win.X, $win.Y, $win.W, $win.Ht, $btn.Cnt)
                [WemeetWin]::SetForegroundWindow($win.H) | Out-Null
                Start-Sleep -Milliseconds 250
                [WemeetWin]::Click([int]$btn.X, [int]$btn.Y)
                Write-Log ("已点击[接受] ({0},{1}) 第{2}次尝试" -f $btn.X, $btn.Y, ($tried + 1))
                Start-Sleep -Seconds 2
                # 验证：弹窗是否已消失
                $still = Get-WeMeetWindows | Where-Object { [int64]$_.H -eq $h }
                if (-not $still) {
                    Write-Log "弹窗已消失，接受成功"
                    $attempts[$h] = 6
                } else {
                    $attempts[$h] = $tried + 1
                    Write-Log "弹窗仍在，稍后重试"
                }
            }
        }
        # 每 60 秒打一条存活日志，便于确认监控在跑
        if (((Get-Date) - $lastAliveLog).TotalSeconds -ge 60) {
            Write-Log ("存活检查: 已扫描窗口数={0}" -f $wins.Count)
            $lastAliveLog = Get-Date
        }
        if ($attempts.Count -gt 100) { $attempts.Clear() }
    } catch {
        Write-Log "轮询异常: $($_.Exception.Message)"
    }
    Start-Sleep -Seconds $IntervalSec
}

Write-Log "监控结束"
