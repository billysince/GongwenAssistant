$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Drawing.Imaging;
public class W {
    [DllImport("user32.dll")]
    public static extern IntPtr GetDesktopWindow();
    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hwnd, out RECT rc);
    [DllImport("user32.dll")]
    public static extern bool PrintWindow(IntPtr hwnd, IntPtr hdcBlt, uint nFlags);
    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hwnd);
    [DllImport("user32.dll")]
    public static extern int GetWindowTextLength(IntPtr hwnd);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int GetWindowText(IntPtr hwnd, System.Text.StringBuilder s, int n);
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hwnd);
    [DllImport("user32.dll")]
    public static extern bool ShowWindowAsync(IntPtr hwnd, int nCmdShow);
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left, Top, Right, Bottom; }
}
"@ -ReferencedAssemblies System.Drawing

Param(
    [int]$Hwnd,
    [string]$OutPath = 'D:\工作\20260520\GongwenAssistant\dist\wps_screenshot.png'
)

if (-not $Hwnd) {
    $p = Get-Process wps -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1
    if (-not $p) { Write-Host 'no WPS window'; exit 1 }
    $Hwnd = [int]$p.MainWindowHandle
}
Write-Host ('hwnd=' + $Hwnd)

# 让 WPS 顶起来
[void][W]::ShowWindowAsync([IntPtr]$Hwnd, 9)  # SW_RESTORE
Start-Sleep -Milliseconds 500
[void][W]::SetForegroundWindow([IntPtr]$Hwnd)
Start-Sleep -Milliseconds 800

$rc = New-Object W+RECT
[void][W]::GetWindowRect([IntPtr]$Hwnd, [ref]$rc)
$w = $rc.Right - $rc.Left
$h = $rc.Bottom - $rc.Top
Write-Host ("rect: $($rc.Left),$($rc.Top) - $($rc.Right),$($rc.Bottom)  size: ${w}x${h}")

if ($w -le 0 -or $h -le 0) { Write-Host 'window size invalid'; exit 2 }

$bmp = New-Object System.Drawing.Bitmap $w, $h
$g = [System.Drawing.Graphics]::FromImage($bmp)
$hdc = $g.GetHdc()
[void][W]::PrintWindow([IntPtr]$Hwnd, $hdc, 2)  # PW_RENDERFULLCONTENT
$g.ReleaseHdc($hdc)
$g.Dispose()

$outDir = Split-Path -Parent $OutPath
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Force -Path $outDir | Out-Null }
$bmp.Save($OutPath, [System.Drawing.Imaging.ImageFormat]::Png)
$bmp.Dispose()

Write-Host ('saved: ' + $OutPath)
Write-Host ('size : ' + (Get-Item $OutPath).Length + ' bytes')
