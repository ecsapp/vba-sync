# Screenshot helpers for excel-control.
#
# Save-WindowImage  → Win32 PrintWindow (works on obscured / hidden-behind windows)
# Save-ScreenImage  → System.Drawing CopyFromScreen of a screen rectangle
#
# Both write a PNG file and return the path.

Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue

if (-not ([System.Management.Automation.PSTypeName]'XcCapture.Win32').Type) {
  Add-Type @"
using System;
using System.Runtime.InteropServices;

namespace XcCapture {
  [StructLayout(LayoutKind.Sequential)]
  public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }

  public static class Win32 {
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr hWnd, IntPtr hdcBlt, uint nFlags);
    [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr hWnd);
    public const uint PW_CLIENTONLY     = 0x1;
    public const uint PW_RENDERFULLCONTENT = 0x2;  // composited capture; Win 8.1+
  }
}
"@
}

# Captures a window by HWND, including content obscured by other windows.
# Returns the path on success; throws on invalid HWND or capture failure.
function Save-WindowImage {
    param(
        [Parameter(Mandatory=$true)] [int64]$Hwnd,
        [Parameter(Mandatory=$true)] [string]$Path
    )
    $h = [IntPtr]$Hwnd
    if (-not [XcCapture.Win32]::IsWindow($h)) {
        throw "Save-WindowImage: not a valid HWND ($Hwnd)"
    }
    $rect = New-Object XcCapture.RECT
    [void][XcCapture.Win32]::GetWindowRect($h, [ref]$rect)
    $w = $rect.Right - $rect.Left
    $hgt = $rect.Bottom - $rect.Top
    if ($w -le 0 -or $hgt -le 0) { throw "Save-WindowImage: window has zero dimensions" }

    $bmp = New-Object System.Drawing.Bitmap $w, $hgt
    $g   = [System.Drawing.Graphics]::FromImage($bmp)
    $hdc = $g.GetHdc()
    try {
        # PW_RENDERFULLCONTENT (0x2) needed for composited content (Edge,
        # Chromium-host controls). Falls back to standard if unsupported.
        [void][XcCapture.Win32]::PrintWindow($h, $hdc, [XcCapture.Win32]::PW_RENDERFULLCONTENT)
    } finally {
        $g.ReleaseHdc($hdc)
        $g.Dispose()
    }
    $bmp.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
    return $Path
}

# Captures a screen rectangle (or full virtual screen if -Rect not supplied).
# Returns the path. Use for the main Excel window or specific monitor regions.
function Save-ScreenImage {
    param(
        [Parameter(Mandatory=$true)] [string]$Path,
        [int]$X      = -1,
        [int]$Y      = -1,
        [int]$Width  = -1,
        [int]$Height = -1
    )
    if ($X -lt 0) {
        Add-Type -AssemblyName System.Windows.Forms
        $vs = [System.Windows.Forms.SystemInformation]::VirtualScreen
        $X = $vs.X; $Y = $vs.Y; $Width = $vs.Width; $Height = $vs.Height
    }
    $bmp = New-Object System.Drawing.Bitmap $Width, $Height
    $g   = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($X, $Y, 0, 0, (New-Object System.Drawing.Size($Width, $Height)))
    $g.Dispose()
    $bmp.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
    return $Path
}
