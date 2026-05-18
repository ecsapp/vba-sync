# Unlock a password-protected VBA project programmatically.
#
# WHY
#   Excel's VBA project can be password-locked (VBProject.Protection = 1).
#   While locked, VBComponents.Count returns 0 and Import() fails — so the
#   strip-and-import sync step that vba-sync's ImportProject relies on can't
#   work, and most COM automation against the project breaks.
#
# HOW
#   Microsoft doesn't expose unlock via COM directly, but VBE.CommandBars
#   does — same UI model as Excel's own. The "VBAProject Properties..."
#   menu item has control id 2578 in the VBE Tools bar; calling .Execute()
#   on it raises the modal "VBAProject Password" dialog whether VBE is
#   foreground or not.
#
#   Once the dialog exists we drive it via Win32:
#     1. Enumerate top-level #32770 dialogs in the Excel PID, find the one
#        titled "VBAProject Password".
#     2. Find its first Edit child control, WM_SETTEXT the password,
#        PostMessage WM_COMMAND IDOK to click OK.
#     3. Project unlocks (Protection -> 0, VBComponents enumerable).
#     4. Close the Properties dialog that surfaces underneath via WM_CLOSE.
#
#   No SendKeys = no foreground focus required — survives the user
#   interacting with other apps while the script runs.
#
# USAGE
#   . excel-control/unlock-vba-project.ps1
#   $ok = Unlock-VbaProject -Xl $xl -Password 'secret'
#   if (-not $ok) { throw "Unlock failed" }

if (-not ([System.Management.Automation.PSTypeName]'VbaUnlock.Win32').Type) {
  Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Text;
namespace VbaUnlock {
  public static class Win32 {
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr p);
    [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr parent, EnumProc cb, IntPtr p);
    [DllImport("user32.dll", CharSet=CharSet.Auto)] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll", CharSet=CharSet.Auto)] public static extern int GetClassName(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll", SetLastError=true)] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll", CharSet=CharSet.Auto)] public static extern IntPtr SendMessage(IntPtr h, uint msg, IntPtr w, string l);
    [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h, uint msg, IntPtr w, IntPtr l);
    public const uint WM_SETTEXT = 0x000C;
    public const uint WM_COMMAND = 0x0111;
    public const uint WM_CLOSE = 0x0010;
    public const int IDOK = 1;
    public delegate bool EnumProc(IntPtr h, IntPtr l);
  }
}
"@
}

function Find-VbaPasswordDialog {
  param([uint32]$ProcId)
  $script:_dlg = [IntPtr]::Zero
  $cb = [VbaUnlock.Win32+EnumProc]{ param([IntPtr]$h, [IntPtr]$l)
    if (-not [VbaUnlock.Win32]::IsWindowVisible($h)) { return $true }
    $cls = New-Object System.Text.StringBuilder 256
    [VbaUnlock.Win32]::GetClassName($h, $cls, 256) | Out-Null
    if ($cls.ToString() -ne '#32770') { return $true }
    $wpid = [uint32]0
    [VbaUnlock.Win32]::GetWindowThreadProcessId($h, [ref]$wpid) | Out-Null
    if ($wpid -ne $ProcId) { return $true }
    $t = New-Object System.Text.StringBuilder 256
    [VbaUnlock.Win32]::GetWindowText($h, $t, 256) | Out-Null
    if ($t.ToString() -eq 'VBAProject Password') { $script:_dlg = $h; return $false }
    return $true
  }
  [VbaUnlock.Win32]::EnumWindows($cb, [IntPtr]::Zero) | Out-Null
  return $script:_dlg
}

function Find-VbaPropertiesDialog {
  param([uint32]$ProcId)
  $script:_propDlg = [IntPtr]::Zero
  $cb = [VbaUnlock.Win32+EnumProc]{ param([IntPtr]$h, [IntPtr]$l)
    if (-not [VbaUnlock.Win32]::IsWindowVisible($h)) { return $true }
    $cls = New-Object System.Text.StringBuilder 256
    [VbaUnlock.Win32]::GetClassName($h, $cls, 256) | Out-Null
    if ($cls.ToString() -ne '#32770') { return $true }
    $wpid = [uint32]0
    [VbaUnlock.Win32]::GetWindowThreadProcessId($h, [ref]$wpid) | Out-Null
    if ($wpid -ne $ProcId) { return $true }
    $t = New-Object System.Text.StringBuilder 256
    [VbaUnlock.Win32]::GetWindowText($h, $t, 256) | Out-Null
    if ($t.ToString() -like 'VBAProject*Properties*' -or $t.ToString() -like '*Project Properties*') {
      $script:_propDlg = $h; return $false
    }
    return $true
  }
  [VbaUnlock.Win32]::EnumWindows($cb, [IntPtr]::Zero) | Out-Null
  return $script:_propDlg
}

function Find-FirstEditChild {
  param([IntPtr]$Parent)
  $script:_edit = [IntPtr]::Zero
  $cb = [VbaUnlock.Win32+EnumProc]{ param([IntPtr]$h, [IntPtr]$l)
    $cls = New-Object System.Text.StringBuilder 256
    [VbaUnlock.Win32]::GetClassName($h, $cls, 256) | Out-Null
    if ($cls.ToString() -like 'Edit*' -or $cls.ToString() -like 'RichEdit*') {
      $script:_edit = $h
      return $false
    }
    return $true
  }
  [VbaUnlock.Win32]::EnumChildWindows($Parent, $cb, [IntPtr]::Zero) | Out-Null
  return $script:_edit
}

function Unlock-VbaProject {
  param(
    [Parameter(Mandatory=$true)] $Xl,
    [Parameter(Mandatory=$true)] [string]$Password,
    [int]$DialogWaitMs = 2500,
    [switch]$Trace
  )

  function _trace($msg) { if ($Trace) { Write-Host "  [unlock] $msg" -ForegroundColor DarkGray } }

  if ($Xl.ActiveWorkbook.VBProject.Protection -eq 0) { _trace 'already unlocked'; return $true }
  _trace 'protection=1, starting unlock'

  $xlPid = [uint32]0
  [VbaUnlock.Win32]::GetWindowThreadProcessId([IntPtr]$Xl.Hwnd, [ref]$xlPid) | Out-Null

  # Both Excel main window and VBE main window must be alive for the
  # menu item's Execute() to surface the modal Password dialog. Restore
  # visibility on exit.
  $xlWasVisible  = $Xl.Visible
  $vbeWasVisible = $Xl.VBE.MainWindow.Visible
  $Xl.Visible = $true
  Start-Sleep -Milliseconds 300
  $Xl.VBE.MainWindow.Visible = $true
  Start-Sleep -Milliseconds 300

  # Control id 2578 = "VBAProject Properties..." in the VBE Tools bar.
  $ctl = $null
  try { $ctl = $Xl.VBE.CommandBars.FindControl([Type]::Missing, 2578) } catch { }
  if ($null -eq $ctl) {
    _trace 'could not FindControl(2578); falling back to Tools bar lookup'
    try { $ctl = $Xl.VBE.CommandBars.Item('Tools').Controls.Item(5) } catch { }
  }
  if ($null -eq $ctl) {
    _trace 'no VBAProject Properties control found'
    $Xl.VBE.MainWindow.Visible = $vbeWasVisible
    return $false
  }

  try { $ctl.Execute() } catch { _trace ("Execute threw: " + $_.Exception.Message) }

  $dlg = [IntPtr]::Zero
  $elapsed = 0
  while ($elapsed -lt $DialogWaitMs) {
    Start-Sleep -Milliseconds 150
    $elapsed += 150
    $dlg = Find-VbaPasswordDialog -ProcId $xlPid
    if ($dlg -ne [IntPtr]::Zero) { break }
  }
  _trace ("password dialog: 0x{0:X} after {1}ms" -f $dlg.ToInt64(), $elapsed)
  if ($dlg -eq [IntPtr]::Zero) {
    $Xl.VBE.MainWindow.Visible = $vbeWasVisible
    $Xl.Visible = $xlWasVisible
    return $false
  }

  Start-Sleep -Milliseconds 250
  $edit = Find-FirstEditChild -Parent $dlg
  _trace ("edit child: 0x{0:X}" -f $edit.ToInt64())
  if ($edit -eq [IntPtr]::Zero) {
    [VbaUnlock.Win32]::PostMessage($dlg, [VbaUnlock.Win32]::WM_CLOSE, [IntPtr]::Zero, [IntPtr]::Zero) | Out-Null
    $Xl.VBE.MainWindow.Visible = $vbeWasVisible
    $Xl.Visible = $xlWasVisible
    return $false
  }
  [VbaUnlock.Win32]::SendMessage($edit, [VbaUnlock.Win32]::WM_SETTEXT, [IntPtr]::Zero, $Password) | Out-Null
  Start-Sleep -Milliseconds 250
  [VbaUnlock.Win32]::PostMessage($dlg, [VbaUnlock.Win32]::WM_COMMAND, [IntPtr][VbaUnlock.Win32]::IDOK, [IntPtr]::Zero) | Out-Null
  Start-Sleep -Milliseconds 600

  # The Properties dialog is now visible beneath; close it.
  $propDlg = Find-VbaPropertiesDialog -ProcId $xlPid
  if ($propDlg -ne [IntPtr]::Zero) {
    [VbaUnlock.Win32]::PostMessage($propDlg, [VbaUnlock.Win32]::WM_CLOSE, [IntPtr]::Zero, [IntPtr]::Zero) | Out-Null
    Start-Sleep -Milliseconds 300
  }

  $Xl.VBE.MainWindow.Visible = $vbeWasVisible
  $Xl.Visible = $xlWasVisible

  $protAfter = $Xl.ActiveWorkbook.VBProject.Protection
  _trace ("final Protection=$protAfter")
  return ($protAfter -eq 0)
}
