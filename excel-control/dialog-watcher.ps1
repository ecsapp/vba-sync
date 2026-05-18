# Win32 dialog watcher for headless Excel automation.
#
# WHAT IT DOES
#   Spawns a background runspace that polls (every 250ms) for any
#   Win32 dialog (class "#32770") owned by EXCEL.EXE. When one
#   appears, it captures the dialog's window title + visible text,
#   dismisses it (sends Enter), and stops watching.
#
# WHY
#   When a VBA macro raises an unhandled error mid-Application.Run,
#   Excel pops a modal "Microsoft Visual Basic" dialog. The COM
#   thread blocks until the dialog is dismissed. Without dialog
#   capture, headless Excel either hangs forever or returns a
#   useless generic HRESULT. With this watcher, we read the actual
#   error message ("Run-time error '9006': Dataset: Invalid number
#   of columns") and surface it to the caller.
#
# USAGE
#   . tools/dialog-watcher.ps1
#   $watcher = Start-DialogWatcher -ExcelPid $xl.Hwnd  # or by PID
#   try { $result = $xl.Run("...") } finally {
#     $captured = Stop-DialogWatcher $watcher
#     if ($captured) { Write-Host "Dialog captured: $captured" }
#   }
#
# CAVEATS
#   - Dismisses the dialog by sending IDOK (Enter). For dialogs
#     with multiple buttons (e.g. Debug/End/Help), this picks the
#     default — usually End for VBA runtime errors, OK for compile
#     errors. End is what we want for headless: it terminates the
#     macro and returns control to the caller.
#   - Polls at 250ms. A *very* fast macro that finishes before the
#     first poll won't trigger the watcher even if an error fires
#     on the way out. Acceptable; the macro returns its result via
#     normal COM in that case.

# ----- Load Win32 P/Invoke wrappers (only once) -----
if (-not ([System.Management.Automation.PSTypeName]'DialogWatcher.Win32').Type) {
  Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Text;

namespace DialogWatcher {
  public static class Win32 {
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int GetWindowTextLength(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

    public const uint WM_GETTEXT = 0x000D;
    public const uint WM_GETTEXTLENGTH = 0x000E;
    public const uint WM_COMMAND = 0x0111;
    public const uint WM_CLOSE = 0x0010;
    public const int IDOK = 1;
    public const int IDCANCEL = 2;
  }
}
"@
}

# ----- Helpers (script-scope, not exported) -----

function Get-WindowText {
  param([IntPtr]$Hwnd)
  $len = [DialogWatcher.Win32]::GetWindowTextLength($Hwnd)
  if ($len -le 0) { return "" }
  $sb = New-Object System.Text.StringBuilder ($len + 1)
  [DialogWatcher.Win32]::GetWindowText($Hwnd, $sb, $sb.Capacity) | Out-Null
  return $sb.ToString()
}

function Get-WindowClassName {
  param([IntPtr]$Hwnd)
  $sb = New-Object System.Text.StringBuilder 256
  [DialogWatcher.Win32]::GetClassName($Hwnd, $sb, $sb.Capacity) | Out-Null
  return $sb.ToString()
}

function Get-WindowProcessId {
  param([IntPtr]$Hwnd)
  $procId = [uint32]0
  [DialogWatcher.Win32]::GetWindowThreadProcessId($Hwnd, [ref]$procId) | Out-Null
  return $procId
}

# Find all top-level dialog windows (class #32770) owned by a given PID.
function Find-DialogsForPid {
  param([uint32]$ProcessId)
  $found = New-Object System.Collections.ArrayList
  $callback = {
    param([IntPtr]$h, [IntPtr]$lp)
    if (-not [DialogWatcher.Win32]::IsWindowVisible($h)) { return $true }
    $cls = Get-WindowClassName -Hwnd $h
    if ($cls -ne "#32770") { return $true }
    $wpid = Get-WindowProcessId -Hwnd $h
    if ($wpid -ne $ProcessId) { return $true }
    [void]$found.Add($h)
    return $true
  }
  $delegate = [DialogWatcher.Win32+EnumWindowsProc]$callback
  [DialogWatcher.Win32]::EnumWindows($delegate, [IntPtr]::Zero) | Out-Null
  return $found
}

# Walk the children of a dialog and concatenate all visible static-text
# (class "Static") + button labels. That's where the error message lives.
function Get-DialogContent {
  param([IntPtr]$Hwnd)
  $title = Get-WindowText -Hwnd $Hwnd
  $parts = New-Object System.Collections.ArrayList
  if ($title) { [void]$parts.Add("[$title]") }

  $childCallback = {
    param([IntPtr]$h, [IntPtr]$lp)
    $cls = Get-WindowClassName -Hwnd $h
    if ($cls -eq "Static" -or $cls -eq "Button") {
      $txt = Get-WindowText -Hwnd $h
      if ($txt -and $txt.Trim() -ne "") {
        [void]$parts.Add($txt.Trim())
      }
    }
    return $true
  }
  $delegate = [DialogWatcher.Win32+EnumWindowsProc]$childCallback
  [DialogWatcher.Win32]::EnumChildWindows($Hwnd, $delegate, [IntPtr]::Zero) | Out-Null

  return ($parts -join " | ")
}

# Dismiss a dialog by sending it WM_COMMAND IDCANCEL (preferred for
# VBA error dialogs because the default button is usually "End"
# rather than "Continue", and IDCANCEL maps to End on those).
# Falls back to WM_CLOSE if WM_COMMAND doesn't take.
function Dismiss-Dialog {
  param([IntPtr]$Hwnd)
  $WM_COMMAND = [DialogWatcher.Win32]::WM_COMMAND
  $WM_CLOSE = [DialogWatcher.Win32]::WM_CLOSE
  $IDCANCEL = [DialogWatcher.Win32]::IDCANCEL

  # Try WM_COMMAND IDCANCEL (= End for VBA runtime-error dialogs)
  [DialogWatcher.Win32]::PostMessage($Hwnd, $WM_COMMAND, [IntPtr]$IDCANCEL, [IntPtr]::Zero) | Out-Null
  Start-Sleep -Milliseconds 100
  if ([DialogWatcher.Win32]::IsWindowVisible($Hwnd)) {
    # Fallback: WM_CLOSE
    [DialogWatcher.Win32]::PostMessage($Hwnd, $WM_CLOSE, [IntPtr]::Zero, [IntPtr]::Zero) | Out-Null
  }
}

# ----- Public API -----

# Start watching for dialogs from a given Excel process. Returns a
# handle (a hashtable with state) you pass to Stop-DialogWatcher
# later to retrieve any captured dialog text.
function Start-DialogWatcher {
  param(
    [uint32]$ProcessId,
    [int]$PollMs = 250
  )

  $state = [hashtable]::Synchronized(@{
    ProcessId = $ProcessId
    Captured  = $null
    Stop      = $false
  })

  # Build the runspace + thread that polls
  $rs = [runspacefactory]::CreateRunspace()
  $rs.Open()
  $rs.SessionStateProxy.SetVariable("state", $state)
  $rs.SessionStateProxy.SetVariable("PollMs", $PollMs)
  $rs.SessionStateProxy.SetVariable("loadDef", (Get-Item function:\Find-DialogsForPid).Definition)
  $rs.SessionStateProxy.SetVariable("textDef", (Get-Item function:\Get-DialogContent).Definition)
  $rs.SessionStateProxy.SetVariable("dismissDef", (Get-Item function:\Dismiss-Dialog).Definition)
  $rs.SessionStateProxy.SetVariable("getTextDef", (Get-Item function:\Get-WindowText).Definition)
  $rs.SessionStateProxy.SetVariable("getClsDef", (Get-Item function:\Get-WindowClassName).Definition)
  $rs.SessionStateProxy.SetVariable("getPidDef", (Get-Item function:\Get-WindowProcessId).Definition)

  $ps = [powershell]::Create()
  $ps.Runspace = $rs
  [void]$ps.AddScript({
    # Re-import P/Invoke types in the runspace (already loaded
    # globally but each runspace needs its own type cache).
    if (-not ([System.Management.Automation.PSTypeName]'DialogWatcher.Win32').Type) {
      Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Text;
namespace DialogWatcher {
  public static class Win32 {
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int GetWindowTextLength(IntPtr hWnd);
    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);
    [DllImport("user32.dll", SetLastError = true)]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
    [DllImport("user32.dll")]
    public static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")]
    public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
  }
}
"@
    }
    Invoke-Expression "function Get-WindowText { $getTextDef }"
    Invoke-Expression "function Get-WindowClassName { $getClsDef }"
    Invoke-Expression "function Get-WindowProcessId { $getPidDef }"
    Invoke-Expression "function Find-DialogsForPid { $loadDef }"
    Invoke-Expression "function Get-DialogContent { $textDef }"
    Invoke-Expression "function Dismiss-Dialog { $dismissDef }"

    while (-not $state.Stop) {
      $dialogs = Find-DialogsForPid -ProcessId $state.ProcessId
      if ($dialogs.Count -gt 0) {
        # Capture text from the first dialog we see
        $hwnd = [IntPtr]$dialogs[0]
        $content = Get-DialogContent -Hwnd $hwnd
        $state.Captured = $content
        Dismiss-Dialog -Hwnd $hwnd
        # Continue watching in case a *second* dialog pops (e.g.
        # error dialog -> debug dialog after Cancel). Append.
      }
      Start-Sleep -Milliseconds $PollMs
    }
  })

  $handle = $ps.BeginInvoke()

  return [pscustomobject]@{
    State = $state
    PowerShell = $ps
    Handle = $handle
    Runspace = $rs
  }
}

# Stop the watcher, return any captured dialog text (or $null).
function Stop-DialogWatcher {
  param($Watcher)
  if (-not $Watcher) { return $null }
  $Watcher.State.Stop = $true
  try { [void]$Watcher.PowerShell.EndInvoke($Watcher.Handle) } catch {}
  try { $Watcher.PowerShell.Dispose() } catch {}
  try { $Watcher.Runspace.Close() } catch {}
  return $Watcher.State.Captured
}
