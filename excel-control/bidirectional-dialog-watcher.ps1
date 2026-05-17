# Bidirectional dialog watcher — the "reporter, not dismisser" model.
#
# WHAT IT DOES
#   Spawns a background runspace that polls (every PollMs ms) for any visible
#   non-chrome top-level window (Win32 #32770 dialogs AND UserForms) owned by
#   the target Excel PID. When one appears, it:
#
#     1. Reads the modal's caption, body text (Static + Button labels), and
#        button list with their hwnds.
#     2. Writes the snapshot as JSON to $ModalFile.
#     3. Echoes one line to stdout:
#          MODAL hwnd=<n> caption='<caption>' buttons=<csv>
#     4. Blocks (still polling) until $ActionFile exists.
#     5. Reads the chosen button label from $ActionFile.
#     6. Clicks via a fallback chain:
#          a. PostMessage WM_COMMAND with the button's discovered control-id
#             to the parent dialog (most reliable for #32770).
#          b. PostMessage BM_CLICK on the button hwnd.
#          c. PostMessage WM_KEYDOWN/WM_KEYUP VK_RETURN to the dialog
#             (UserForms whose CommandButtons aren't enumerable as Win32
#             Button — triggers the form's Default button Click handler).
#          d. PostMessage WM_CLOSE to the dialog (last resort).
#        After each attempt, polls IsWindow + IsWindowVisible for up to
#        800ms-1200ms to confirm dismissal. Only marks the hwnd "handled"
#        once it's confirmed gone; otherwise the next poll re-pushes the
#        same modal so the click can be retried.
#     7. Deletes both files, dedupes by hwnd, and resumes polling.
#
#   Net effect: an external agent (Claude Code, MCP client, CI script) sees
#   every modal as a discrete event and decides which button to click. No
#   blind auto-dismiss; no information loss.
#
# DIFFERENCE FROM dialog-watcher.ps1 (v1)
#   v1 is a dismisser. It captures the first modal's text into a single
#   field, sends WM_COMMAND IDCANCEL, and stops. Good for the vba-sync
#   export pipeline where every modal is a known failure to surface.
#   v2 reports every modal as an event, never dismisses on its own, and
#   delegates the dismissal action to an external caller. Required for any
#   workflow where modals carry information the agent needs to act on
#   (refresh success vs error, per-tab validation, per-form input).
#
#   Both can coexist. Use v1 for the vba-sync inner loop, v2 for any
#   workbook-driving scenario where the modal text is part of the workflow.
#
# IPC CONTRACT
#   $ModalFile  — JSON written by watcher, read by agent. Schema:
#       { "Hwnd": <int64>, "Caption": "...", "Body": "Static|Static|...",
#         "Buttons": [{ "Caption": "OK", "Hwnd": <int64> }, ...] }
#   $ActionFile — single-line UTF-8 text written by agent, read by watcher.
#       Body is the button label (case-insensitive match against Buttons[].Caption).
#       For modals where the button list is empty (UserForms) the watcher
#       falls through to VK_RETURN regardless of label.
#
#   The protocol is intentionally simple polling so anything that can read
#   and write files can drive it (Claude Code's Read/Write/Monitor tools,
#   plain bash, an MCP client, etc.). The eventual SCOPING.md session
#   protocol replaces these single-event files with append-only JSONL
#   streams under sessions/<id>/. This watcher's event shape is compatible
#   with that — moving to JSONL is a layer change, not a rewrite.
#
# USAGE
#   . tools/bidirectional-dialog-watcher.ps1
#   $watcher = Start-DialogWatcher -ProcessId $xlPid `
#                                  -ModalFile  'c:\tmp\dlg.json' `
#                                  -ActionFile 'c:\tmp\act.txt'
#   try {
#     $xl.Run("'Wbk.xlsm'!SomeMacroThatPopsAModal")
#     # While the macro runs, the agent reads $ModalFile and writes $ActionFile.
#   } finally {
#     Stop-DialogWatcher $watcher
#   }

if (-not ([System.Management.Automation.PSTypeName]'BiDialogWatcher.Win32').Type) {
  Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Text;

namespace BiDialogWatcher {
  public static class Win32 {
    [DllImport("user32.dll", CharSet = CharSet.Auto)] public static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);
    [DllImport("user32.dll", CharSet = CharSet.Auto)] public static extern int GetWindowTextLength(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr hWnd);
    [DllImport("user32.dll", SetLastError = true)] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
    [DllImport("user32.dll", CharSet = CharSet.Auto)] public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")] public static extern int GetDlgCtrlID(IntPtr hwndCtl);
    public const uint BM_CLICK = 0x00F5;
    public const uint WM_COMMAND = 0x0111;
    public const uint WM_CLOSE = 0x0010;
    public const uint WM_KEYDOWN = 0x0100;
    public const uint WM_KEYUP = 0x0101;
    public const int VK_RETURN = 0x0D;
  }
}
"@
}

function Get-WindowText {
  param([IntPtr]$Hwnd)
  $len = [BiDialogWatcher.Win32]::GetWindowTextLength($Hwnd)
  if ($len -le 0) { return "" }
  $sb = New-Object System.Text.StringBuilder ($len + 1)
  [BiDialogWatcher.Win32]::GetWindowText($Hwnd, $sb, $sb.Capacity) | Out-Null
  return $sb.ToString()
}

function Get-WindowClassName {
  param([IntPtr]$Hwnd)
  $sb = New-Object System.Text.StringBuilder 256
  [BiDialogWatcher.Win32]::GetClassName($Hwnd, $sb, $sb.Capacity) | Out-Null
  return $sb.ToString()
}

function Get-WindowProcessId {
  param([IntPtr]$Hwnd)
  $procId = [uint32]0
  [BiDialogWatcher.Win32]::GetWindowThreadProcessId($Hwnd, [ref]$procId) | Out-Null
  return $procId
}

# Top-level visible non-chrome windows owned by the Excel PID. Includes both
# Win32 #32770 dialogs and VBA UserForms (ThunderDFrame / F3 Server / bosa_*).
# Chrome filter strips main Excel windows, the Office ribbon, dropdown popups,
# and other non-modal Excel UI surfaces.
function Find-DialogsForPid {
  param([uint32]$ProcessId)
  $chrome = @('XLMAIN','EXCEL7','XLDESK','wndclass_desked_gsk','NUIDialog','OUTEXVBA','NetUIHWND')
  $found = New-Object System.Collections.ArrayList
  $callback = {
    param([IntPtr]$h, [IntPtr]$lp)
    if (-not [BiDialogWatcher.Win32]::IsWindowVisible($h)) { return $true }
    if ((Get-WindowProcessId -Hwnd $h) -ne $ProcessId) { return $true }
    $cls = Get-WindowClassName -Hwnd $h
    if ($chrome -contains $cls) { return $true }
    if (-not (Get-WindowText -Hwnd $h)) { return $true }
    [void]$found.Add($h)
    return $true
  }
  $delegate = [BiDialogWatcher.Win32+EnumWindowsProc]$callback
  [BiDialogWatcher.Win32]::EnumWindows($delegate, [IntPtr]::Zero) | Out-Null
  return $found
}

# Walk children: collect Static-class text into Body, Button-class controls
# into Buttons (with their hwnds for later click dispatch).
function Get-DialogPayload {
  param([IntPtr]$Hwnd)
  $statics = New-Object System.Collections.ArrayList
  $buttons = New-Object System.Collections.ArrayList
  $cb = {
    param([IntPtr]$h, [IntPtr]$lp)
    $cls = Get-WindowClassName -Hwnd $h
    $txt = Get-WindowText -Hwnd $h
    if (-not $txt) { return $true }
    if ($cls -eq 'Static') { [void]$statics.Add($txt.Trim()) }
    elseif ($cls -eq 'Button') {
      [void]$buttons.Add([pscustomobject]@{ Caption = $txt.Trim(); Hwnd = $h.ToInt64() })
    }
    return $true
  }
  $delegate = [BiDialogWatcher.Win32+EnumWindowsProc]$cb
  [BiDialogWatcher.Win32]::EnumChildWindows($Hwnd, $delegate, [IntPtr]::Zero) | Out-Null
  return [pscustomobject]@{
    Body    = ($statics -join " | ")
    Buttons = $buttons
  }
}

# Spawn the watcher runspace. Main thread will be blocked in xl.Run; this
# polls + IPCs concurrently.
function Start-DialogWatcher {
  param(
    [uint32]$ProcessId,
    [string]$ModalFile  = 'c:\tmp\excel-control-modal.json',
    [string]$ActionFile = 'c:\tmp\excel-control-action.txt',
    [int]$PollMs = 250
  )

  $tmpDir = Split-Path -Parent $ModalFile
  if ($tmpDir -and -not (Test-Path $tmpDir)) { New-Item -ItemType Directory -Path $tmpDir | Out-Null }
  Remove-Item $ModalFile  -ErrorAction SilentlyContinue
  Remove-Item $ActionFile -ErrorAction SilentlyContinue

  $state = [hashtable]::Synchronized(@{
    ProcessId  = $ProcessId
    Stop       = $false
    ModalFile  = $ModalFile
    ActionFile = $ActionFile
    Handled    = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))
  })

  $rs = [runspacefactory]::CreateRunspace(); $rs.Open()
  $rs.SessionStateProxy.SetVariable("state", $state)
  $rs.SessionStateProxy.SetVariable("PollMs", $PollMs)
  foreach ($fn in 'Get-WindowText','Get-WindowClassName','Get-WindowProcessId','Find-DialogsForPid','Get-DialogPayload') {
    $rs.SessionStateProxy.SetVariable("def_$fn", (Get-Item function:\$fn).Definition)
  }

  $ps = [powershell]::Create(); $ps.Runspace = $rs
  [void]$ps.AddScript({
    if (-not ([System.Management.Automation.PSTypeName]'BiDialogWatcher.Win32').Type) {
      Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Text;
namespace BiDialogWatcher {
  public static class Win32 {
    [DllImport("user32.dll", CharSet = CharSet.Auto)] public static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);
    [DllImport("user32.dll", CharSet = CharSet.Auto)] public static extern int GetWindowTextLength(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr hWnd);
    [DllImport("user32.dll", SetLastError = true)] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
    [DllImport("user32.dll", CharSet = CharSet.Auto)] public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")] public static extern int GetDlgCtrlID(IntPtr hwndCtl);
    public const uint BM_CLICK = 0x00F5;
    public const uint WM_COMMAND = 0x0111;
    public const uint WM_CLOSE = 0x0010;
    public const uint WM_KEYDOWN = 0x0100;
    public const uint WM_KEYUP = 0x0101;
    public const int VK_RETURN = 0x0D;
  }
}
"@
    }
    foreach ($fn in 'Get-WindowText','Get-WindowClassName','Get-WindowProcessId','Find-DialogsForPid','Get-DialogPayload') {
      $varName = "def_$fn"
      Invoke-Expression "function $fn { $((Get-Variable -Name $varName).Value) }"
    }

    while (-not $state.Stop) {
      $dialogs = Find-DialogsForPid -ProcessId $state.ProcessId
      $handled = $state.Handled
      foreach ($d in $dialogs) {
        $hwnd = [IntPtr]$d
        $hwndKey = $hwnd.ToInt64()
        if ($handled.Contains($hwndKey)) { continue }

        $caption = Get-WindowText -Hwnd $hwnd
        $payload = Get-DialogPayload -Hwnd $hwnd
        $modal = [pscustomobject]@{
          Hwnd     = $hwndKey
          Caption  = $caption
          Body     = $payload.Body
          Buttons  = $payload.Buttons
        }
        $modal | ConvertTo-Json -Depth 4 | Set-Content -Path $state.ModalFile -Encoding UTF8

        $btnList = ($payload.Buttons | ForEach-Object { $_.Caption }) -join ','
        [Console]::WriteLine("MODAL hwnd=$hwndKey caption='$caption' buttons=$btnList")

        # Wait for caller's action. Bail if the dialog gets closed externally
        # (user-clicked) — no point dispatching a click against a dead hwnd.
        while (-not (Test-Path $state.ActionFile)) {
          if ($state.Stop) { break }
          if (-not [BiDialogWatcher.Win32]::IsWindow($hwnd) -or -not [BiDialogWatcher.Win32]::IsWindowVisible($hwnd)) {
            break
          }
          Start-Sleep -Milliseconds 200
        }
        if (-not (Test-Path $state.ActionFile)) {
          [void]$handled.Add($hwndKey)
          Remove-Item $state.ModalFile -ErrorAction SilentlyContinue
          continue
        }

        $action = (Get-Content $state.ActionFile -Raw).Trim()
        Remove-Item $state.ActionFile -ErrorAction SilentlyContinue
        Remove-Item $state.ModalFile  -ErrorAction SilentlyContinue

        $target = $payload.Buttons | Where-Object { $_.Caption.Trim().ToLower() -eq $action.ToLower() } | Select-Object -First 1
        if (-not $target) {
          $target = $payload.Buttons | Select-Object -First 1
          [Console]::WriteLine("ACTION '$action' did not match; trying first button '$($target.Caption)'")
        }

        function Wait-DialogClosed([IntPtr]$h, [int]$timeoutMs) {
          $deadline = [DateTime]::UtcNow.AddMilliseconds($timeoutMs)
          while ([DateTime]::UtcNow -lt $deadline) {
            if (-not [BiDialogWatcher.Win32]::IsWindow($h)) { return $true }
            if (-not [BiDialogWatcher.Win32]::IsWindowVisible($h)) { return $true }
            Start-Sleep -Milliseconds 50
          }
          return $false
        }

        $closed = $false
        if ($target) {
          $ctrlId = [BiDialogWatcher.Win32]::GetDlgCtrlID([IntPtr]$target.Hwnd)
          if ($ctrlId -ne 0) {
            [Console]::WriteLine("CLICK '$($target.Caption)' via WM_COMMAND ctrlId=$ctrlId on dlg=$hwndKey")
            [BiDialogWatcher.Win32]::PostMessage($hwnd, [BiDialogWatcher.Win32]::WM_COMMAND, [IntPtr]$ctrlId, [IntPtr]$target.Hwnd) | Out-Null
            $closed = Wait-DialogClosed $hwnd 800
          }
          if (-not $closed) {
            [Console]::WriteLine("WM_COMMAND did not close; trying BM_CLICK on button hwnd=$($target.Hwnd)")
            [BiDialogWatcher.Win32]::PostMessage([IntPtr]$target.Hwnd, [BiDialogWatcher.Win32]::BM_CLICK, [IntPtr]::Zero, [IntPtr]::Zero) | Out-Null
            $closed = Wait-DialogClosed $hwnd 800
          }
        }
        if (-not $closed) {
          # UserForm with no Win32-enumerable buttons (Forms.CommandButton). VK_RETURN
          # fires the form's Default button click handler.
          [Console]::WriteLine("No button matched; sending VK_RETURN to dlg=$hwndKey")
          [BiDialogWatcher.Win32]::PostMessage($hwnd, [BiDialogWatcher.Win32]::WM_KEYDOWN, [IntPtr][BiDialogWatcher.Win32]::VK_RETURN, [IntPtr]0) | Out-Null
          [BiDialogWatcher.Win32]::PostMessage($hwnd, [BiDialogWatcher.Win32]::WM_KEYUP,   [IntPtr][BiDialogWatcher.Win32]::VK_RETURN, [IntPtr]0) | Out-Null
          $closed = Wait-DialogClosed $hwnd 1200
        }
        if (-not $closed) {
          [Console]::WriteLine("Falling back to WM_CLOSE on dlg=$hwndKey")
          [BiDialogWatcher.Win32]::PostMessage($hwnd, [BiDialogWatcher.Win32]::WM_CLOSE, [IntPtr]::Zero, [IntPtr]::Zero) | Out-Null
          $closed = Wait-DialogClosed $hwnd 800
        }

        if ($closed) {
          [Console]::WriteLine("DISMISSED hwnd=$hwndKey")
          [void]$handled.Add($hwndKey)
        } else {
          [Console]::WriteLine("FAILED-TO-DISMISS hwnd=$hwndKey — will retry on next poll")
        }
      }
      Start-Sleep -Milliseconds $PollMs
    }
  })

  $handle = $ps.BeginInvoke()
  return [pscustomobject]@{ State = $state; PowerShell = $ps; Handle = $handle; Runspace = $rs }
}

function Stop-DialogWatcher {
  param($Watcher)
  if (-not $Watcher) { return }
  $Watcher.State.Stop = $true
  try { [void]$Watcher.PowerShell.EndInvoke($Watcher.Handle) } catch {}
  try { $Watcher.PowerShell.Dispose() } catch {}
  try { $Watcher.Runspace.Close() } catch {}
}
