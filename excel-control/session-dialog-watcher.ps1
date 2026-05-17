# Session-aware dialog watcher. Dot-sourced by start-session.ps1.
#
# Runs in a background runspace alongside the main session loop. Polls
# for #32770 dialogs and UserForm windows owned by Excel's PID. For each
# new dialog:
#   1. Assigns an id "d<N>"
#   2. Writes a `dialog_appeared` (or `userform_appeared`) event
#   3. Registers the hwnd in a shared map so respond_dialog can find it
#
# Independently polls commands.jsonl for `respond_dialog` commands. On
# match: looks up the hwnd by dialog_id, dispatches a click via
# WM_COMMAND → BM_CLICK → VK_RETURN → WM_CLOSE fallback chain. Emits
# `dialog_dismissed` on success, `respond_failed` on miss.
#
# If a dialog disappears externally (user closed it, macro ended), emits
# `dialog_closed_externally` and frees the id.

if (-not ([System.Management.Automation.PSTypeName]'XcDialog.Win32').Type) {
  Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Text;

namespace XcDialog {
  [StructLayout(LayoutKind.Sequential)]
  public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }
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
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr hWnd, IntPtr hdcBlt, uint nFlags);
    public const uint BM_CLICK = 0x00F5;
    public const uint WM_COMMAND = 0x0111;
    public const uint WM_CLOSE = 0x0010;
    public const uint WM_KEYDOWN = 0x0100;
    public const uint WM_KEYUP = 0x0101;
    public const int VK_RETURN = 0x0D;
    public const uint PW_RENDERFULLCONTENT = 0x2;
  }
}
"@
  Add-Type -AssemblyName System.Drawing
}

function Start-SessionDialogWatcher {
    param(
        [Parameter(Mandatory=$true)] [uint32]$ProcessId,
        [Parameter(Mandatory=$true)] [string]$EventsFile,
        [Parameter(Mandatory=$true)] [string]$CommandsFile,
        [Parameter(Mandatory=$true)] [string]$CapturesDir,
        [int]$PollMs = 200
    )

    $state = [hashtable]::Synchronized(@{
        Stop          = $false
        ProcessId     = $ProcessId
        EventsFile    = $EventsFile
        CommandsFile  = $CommandsFile
        CapturesDir   = $CapturesDir
        PollMs        = $PollMs
        ActiveDialogs = [hashtable]::Synchronized(@{})
        DialogInfo    = [hashtable]::Synchronized(@{})
        NextId        = 1
        CmdOffset     = 0
    })

    $rs = [runspacefactory]::CreateRunspace()
    $rs.Open()
    $rs.SessionStateProxy.SetVariable('state', $state)

    $ps = [powershell]::Create()
    $ps.Runspace = $rs

    [void]$ps.AddScript({
        # Re-declare Win32 inside the runspace (includes capture types)
        if (-not ([System.Management.Automation.PSTypeName]'XcDialog.Win32').Type) {
            Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Text;
namespace XcDialog {
  [StructLayout(LayoutKind.Sequential)]
  public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }
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
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr hWnd, IntPtr hdcBlt, uint nFlags);
    public const uint BM_CLICK = 0x00F5;
    public const uint WM_COMMAND = 0x0111;
    public const uint WM_CLOSE = 0x0010;
    public const uint WM_KEYDOWN = 0x0100;
    public const uint WM_KEYUP = 0x0101;
    public const int VK_RETURN = 0x0D;
    public const uint PW_RENDERFULLCONTENT = 0x2;
  }
}
"@
            Add-Type -AssemblyName System.Drawing
        }

        function Capture-WindowPng([int64]$Hwnd, [string]$Path) {
            $h = [IntPtr]$Hwnd
            if (-not [XcDialog.Win32]::IsWindow($h)) { return $null }
            $rect = New-Object XcDialog.RECT
            [void][XcDialog.Win32]::GetWindowRect($h, [ref]$rect)
            $w = $rect.Right - $rect.Left
            $hgt = $rect.Bottom - $rect.Top
            if ($w -le 0 -or $hgt -le 0) { return $null }
            $bmp = New-Object System.Drawing.Bitmap $w, $hgt
            $g   = [System.Drawing.Graphics]::FromImage($bmp)
            $hdc = $g.GetHdc()
            try { [void][XcDialog.Win32]::PrintWindow($h, $hdc, [XcDialog.Win32]::PW_RENDERFULLCONTENT) }
            finally { $g.ReleaseHdc($hdc); $g.Dispose() }
            $bmp.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
            $bmp.Dispose()
            return $Path
        }

        function Get-WText([IntPtr]$h) {
            $len = [XcDialog.Win32]::GetWindowTextLength($h)
            if ($len -le 0) { return '' }
            $sb = New-Object System.Text.StringBuilder ($len + 1)
            [XcDialog.Win32]::GetWindowText($h, $sb, $sb.Capacity) | Out-Null
            return $sb.ToString()
        }
        function Get-WClass([IntPtr]$h) {
            $sb = New-Object System.Text.StringBuilder 256
            [XcDialog.Win32]::GetClassName($h, $sb, $sb.Capacity) | Out-Null
            return $sb.ToString()
        }
        function Get-WPid([IntPtr]$h) {
            $p = [uint32]0
            [XcDialog.Win32]::GetWindowThreadProcessId($h, [ref]$p) | Out-Null
            return $p
        }

        function Find-Dialogs($targetPid) {
            $chrome = @('XLMAIN','EXCEL7','XLDESK','wndclass_desked_gsk','NUIDialog','OUTEXVBA','NetUIHWND','MsoCommandBar','MsoCommandBarPopup','MsoWorkPane')
            $found = New-Object System.Collections.ArrayList
            $cb = {
                param([IntPtr]$h, [IntPtr]$lp)
                if (-not [XcDialog.Win32]::IsWindowVisible($h)) { return $true }
                if ((Get-WPid -h $h) -ne $targetPid) { return $true }
                $cls = Get-WClass -h $h
                if ($chrome -contains $cls) { return $true }
                if (-not (Get-WText -h $h)) { return $true }
                [void]$found.Add($h)
                return $true
            }
            $del = [XcDialog.Win32+EnumWindowsProc]$cb
            [XcDialog.Win32]::EnumWindows($del, [IntPtr]::Zero) | Out-Null
            return $found
        }

        function Get-Payload([IntPtr]$dlg) {
            $statics = New-Object System.Collections.ArrayList
            $buttons = New-Object System.Collections.ArrayList
            $cb = {
                param([IntPtr]$h, [IntPtr]$lp)
                $cls = Get-WClass -h $h
                $txt = Get-WText -h $h
                if (-not $txt) { return $true }
                if ($cls -eq 'Static') { [void]$statics.Add($txt.Trim()) }
                elseif ($cls -eq 'Button') {
                    [void]$buttons.Add([pscustomobject]@{ Caption = $txt.Trim(); Hwnd = $h.ToInt64() })
                }
                return $true
            }
            $del = [XcDialog.Win32+EnumWindowsProc]$cb
            [XcDialog.Win32]::EnumChildWindows($dlg, $del, [IntPtr]::Zero) | Out-Null
            return [pscustomobject]@{ Body = ($statics -join ' | '); Buttons = $buttons }
        }

        function Wait-Closed([IntPtr]$h, [int]$timeoutMs) {
            $end = [DateTime]::UtcNow.AddMilliseconds($timeoutMs)
            while ([DateTime]::UtcNow -lt $end) {
                if (-not [XcDialog.Win32]::IsWindow($h)) { return $true }
                if (-not [XcDialog.Win32]::IsWindowVisible($h)) { return $true }
                Start-Sleep -Milliseconds 50
            }
            return $false
        }

        function Append-Event([hashtable]$ev) {
            $line = $ev | ConvertTo-Json -Compress -Depth 10
            for ($i = 0; $i -lt 5; $i++) {
                try { Add-Content -LiteralPath $state.EventsFile -Value $line -Encoding UTF8; return }
                catch { Start-Sleep -Milliseconds 30 }
            }
            Add-Content -LiteralPath $state.EventsFile -Value $line -Encoding UTF8
        }

        function Read-RespondCommands {
            $out = New-Object System.Collections.ArrayList
            if (-not (Test-Path $state.CommandsFile)) { return $out }
            $fs = [System.IO.File]::Open($state.CommandsFile, 'Open', 'Read', 'ReadWrite')
            try {
                if ($state.CmdOffset -gt $fs.Length) { $state.CmdOffset = 0 }
                $fs.Position = $state.CmdOffset
                $sr = New-Object System.IO.StreamReader($fs, [System.Text.Encoding]::UTF8, $true, 1024, $true)
                while (-not $sr.EndOfStream) {
                    $line = $sr.ReadLine()
                    if ($null -eq $line) { break }
                    $trim = $line.Trim()
                    if ($trim) {
                        try {
                            $obj = $trim | ConvertFrom-Json
                            if ($obj.cmd -eq 'respond_dialog') { [void]$out.Add($obj) }
                        } catch {}
                    }
                }
                $state.CmdOffset = $fs.Position
                $sr.Dispose()
            } finally { $fs.Dispose() }
            return $out
        }

        function Dispatch-Click($dialogInfo, [string]$buttonLabel) {
            $hwnd = [IntPtr]$dialogInfo.Hwnd
            $target = $null
            foreach ($b in $dialogInfo.Buttons) {
                if ($b.Caption.Trim().ToLower() -eq $buttonLabel.ToLower()) { $target = $b; break }
            }
            $closed = $false
            if ($target) {
                $ctrlId = [XcDialog.Win32]::GetDlgCtrlID([IntPtr]$target.Hwnd)
                if ($ctrlId -ne 0) {
                    [XcDialog.Win32]::PostMessage($hwnd, [XcDialog.Win32]::WM_COMMAND, [IntPtr]$ctrlId, [IntPtr]$target.Hwnd) | Out-Null
                    $closed = Wait-Closed -h $hwnd -timeoutMs 800
                }
                if (-not $closed) {
                    [XcDialog.Win32]::PostMessage([IntPtr]$target.Hwnd, [XcDialog.Win32]::BM_CLICK, [IntPtr]::Zero, [IntPtr]::Zero) | Out-Null
                    $closed = Wait-Closed -h $hwnd -timeoutMs 800
                }
            }
            if (-not $closed) {
                # UserForm with Forms.CommandButton — VK_RETURN fires Default click handler
                [XcDialog.Win32]::PostMessage($hwnd, [XcDialog.Win32]::WM_KEYDOWN, [IntPtr][XcDialog.Win32]::VK_RETURN, [IntPtr]0) | Out-Null
                [XcDialog.Win32]::PostMessage($hwnd, [XcDialog.Win32]::WM_KEYUP,   [IntPtr][XcDialog.Win32]::VK_RETURN, [IntPtr]0) | Out-Null
                $closed = Wait-Closed -h $hwnd -timeoutMs 1200
            }
            if (-not $closed) {
                [XcDialog.Win32]::PostMessage($hwnd, [XcDialog.Win32]::WM_CLOSE, [IntPtr]::Zero, [IntPtr]::Zero) | Out-Null
                $closed = Wait-Closed -h $hwnd -timeoutMs 800
            }
            return $closed
        }

        # ---- main loop ----
        while (-not $state.Stop) {
            $dialogs = Find-Dialogs -targetPid $state.ProcessId
            $stillActive = @{}
            foreach ($d in $dialogs) { $stillActive[([IntPtr]$d).ToInt64()] = $true }

            # New dialogs
            foreach ($d in $dialogs) {
                $hwnd = [IntPtr]$d
                $key  = $hwnd.ToInt64()
                if ($state.ActiveDialogs.ContainsKey($key)) { continue }

                $caption = Get-WText -h $hwnd
                $cls     = Get-WClass -h $hwnd
                $payload = Get-Payload -dlg $hwnd

                $id = "d$($state.NextId)"
                $state.NextId++
                $state.ActiveDialogs[$key] = $id
                $state.DialogInfo[$id] = [pscustomobject]@{
                    Hwnd    = $hwnd.ToInt64()
                    Buttons = $payload.Buttons
                    Caption = $caption
                }

                $buttonNames = @($payload.Buttons | ForEach-Object { $_.Caption })
                # Classify the dialog. VBA's "Microsoft Visual Basic"
                # runtime-error dialog has a distinctive body
                # "Run-time error 'N':\n\n<description>" and buttons
                # &Continue / &End / &Debug / &Help. We turn that into a
                # structured `runtime_error` event AND auto-dispatch End
                # (the macro is unrecoverable from outside).
                $isRtError = ($caption -eq 'Microsoft Visual Basic' -and $payload.Body -match "Run-time error '(\d+)'\s*[:\.]?\s*(.*)$")
                $eventType =
                    if ($isRtError) { 'runtime_error' }
                    elseif ($cls -eq 'ThunderDFrame' -or $cls -like 'F3*' -or $cls -like 'bosa*') { 'userform_appeared' }
                    else { 'dialog_appeared' }

                $shotPath = $null
                $shotErr  = $null
                try {
                    $shotPath = Join-Path $state.CapturesDir "$id.png"
                    [void](Capture-WindowPng -Hwnd $hwnd.ToInt64() -Path $shotPath)
                    if (-not (Test-Path $shotPath)) { $shotPath = $null; $shotErr = 'png not written' }
                } catch { $shotPath = $null; $shotErr = $_.Exception.Message }

                $ev = @{
                    t = $eventType
                    id = $id
                    title = $caption
                    text = $payload.Body
                    buttons = $buttonNames
                    class = $cls
                    screenshot = $shotPath
                }
                if ($shotErr) { $ev.screenshot_error = $shotErr }
                if ($isRtError) {
                    $ev.number = [int]$Matches[1]
                    $ev.description = $Matches[2].Trim()
                }
                Append-Event $ev

                # For runtime errors, immediately dispatch End so the macro
                # returns to COM and PowerShell unblocks. The agent can't
                # respond to a runtime-error mid-macro anyway.
                if ($isRtError) {
                    $info = $state.DialogInfo[$id]
                    $null = Dispatch-Click -dialogInfo $info -buttonLabel '&End'
                    [void]$state.ActiveDialogs.Remove($key)
                    [void]$state.DialogInfo.Remove($id)
                }
            }

            # Externally-closed dialogs
            $toRemove = @()
            foreach ($key in @($state.ActiveDialogs.Keys)) {
                if (-not $stillActive.ContainsKey([int64]$key)) { $toRemove += $key }
            }
            foreach ($key in $toRemove) {
                $id = $state.ActiveDialogs[$key]
                [void]$state.ActiveDialogs.Remove($key)
                [void]$state.DialogInfo.Remove($id)
                Append-Event @{ t = 'dialog_closed_externally'; id = $id }
            }

            # respond_dialog commands
            $cmds = Read-RespondCommands
            foreach ($cmd in $cmds) {
                Append-Event @{ t = 'command_ack'; id = $cmd.id; cmd = 'respond_dialog' }
                $dialogId = $cmd.dialog_id
                $info = $state.DialogInfo[$dialogId]
                if (-not $info) {
                    Append-Event @{ t = 'respond_failed'; id = $cmd.id; dialog_id = $dialogId; reason = 'unknown_or_closed_dialog' }
                    continue
                }
                $closed = Dispatch-Click -dialogInfo $info -buttonLabel $cmd.button
                if ($closed) {
                    Append-Event @{ t = 'dialog_dismissed'; id = $cmd.id; dialog_id = $dialogId; button = $cmd.button }
                    [void]$state.ActiveDialogs.Remove([int64]$info.Hwnd)
                    [void]$state.DialogInfo.Remove($dialogId)
                } else {
                    Append-Event @{ t = 'respond_failed'; id = $cmd.id; dialog_id = $dialogId; reason = 'dialog_did_not_close' }
                }
            }

            Start-Sleep -Milliseconds $state.PollMs
        }
    })

    $handle = $ps.BeginInvoke()
    return [pscustomobject]@{
        State      = $state
        PowerShell = $ps
        Handle     = $handle
        Runspace   = $rs
    }
}

function Stop-SessionDialogWatcher($Watcher) {
    if (-not $Watcher) { return }
    $Watcher.State.Stop = $true
    try { [void]$Watcher.PowerShell.EndInvoke($Watcher.Handle) } catch {}
    try { $Watcher.PowerShell.Dispose() } catch {}
    try { $Watcher.Runspace.Close() } catch {}
}
