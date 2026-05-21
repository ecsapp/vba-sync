# Session-aware dialog + UserForm watcher. Dot-sourced by start-session.ps1.
#
# Runs in a background STA runspace alongside the main session loop. The
# main loop blocks inside $Xl.Run while a modal MsgBox or UserForm is up;
# this watcher detects that window and drives it so COM unblocks.
#
# Two window kinds, two interaction paths:
#
#   #32770        - standard dialog / MsgBox / VBA runtime-error box.
#                   Buttons are real child windows -> WM_COMMAND / BM_CLICK.
#
#   ThunderDFrame - MSForms UserForm. Its controls are WINDOWLESS (no
#                   child HWNDs - EnumChildWindows finds nothing). They
#                   are driven through MSAA / IAccessible (oleacc.dll),
#                   the same accessibility layer screen readers use:
#                   accDoDefaultAction clicks a button, accSelect /
#                   accDoDefaultAction picks an option button. This works
#                   cross-process while VBA is blocked on the modal form
#                   and needs no on-screen position or foreground focus.
#
# Window detection uses the window's OWN WS_VISIBLE style, not
# IsWindowVisible: a dialog owned by a hidden (headless) Excel fails
# IsWindowVisible because that walks the owner chain to the hidden main
# window. Checking the window's own style finds it either way - and the
# same test detects a UserForm dismissed via Me.Hide (window survives,
# WS_VISIBLE clears).
#
# For each new window the watcher assigns id "d<N>", writes a
# dialog_appeared / userform_appeared event (UserForms carry a controls[]
# list), and registers it so the commands below can act on it.
#
# Commands polled from commands.jsonl (owned here, not by the main loop):
#   respond_dialog    {dialog_id, button}          - click a button
#   set_form_control  {dialog_id, control, value}  - set a UserForm
#                                                    control (radio /
#                                                    checkbox / text)

$script:XcWatcherCs = @"
using System;
using System.Runtime.InteropServices;
using System.Text;

namespace XcDialog {
  [StructLayout(LayoutKind.Sequential)]
  public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }
  public static class Win32 {
    [DllImport("user32.dll", CharSet = CharSet.Auto)] public static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);
    [DllImport("user32.dll", CharSet = CharSet.Auto)] public static extern int GetWindowTextLength(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr hWnd);
    [DllImport("user32.dll", SetLastError = true)] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
    [DllImport("user32.dll", CharSet = CharSet.Auto)] public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);
    [DllImport("user32.dll")] public static extern int GetWindowLong(IntPtr hWnd, int nIndex);
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")] public static extern int GetDlgCtrlID(IntPtr hwndCtl);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr hWnd, IntPtr hdcBlt, uint nFlags);
    [DllImport("oleacc.dll")] public static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint dwId, ref Guid riid, [MarshalAs(UnmanagedType.Interface)] out object ppvObject);
    [DllImport("oleacc.dll")] public static extern int AccessibleChildren([MarshalAs(UnmanagedType.Interface)] object paccContainer, int iChildStart, int cChildren, [Out] object[] rgvarChildren, out int pcObtained);
    public const uint BM_CLICK = 0x00F5;
    public const uint WM_COMMAND = 0x0111;
    public const uint WM_CLOSE = 0x0010;
    public const int  GWL_STYLE = -16;
    public const int  WS_VISIBLE = 0x10000000;
    public const uint OBJID_CLIENT = 0xFFFFFFFC;
    public const uint PW_RENDERFULLCONTENT = 0x2;
    // -Visible mode: bring a background Excel to the foreground so a
    // modal window it deferred actually surfaces (item 2).
    [DllImport("user32.dll")] public static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("kernel32.dll")] public static extern uint GetCurrentThreadId();
    [DllImport("user32.dll", SetLastError = true)] public static extern IntPtr SendMessageTimeout(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam, uint fuFlags, uint uTimeout, out IntPtr lpdwResult);
    public const uint WM_NULL = 0x0000;
    public const uint SMTO_ABORTIFHUNG = 0x0002;
    public const int  SW_RESTORE = 9;
  }
}
"@

if (-not ([System.Management.Automation.PSTypeName]'XcDialog.Win32').Type) {
  Add-Type -TypeDefinition $script:XcWatcherCs
  Add-Type -AssemblyName System.Drawing
}

function Start-SessionDialogWatcher {
    param(
        [Parameter(Mandatory=$true)] [uint32]$ProcessId,
        [Parameter(Mandatory=$true)] [string]$EventsFile,
        [Parameter(Mandatory=$true)] [string]$CommandsFile,
        [Parameter(Mandatory=$true)] [string]$CapturesDir,
        [int]$PollMs = 200,
        [bool]$Visible = $false
    )

    $state = [hashtable]::Synchronized(@{
        Stop          = $false
        ProcessId     = $ProcessId
        EventsFile    = $EventsFile
        CommandsFile  = $CommandsFile
        CapturesDir   = $CapturesDir
        PollMs        = $PollMs
        MacroRunning  = $false
        Visible       = $Visible
        ActiveDialogs = [hashtable]::Synchronized(@{})
        DialogInfo    = [hashtable]::Synchronized(@{})
        NextId        = 1
        CmdOffset     = 0
        Cs            = $script:XcWatcherCs
    })

    $rs = [runspacefactory]::CreateRunspace()
    # STA: MSAA / IAccessible COM calls are happiest in a single-threaded
    # apartment, and the watcher only ever calls OUT to the form, so it
    # never needs to pump messages for incoming callbacks.
    $rs.ApartmentState = [System.Threading.ApartmentState]::STA
    $rs.Open()
    $rs.SessionStateProxy.SetVariable('state', $state)

    $ps = [powershell]::Create()
    $ps.Runspace = $rs

    [void]$ps.AddScript({
        # Re-declare the Win32 + oleacc interop inside the runspace.
        if (-not ([System.Management.Automation.PSTypeName]'XcDialog.Win32').Type) {
            Add-Type -TypeDefinition $state.Cs
            Add-Type -AssemblyName System.Drawing
        }

        $IID_IAccessible = [Guid]'618736e0-3c3d-11cf-810c-00aa00389b71'

        # MSAA role constants (winuser.h ROLE_SYSTEM_*).
        $ROLE_WINDOW = 9; $ROLE_CLIENT = 10; $ROLE_PANE = 16
        $ROLE_GROUPING = 20; $ROLE_PAGETABLIST = 59
        $ROLE_PUSHBUTTON = 43; $ROLE_CHECKBUTTON = 44; $ROLE_RADIOBUTTON = 45
        # MSAA state bits.
        $STATE_UNAVAILABLE = 0x1; $STATE_SELECTED = 0x2
        $STATE_CHECKED = 0x10; $STATE_INVISIBLE = 0x8000

        # -------- window helpers --------------------------------------

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
        # Is the window itself shown? Checks the window's OWN WS_VISIBLE
        # style - unlike IsWindowVisible, which also requires every owner
        # to be visible and so reports $false for a dialog owned by a
        # hidden (headless) Excel. Also returns $false once a UserForm is
        # dismissed with Me.Hide (the window survives, WS_VISIBLE clears).
        function Test-WindowShown([IntPtr]$h) {
            if (-not [XcDialog.Win32]::IsWindow($h)) { return $false }
            $style = [XcDialog.Win32]::GetWindowLong($h, [XcDialog.Win32]::GWL_STYLE)
            return (($style -band [XcDialog.Win32]::WS_VISIBLE) -ne 0)
        }

        function Save-DialogPng([int64]$Hwnd, [string]$Path) {
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

        # Top-level #32770 dialogs and ThunderDFrame UserForms owned by
        # the session's Excel that are currently shown.
        function Find-Dialogs($targetPid) {
            $found = New-Object System.Collections.ArrayList
            $cb = {
                param([IntPtr]$h, [IntPtr]$lp)
                $cls = Get-WClass -h $h
                if ($cls -ne '#32770' -and $cls -ne 'ThunderDFrame') { return $true }
                if ((Get-WPid -h $h) -ne $targetPid) { return $true }
                if (-not (Test-WindowShown -h $h)) { return $true }
                [void]$found.Add($h)
                return $true
            }
            $del = [XcDialog.Win32+EnumWindowsProc]$cb
            [XcDialog.Win32]::EnumWindows($del, [IntPtr]::Zero) | Out-Null
            return $found
        }

        # Excel's main top-level window (class XLMAIN) for the session.
        function Find-ExcelMain([uint32]$targetPid) {
            $found = New-Object System.Collections.ArrayList
            $cb = {
                param([IntPtr]$h, [IntPtr]$lp)
                if ((Get-WClass -h $h) -eq 'XLMAIN' -and (Get-WPid -h $h) -eq $targetPid) {
                    [void]$found.Add($h); return $false
                }
                return $true
            }
            $del = [XcDialog.Win32+EnumWindowsProc]$cb
            [XcDialog.Win32]::EnumWindows($del, [IntPtr]::Zero) | Out-Null
            if ($found.Count -gt 0) { return [IntPtr]$found[0] }
            return [IntPtr]::Zero
        }

        # Bring a window to the foreground, bypassing the foreground lock
        # via AttachThreadInput (no synthetic ALT key). Surfaces a modal a
        # background Excel deferred in -Visible mode.
        function Set-ExcelForeground([IntPtr]$h) {
            if ($h -eq [IntPtr]::Zero) { return }
            $fg = [XcDialog.Win32]::GetForegroundWindow()
            $ourTid = [XcDialog.Win32]::GetCurrentThreadId()
            $fgPid = [uint32]0
            $fgTid = [XcDialog.Win32]::GetWindowThreadProcessId($fg, [ref]$fgPid)
            $attached = $false
            try {
                if ($fgTid -ne 0 -and $fgTid -ne $ourTid) {
                    $attached = [XcDialog.Win32]::AttachThreadInput($ourTid, $fgTid, $true)
                }
                if ([XcDialog.Win32]::IsIconic($h)) {
                    [void][XcDialog.Win32]::ShowWindow($h, [XcDialog.Win32]::SW_RESTORE)
                }
                [void][XcDialog.Win32]::SetForegroundWindow($h)
            } finally {
                if ($attached) {
                    [void][XcDialog.Win32]::AttachThreadInput($ourTid, $fgTid, $false)
                }
            }
        }

        # -------- #32770 child-window payload -------------------------

        function Get-Payload([IntPtr]$dlg) {
            # Standard dialogs have real child windows: Static = body
            # text, Button = a clickable button.
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

        # -------- MSAA / IAccessible (UserForm controls) --------------

        function Get-AccFromWindow([IntPtr]$h) {
            # OBJID_CLIENT IAccessible for a window; $null on failure.
            $iid = [Guid]'618736e0-3c3d-11cf-810c-00aa00389b71'
            $obj = $null
            try {
                $hr = [XcDialog.Win32]::AccessibleObjectFromWindow($h, [XcDialog.Win32]::OBJID_CLIENT, [ref]$iid, [ref]$obj)
                if ($hr -eq 0 -and $obj) { return $obj }
            } catch {}
            return $null
        }

        # Recursively collect leaf controls under an IAccessible
        # container into $sink as { Name Role State Value Acc ChildId }.
        # MSForms returns each control as a full child IAccessible
        # (ChildId 0); a simple element (int childId) is a leaf of its
        # parent. Containers are recursed; window/client wrappers are
        # walked but not themselves recorded.
        function Collect-AccControls($container, [int]$depth, $sink) {
            if ($depth -gt 6 -or -not $container) { return }
            $count = 0
            try { $count = [int]$container.accChildCount } catch { return }
            if ($count -le 0) { return }
            $arr = New-Object 'object[]' $count
            $got = 0
            try { [void][XcDialog.Win32]::AccessibleChildren($container, 0, $count, $arr, [ref]$got) }
            catch { return }
            for ($i = 0; $i -lt $got; $i++) {
                $raw = $arr[$i]
                if ($null -eq $raw) { continue }
                $cAcc = $null; $cId = 0
                if ($raw -is [int]) { $cAcc = $container; $cId = [int]$raw }
                else                { $cAcc = $raw;       $cId = 0 }

                $role = 0; $name = ''; $stt = 0; $val = ''
                try { $role = [int]$cAcc.accRole($cId) }  catch {}
                try { $n = $cAcc.accName($cId);  if ($n) { $name = [string]$n } } catch {}
                try { $stt = [int]$cAcc.accState($cId) }  catch {}
                try { $v = $cAcc.accValue($cId); if ($v) { $val = [string]$v } } catch {}

                if (($stt -band $STATE_INVISIBLE) -eq 0) {
                    if ($role -ne $ROLE_WINDOW -and $role -ne $ROLE_CLIENT) {
                        [void]$sink.Add([pscustomobject]@{
                            Name = $name; Role = $role; State = $stt
                            Value = $val; Acc = $cAcc; ChildId = $cId
                        })
                    }
                }
                # Recurse into full-object containers only (a simple
                # element has no children and re-walking its parent
                # would loop).
                if ($cId -eq 0 -and ($role -eq $ROLE_WINDOW -or $role -eq $ROLE_CLIENT -or `
                    $role -eq $ROLE_PANE -or $role -eq $ROLE_GROUPING -or $role -eq $ROLE_PAGETABLIST)) {
                    Collect-AccControls $cAcc ($depth + 1) $sink
                }
            }
        }

        function Get-FormControls([IntPtr]$hwnd) {
            $sink = New-Object System.Collections.ArrayList
            $root = Get-AccFromWindow $hwnd
            if ($root) { Collect-AccControls $root 0 $sink }
            return $sink
        }

        function Get-RoleName([int]$r) {
            switch ($r) {
                43 { 'button' } 44 { 'checkbox' } 45 { 'radiobutton' }
                42 { 'text' }   46 { 'combobox' } 41 { 'label' }
                33 { 'list' }   34 { 'listitem' } 20 { 'grouping' }
                16 { 'pane' }   50 { 'spinbutton' } 60 { 'pagetab' }
                default { "role$r" }
            }
        }

        function Norm([string]$s) { return (($s -replace '&', '').Trim().ToLower()) }

        # Serialise the control list for a userform_appeared event.
        function Format-Controls($controls) {
            $arr = @()
            foreach ($c in $controls) {
                $checked = ((($c.State -band $STATE_CHECKED) -ne 0) -or (($c.State -band $STATE_SELECTED) -ne 0))
                $enabled = (($c.State -band $STATE_UNAVAILABLE) -eq 0)
                $arr += [pscustomobject]@{
                    name    = $c.Name
                    role    = (Get-RoleName ([int]$c.Role))
                    value   = $c.Value
                    checked = $checked
                    enabled = $enabled
                }
            }
            return ,$arr
        }

        function Wait-Closed([IntPtr]$h, [int]$timeoutMs) {
            $end = [DateTime]::UtcNow.AddMilliseconds($timeoutMs)
            while ([DateTime]::UtcNow -lt $end) {
                if (-not (Test-WindowShown $h)) { return $true }
                Start-Sleep -Milliseconds 50
            }
            return $false
        }

        function Write-SessionEvent([hashtable]$ev) {
            $line = $ev | ConvertTo-Json -Compress -Depth 12
            for ($i = 0; $i -lt 5; $i++) {
                try { Add-Content -LiteralPath $state.EventsFile -Value $line -Encoding UTF8; return }
                catch { Start-Sleep -Milliseconds 30 }
            }
            Add-Content -LiteralPath $state.EventsFile -Value $line -Encoding UTF8
        }

        function Read-WatcherCommands {
            # respond_dialog + set_form_control are owned by this runspace.
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
                            if ($obj.cmd -eq 'respond_dialog' -or $obj.cmd -eq 'set_form_control') {
                                [void]$out.Add($obj)
                            }
                        } catch {}
                    }
                }
                $state.CmdOffset = $fs.Position
                $sr.Dispose()
            } finally { $fs.Dispose() }
            return $out
        }

        # -------- click / set actions ---------------------------------

        # #32770: post WM_COMMAND to the button's control id, then
        # BM_CLICK, then (no match) WM_CLOSE.
        function Send-DialogClick($info, [string]$buttonLabel) {
            $hwnd = [IntPtr]$info.Hwnd
            $want = Norm $buttonLabel
            $target = $null
            foreach ($b in $info.Buttons) {
                if ((Norm $b.Caption) -eq $want) { $target = $b; break }
            }
            $closed = $false
            if ($target) {
                $ctrlId = [XcDialog.Win32]::GetDlgCtrlID([IntPtr]$target.Hwnd)
                if ($ctrlId -ne 0) {
                    [void][XcDialog.Win32]::PostMessage($hwnd, [XcDialog.Win32]::WM_COMMAND, [IntPtr]$ctrlId, [IntPtr]$target.Hwnd)
                    $closed = Wait-Closed $hwnd 800
                }
                if (-not $closed) {
                    [void][XcDialog.Win32]::PostMessage([IntPtr]$target.Hwnd, [XcDialog.Win32]::BM_CLICK, [IntPtr]::Zero, [IntPtr]::Zero)
                    $closed = Wait-Closed $hwnd 800
                }
            }
            if ($closed) { return [pscustomobject]@{ Ok = $true; Reason = $null; Available = @() } }
            if (-not $target) {
                $avail = @($info.Buttons | ForEach-Object { $_.Caption })
                return [pscustomobject]@{ Ok = $false; Reason = 'button_not_found'; Available = $avail }
            }
            return [pscustomobject]@{ Ok = $false; Reason = 'dialog_did_not_close'; Available = @() }
        }

        # ThunderDFrame: click a UserForm button via MSAA.
        function Send-FormClick([IntPtr]$hwnd, [string]$buttonLabel) {
            $controls = Get-FormControls $hwnd
            $want = Norm $buttonLabel
            $target = $null
            foreach ($c in $controls) {
                if ([int]$c.Role -ne $ROLE_PUSHBUTTON) { continue }
                if ((Norm $c.Name) -eq $want) { $target = $c; break }
            }
            if (-not $target) {
                $avail = @($controls | Where-Object { [int]$_.Role -eq $ROLE_PUSHBUTTON } | ForEach-Object { $_.Name })
                return [pscustomobject]@{ Ok = $false; Reason = 'button_not_found'; Available = $avail }
            }
            try { [void]$target.Acc.accDoDefaultAction($target.ChildId) }
            catch { return [pscustomobject]@{ Ok = $false; Reason = ('invoke_failed: ' + $_.Exception.Message); Available = @() } }
            $closed = Wait-Closed $hwnd 1500
            if ($closed) { return [pscustomobject]@{ Ok = $true; Reason = $null; Available = @() } }
            return [pscustomobject]@{ Ok = $false; Reason = 'form_did_not_close'; Available = @() }
        }

        # ThunderDFrame: set a UserForm control via MSAA. Radio buttons
        # are selected; checkboxes toggled to the requested bool; text /
        # combobox set best-effort through accValue.
        function Set-FormControl([IntPtr]$hwnd, [string]$controlName, $value) {
            $controls = Get-FormControls $hwnd
            $want = Norm $controlName
            $target = $null
            foreach ($c in $controls) {
                if ((Norm $c.Name) -eq $want) { $target = $c; break }
            }
            if (-not $target) {
                return [pscustomobject]@{ Ok = $false; Reason = 'control_not_found'
                    Available = @($controls | ForEach-Object { $_.Name }) }
            }
            $sv = ''
            if ($null -ne $value) { $sv = [string]$value }
            # Empty / true-ish means select-or-check; explicit false unchecks.
            $desired = ($sv -eq '') -or ($sv -match '^(?i:true|1|yes|on|checked|selected)$')
            $role = [int]$target.Role
            $stt  = [int]$target.State
            try {
                if ($role -eq $ROLE_RADIOBUTTON) {
                    $isSel = ((($stt -band $STATE_SELECTED) -ne 0) -or (($stt -band $STATE_CHECKED) -ne 0))
                    if (-not $isSel) { [void]$target.Acc.accDoDefaultAction($target.ChildId) }
                }
                elseif ($role -eq $ROLE_CHECKBUTTON) {
                    $isChk = (($stt -band $STATE_CHECKED) -ne 0)
                    if ($desired -ne $isChk) { [void]$target.Acc.accDoDefaultAction($target.ChildId) }
                }
                elseif ($role -eq $ROLE_PUSHBUTTON) {
                    [void]$target.Acc.accDoDefaultAction($target.ChildId)
                }
                else {
                    # text box / combobox - best effort
                    $target.Acc.accValue($target.ChildId) = $sv
                }
            } catch {
                return [pscustomobject]@{ Ok = $false; Reason = ('set_failed: ' + $_.Exception.Message); Available = @() }
            }
            Start-Sleep -Milliseconds 90
            # Re-read so the event reports the control's resulting state.
            $now = $null
            foreach ($c in (Get-FormControls $hwnd)) {
                if ((Norm $c.Name) -eq $want) { $now = $c; break }
            }
            $chk = $false; $vv = ''
            if ($now) {
                $ns = [int]$now.State
                $chk = ((($ns -band $STATE_CHECKED) -ne 0) -or (($ns -band $STATE_SELECTED) -ne 0))
                $vv = [string]$now.Value
            }
            return [pscustomobject]@{ Ok = $true; Reason = $null
                Role = (Get-RoleName $role); Checked = $chk; Value = $vv }
        }

        # -------- main loop -------------------------------------------
        $deferredPolls   = 0
        $lastActivateUtc = [DateTime]::MinValue
        while (-not $state.Stop) {
            $dialogs = Find-Dialogs -targetPid $state.ProcessId
            $stillActive = @{}
            foreach ($d in $dialogs) { $stillActive[([IntPtr]$d).ToInt64()] = $true }

            # New windows
            foreach ($d in $dialogs) {
                $hwnd = [IntPtr]$d
                $key  = $hwnd.ToInt64()
                if ($state.ActiveDialogs.ContainsKey($key)) { continue }

                $caption = Get-WText -h $hwnd
                $cls     = Get-WClass -h $hwnd
                $isForm  = ($cls -eq 'ThunderDFrame')

                $bodyText    = ''
                $buttonNames = @()
                $controlList = @()
                $childButtons = $null

                if ($isForm) {
                    $formControls = Get-FormControls $hwnd
                    $controlList  = Format-Controls $formControls
                    $buttonNames  = @($formControls | Where-Object { [int]$_.Role -eq 43 } | ForEach-Object { $_.Name })
                    $bodyText     = (@($formControls | Where-Object { [int]$_.Role -eq 41 } | ForEach-Object { $_.Name }) -join ' | ')
                } else {
                    $payload      = Get-Payload -dlg $hwnd
                    $bodyText     = $payload.Body
                    $buttonNames  = @($payload.Buttons | ForEach-Object { $_.Caption })
                    $childButtons = $payload.Buttons
                }

                $id = "d$($state.NextId)"
                $state.NextId++
                $state.ActiveDialogs[$key] = $id
                $state.DialogInfo[$id] = [pscustomobject]@{
                    Hwnd    = $hwnd.ToInt64()
                    Class   = $cls
                    Caption = $caption
                    Buttons = $childButtons   # #32770 only; $null for forms
                }

                # The VBA runtime-error dialog is a #32770 captioned
                # "Microsoft Visual Basic" with &Continue / &End / &Debug.
                # It is unrecoverable from outside the VBA runtime, so we
                # emit a structured runtime_error and auto-dispatch End.
                $isRtError = ((-not $isForm) -and ($caption -eq 'Microsoft Visual Basic'))
                $rtNum  = 0
                $rtDesc = $bodyText
                if ($isRtError -and $bodyText -match "Run-time error '(-?\d+)(?:\s*\([0-9A-Fa-f]+\))?'\s*[:\.]?\s*(.*)$") {
                    $rtNum  = [int64]$Matches[1]
                    $rtDesc = $Matches[2].Trim()
                }
                $eventType =
                    if     ($isRtError) { 'runtime_error' }
                    elseif ($isForm)    { 'userform_appeared' }
                    else                { 'dialog_appeared' }

                $shotPath = $null
                $shotErr  = $null
                try {
                    $shotPath = Join-Path $state.CapturesDir "$id.png"
                    [void](Save-DialogPng -Hwnd $hwnd.ToInt64() -Path $shotPath)
                    if (-not (Test-Path $shotPath)) { $shotPath = $null; $shotErr = 'png not written' }
                } catch { $shotPath = $null; $shotErr = $_.Exception.Message }

                $ev = @{
                    t          = $eventType
                    id         = $id
                    title      = $caption
                    text       = $bodyText
                    buttons    = $buttonNames
                    class      = $cls
                    screenshot = $shotPath
                }
                if ($isForm)   { $ev.controls = $controlList }
                if ($shotErr)  { $ev.screenshot_error = $shotErr }
                if ($isRtError) {
                    $ev.number = $rtNum
                    $ev.description = $rtDesc
                }
                Write-SessionEvent $ev

                if ($isRtError) {
                    $info = $state.DialogInfo[$id]
                    $null = Send-DialogClick $info '&End'
                    [void]$state.ActiveDialogs.Remove($key)
                    [void]$state.DialogInfo.Remove($id)
                }
            }

            # Externally-closed windows
            $toRemove = @()
            foreach ($key in @($state.ActiveDialogs.Keys)) {
                if (-not $stillActive.ContainsKey([int64]$key)) { $toRemove += $key }
            }
            foreach ($key in $toRemove) {
                $id = $state.ActiveDialogs[$key]
                [void]$state.ActiveDialogs.Remove($key)
                [void]$state.DialogInfo.Remove($id)
                Write-SessionEvent @{ t = 'dialog_closed_externally'; id = $id }
            }

            # respond_dialog / set_form_control commands
            $cmds = Read-WatcherCommands
            foreach ($cmd in $cmds) {
                Write-SessionEvent @{ t = 'command_ack'; id = $cmd.id; cmd = $cmd.cmd }
                $dialogId = $cmd.dialog_id
                $info = $state.DialogInfo[$dialogId]

                if ($cmd.cmd -eq 'respond_dialog') {
                    if (-not $info) {
                        Write-SessionEvent @{ t = 'respond_failed'; id = $cmd.id; dialog_id = $dialogId; reason = 'unknown_or_closed_dialog' }
                        continue
                    }
                    if ($info.Class -eq 'ThunderDFrame') {
                        $r = Send-FormClick ([IntPtr]$info.Hwnd) $cmd.button
                    } else {
                        $r = Send-DialogClick $info $cmd.button
                    }
                    if ($r.Ok) {
                        Write-SessionEvent @{ t = 'dialog_dismissed'; id = $cmd.id; dialog_id = $dialogId; button = $cmd.button }
                        [void]$state.ActiveDialogs.Remove([int64]$info.Hwnd)
                        [void]$state.DialogInfo.Remove($dialogId)
                    } else {
                        $ev = @{ t = 'respond_failed'; id = $cmd.id; dialog_id = $dialogId; reason = $r.Reason }
                        if ($r.Available -and $r.Available.Count -gt 0) { $ev.available = $r.Available }
                        Write-SessionEvent $ev
                    }
                }
                elseif ($cmd.cmd -eq 'set_form_control') {
                    if (-not $info) {
                        Write-SessionEvent @{ t = 'form_control_failed'; id = $cmd.id; dialog_id = $dialogId; control = $cmd.control; reason = 'unknown_or_closed_dialog' }
                        continue
                    }
                    if ($info.Class -ne 'ThunderDFrame') {
                        Write-SessionEvent @{ t = 'form_control_failed'; id = $cmd.id; dialog_id = $dialogId; control = $cmd.control; reason = 'not_a_userform' }
                        continue
                    }
                    $r = Set-FormControl ([IntPtr]$info.Hwnd) $cmd.control $cmd.value
                    if ($r.Ok) {
                        Write-SessionEvent @{ t = 'form_control_set'; id = $cmd.id; dialog_id = $dialogId
                            control = $cmd.control; role = $r.Role; checked = $r.Checked; value = $r.Value }
                    } else {
                        $ev = @{ t = 'form_control_failed'; id = $cmd.id; dialog_id = $dialogId; control = $cmd.control; reason = $r.Reason }
                        if ($r.Available -and $r.Available.Count -gt 0) { $ev.available = $r.Available }
                        Write-SessionEvent $ev
                    }
                }
            }

            # -Visible deferred-modal rescue. A background Excel can leave a
            # modal window created-but-not-shown (WS_VISIBLE clear, so
            # Find-Dialogs misses it) and the macro hangs. When a macro is
            # running, nothing is surfaced, and Excel's UI thread is pumping
            # a message loop (a modal IS up), bring Excel forward so the
            # window paints and the next poll catches it. WM_NULL times out
            # while the thread computes — that tells computing from
            # blocked-on-modal, so activation never lands mid-computation.
            if ($dialogs.Count -eq 0 -and $state.Visible -and $state.MacroRunning) {
                $xlMain = Find-ExcelMain $state.ProcessId
                $pumping = $false
                if ($xlMain -ne [IntPtr]::Zero) {
                    $res = [IntPtr]::Zero
                    $rc = [XcDialog.Win32]::SendMessageTimeout($xlMain,
                              [XcDialog.Win32]::WM_NULL, [IntPtr]::Zero, [IntPtr]::Zero,
                              [XcDialog.Win32]::SMTO_ABORTIFHUNG, 300, [ref]$res)
                    $pumping = ($rc -ne [IntPtr]::Zero)
                }
                if ($pumping) {
                    $deferredPolls++
                    $sinceMs = ([DateTime]::UtcNow - $lastActivateUtc).TotalMilliseconds
                    # 3 consecutive polls filters a transient pump (a DoEvents
                    # loop) from a genuinely stuck modal; 750ms throttle bounds
                    # re-activation when foreground is lost again mid-run.
                    if ($deferredPolls -ge 3 -and $sinceMs -gt 750) {
                        Set-ExcelForeground $xlMain
                        $lastActivateUtc = [DateTime]::UtcNow
                        Write-SessionEvent @{ t = 'dialog_activated'; polls = $deferredPolls }
                    }
                } else {
                    $deferredPolls = 0
                }
            } else {
                $deferredPolls = 0
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
