# excel-control debug-capture helper (item 3).
#
# Runs as a short-lived SEPARATE process with its own clean COM
# apartment. The dialog watcher spawns it after clicking Debug on a VBA
# runtime-error dialog: driving COM into a break-mode Excel from the
# watcher's own STA runspace destabilises and crashes Excel, so the
# break-site read and the VBE Reset must happen out-of-process.
#
# Attaches to the harness Excel by a PID-guarded ROT lookup (never a
# foreign Excel instance), reads the VBA break site (module / line /
# surrounding source), drives a VBE Reset to end the break, writes the
# result as JSON to -OutFile, exits.
#
# Best-effort by design: a Bug-4-class corrupt Excel may crash mid-read;
# whatever was captured is written, the rest comes back as an `error`.

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)] [int]$ExcelPid,
    [Parameter(Mandatory=$true)] [string]$OutFile
)

$result = [ordered]@{ ok = $false; reset = $false }

if (-not ([System.Management.Automation.PSTypeName]'XcDbg.Win32').Type) {
    Add-Type @"
using System;
using System.Runtime.InteropServices;
namespace XcDbg {
  public static class Win32 {
    [DllImport("user32.dll", SetLastError=true)] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
    [DllImport("oleaut32.dll", PreserveSig=false)] public static extern void GetActiveObject(ref Guid rclsid, IntPtr reserved, [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);
  }
}
"@
}

# The harness Excel via the ROT, verified by PID — never a foreign Excel.
function Get-HarnessXl([int]$targetPid) {
    $clsid = [Guid]'00024500-0000-0000-C000-000000000046'  # Excel.Application
    $obj = $null
    [XcDbg.Win32]::GetActiveObject([ref]$clsid, [IntPtr]::Zero, [ref]$obj)
    if (-not $obj) { return $null }
    $p = [uint32]0
    [void][XcDbg.Win32]::GetWindowThreadProcessId([IntPtr][int64]$obj.Hwnd, [ref]$p)
    if ($p -ne $targetPid) { return $null }
    return $obj
}

# Depth-first search of the VBE command bars for the Reset control.
function Find-ResetControl($controls) {
    foreach ($ctl in $controls) {
        try { if ((($ctl.Caption -replace '&','').Trim()) -eq 'Reset') { return $ctl } } catch {}
        $sub = $null
        try { $sub = $ctl.Controls } catch {}
        if ($sub) { $r = Find-ResetControl $sub; if ($r) { return $r } }
    }
    return $null
}

try {
    $xl = Get-HarnessXl $ExcelPid
    if (-not $xl) {
        $result.error = 'could not attach to the harness Excel (PID-guarded ROT lookup)'
    } else {
        # --- Tier A2: structured break-site read ---
        try {
            $pane = $xl.VBE.ActiveCodePane
            if ($pane) {
                $cm = $pane.CodeModule
                $sL = 0; $sC = 0; $eL = 0; $eC = 0
                [void]$pane.GetSelection([ref]$sL, [ref]$sC, [ref]$eL, [ref]$eC)
                $a = [Math]::Max(1, $sL - 3)
                $b = [Math]::Min($cm.CountOfLines, $sL + 3)
                $ctx = @()
                for ($i = $a; $i -le $b; $i++) {
                    $mark = $(if ($i -eq $sL) { '>>' } else { '  ' })
                    $ctx += ('{0} {1}: {2}' -f $mark, $i, $cm.Lines($i, 1))
                }
                $result.module         = $cm.Name
                $result.line           = $sL
                $result.source_context = ($ctx -join "`n")
                $result.ok             = $true
            } else {
                $result.error = 'no ActiveCodePane (Excel not in break mode, or already crashed)'
            }
        } catch {
            $result.error = "break-site read failed: $($_.Exception.Message)"
        }
        # --- recovery: drive a VBE Reset to end the break ---
        try {
            foreach ($bar in $xl.VBE.CommandBars) {
                $r = Find-ResetControl $bar.Controls
                if ($r) { $r.Execute(); $result.reset = $true; break }
            }
        } catch { $result.reset = $false }
    }
} catch {
    $result.error = "debug-capture failed: $($_.Exception.Message)"
}

try {
    $result | ConvertTo-Json -Compress -Depth 6 | Set-Content -LiteralPath $OutFile -Encoding UTF8
} catch {}
