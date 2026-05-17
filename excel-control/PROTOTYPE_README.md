# vba-sync test harness (local, .scratch/ — gitignored)

Headless Excel automation harness for iterating on the vba-sync export pipeline
without a human in the loop. Used to find compile errors, time-budget regressions,
and runtime failures inside `modExcelExport.DoExportExcelStructure`.

## Files

- `dialog-watcher.ps1` — Win32 modal-dialog watcher in a separate PS runspace.
  Polls for `#32770` dialogs owned by our Excel PID, captures their static-text
  content, then dismisses via `WM_COMMAND IDCANCEL` (= "End" for VBA error
  dialogs, "Cancel" for system prompts). Sourced from
  `~/Dev/vba-dataset/tools/dialog-watcher.ps1` — keep in sync if you tweak it.

- `harness.ps1` — main runner. Rebuilds `VBA Sync.xlam` by stripping +
  re-importing all `.bas/.cls/.frm` from `VBA Sync/`, opens a target workbook
  read-only, then calls `modExcelExport.DoExportExcelStructure` directly via
  `Application.Run` (bypasses MsgBox-heavy `DoExportProject`). Reads
  `.export_error.log` and `.export_timing.log` from the target's `Excel/` dir
  to report what happened.

- `find-compile-error.ps1` — diagnostic. Rebuilds the .xlam, forces a full
  compile via the VBE "Compile VBAProject" command (CommandBar control 578),
  and on failure reads `VBE.ActiveCodePane` to print the offending module +
  line + source context. Use when the harness reports a captured "Compile
  error" dialog but you need to know WHICH line.

## Usage

```powershell
# Pinpoint a compile error after editing .bas:
& .\.scratch\find-compile-error.ps1

# End-to-end run against a target workbook:
& .\.scratch\harness.ps1 -Target 'C:\path\to\Whatever.xlsm'
```

## Why these matter

VBA's normal MsgBox + runtime-error dialogs block COM automation. Without the
watcher, a hung dialog hangs the PowerShell COM thread forever. With it:
- Errors get captured as text and surfaced to the caller
- Excel can run unattended

The harness also bypasses `DoExportProject`'s success-MsgBox by calling
`DoExportExcelStructure` directly. The result: a full export + artifact
inspection in one PowerShell invocation, suitable for an agent to iterate on.

## Critical knobs in the harness

- `xl.AutomationSecurity = 1` (default) — leave this alone. Setting it to 3
  (`msoAutomationSecurityForceDisable`) breaks the addin's own macros too, so
  `Application.Run` fails with "macros disabled".
- `xl.EnableEvents = $false` — prevents the target's `Workbook_Open` from
  firing. Important for workbooks with auto-handlers.
- `xl.Hwnd → GetWindowProcessId` — reliable way to learn our Excel's PID
  (avoids race conditions with other Excel instances).

## Future extraction

This pattern (dialog watcher + rebuild-and-run) applies to every VBA project
Arnaud maintains. Once it has stabilised across vba-sync, vba-dataset, p6-vba,
extract to a standalone repo (working name: `excel-com-harness`). At that
point, vendor the watcher as a single .ps1 module that callers dot-source.

To make extraction easy:
- Both scripts have NO project-specific path assumptions in code — all paths
  come from parameters with overridable defaults
- `dialog-watcher.ps1` is fully self-contained (no external types)
- `harness.ps1` parameters: `-Target`, `-AddinPath`, `-SourceRoot` — generic
- `find-compile-error.ps1` parameters: `-AddinPath`, `-SourceRoot`

## Known caveats

- The watcher dismisses ALL modals owned by Excel. If you want to inspect a
  particular dialog manually, comment out the `Dismiss-Dialog` call inside the
  watcher's runspace loop and re-run.
- After a harness run, an orphan EXCEL.EXE can survive if `xl.Quit()` raises.
  The harness force-kills its own PID at the end, but if your test instance
  is stuck, `Stop-Process -Name EXCEL` (be careful — kills all Excel).
- `find-compile-error.ps1` relies on VBE positioning the cursor on the broken
  line. If compile succeeds (no dialog), the "no active code pane" message is
  the green-flag.
