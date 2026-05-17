# Open a workbook, run vba-sync ImportProject via COM, save, close.
#
# WHY
#   The vba-sync addin's manual flow is: open the workbook, click "Import
#   Project" on the ribbon, click OK on the success dialog, Ctrl+S. Fine for
#   humans but doesn't compose with CI / agent workflows that want a one-line
#   "bake the latest committed .bas/.cls/.frm into the .xlsm" step before
#   pushing to a client or running a test.
#
#   This script automates that flow. Net effect:
#     - Workbook's VBProject is wiped and re-imported from <Workbook>/Modules,
#       <Workbook>/ClassModules, <Workbook>/Forms (vba-sync's source folder
#       convention).
#     - Workbook is saved.
#
# KEY GOTCHAS THIS SCRIPT HANDLES
#   1. ReadOnlyRecommended: many production workbooks set this flag (a "you
#      probably shouldn't edit this" hint). Excel COM honours it and opens
#      the workbook read-only by default, which makes Save trip a
#      SaveAs/NUIDialog the watcher can't dismiss. Pass IgnoreReadOnlyRecommended=True
#      on Workbooks.Open to bypass.
#   2. VBA project password: if VBProject.Protection = 1, ImportProject
#      can't enumerate components. Caller passes -VbaPassword; we route
#      through unlock-vba-project.ps1 to unlock before importing.
#   3. vba-sync's success modal: ImportProject ends on a MsgBox the caller
#      has to dismiss for xl.Run to return cleanly. We start the
#      bidirectional dialog watcher and auto-OK that single known modal
#      (it's always "OK"-only, no decision to surface).
#
# USAGE
#   .\excel-control\sync-vba-to-workbook.ps1 `
#       -WorkbookPath "C:\path\to\Workbook.xlsm" `
#       -VbaPassword 'secret'
#
#   Or import the script and call Sync-VbaToWorkbook directly:
#     . .\excel-control\sync-vba-to-workbook.ps1
#     Sync-VbaToWorkbook -WorkbookPath '...' -VbaPassword '...'

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)] [string]$WorkbookPath,
  [string]$VbaPassword   = '',
  [string]$AddinPath     = (Join-Path $env:APPDATA 'Microsoft\AddIns\VBA Sync.xlam'),
  [string]$ModalFile     = 'c:\tmp\excel-control-modal.json',
  [string]$ActionFile    = 'c:\tmp\excel-control-action.txt',
  [switch]$Trace
)

$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot 'unlock-vba-project.ps1')
. (Join-Path $PSScriptRoot 'bidirectional-dialog-watcher.ps1')

function Sync-VbaToWorkbook {
  param(
    [Parameter(Mandatory=$true)] [string]$WorkbookPath,
    [string]$VbaPassword = '',
    [string]$AddinPath   = (Join-Path $env:APPDATA 'Microsoft\AddIns\VBA Sync.xlam'),
    [string]$ModalFile   = 'c:\tmp\excel-control-modal.json',
    [string]$ActionFile  = 'c:\tmp\excel-control-action.txt',
    [switch]$Trace
  )

  $resolved = (Resolve-Path -LiteralPath $WorkbookPath -ErrorAction Stop).Path
  if (-not (Test-Path $AddinPath)) { throw "VBA Sync.xlam not found at $AddinPath" }

  $xl      = $null
  $wb      = $null
  $addin   = $null
  $watcher = $null
  try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $true       # required for VBE / Names lookups; UserForm modals
    $xl.DisplayAlerts = $false
    $xl.AskToUpdateLinks = $false

    if ($Trace) { Write-Host "Opening $resolved (read-write, IgnoreReadOnlyRecommended)" -ForegroundColor Cyan }
    # Workbooks.Open params 1-7: Filename, UpdateLinks, ReadOnly, Format, Password,
    # WriteResPassword, IgnoreReadOnlyRecommended. The 7th = True is the only fix
    # for workbooks that set the ReadOnlyRecommended flag internally.
    $wb = $xl.Workbooks.Open($resolved, [Type]::Missing, $false, [Type]::Missing, [Type]::Missing, [Type]::Missing, $true)
    if ($wb.ReadOnly) {
      throw "Workbook opened read-only after Open. File may be locked by another process — close it elsewhere and rerun."
    }

    try { $null = $wb.VBProject.VBComponents.Count } catch {
      throw "Enable Trust Center -> Trust access to the VBA project object model"
    }

    if ($wb.VBProject.Protection -eq 1) {
      if (-not $VbaPassword) { throw "VBA project is password-protected; pass -VbaPassword to unlock." }
      if ($Trace) { Write-Host "Unlocking VBA project..." -ForegroundColor Cyan }
      $unlocked = Unlock-VbaProject -Xl $xl -Password $VbaPassword
      if (-not $unlocked) { throw "Could not unlock VBA project — wrong password or VBE not responding." }
    }

    $excelPid = [uint32]0
    [BiDialogWatcher.Win32]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd, [ref]$excelPid) | Out-Null
    $watcher = Start-DialogWatcher -ProcessId $excelPid -ModalFile $ModalFile -ActionFile $ActionFile

    # Auto-dismiss the single known "Import completed successfully" modal
    # vba-sync emits at the end of ImportProject. Spawned as a background
    # job that writes "OK" to ActionFile when ModalFile appears, then exits.
    $autoOkJob = Start-Job -ArgumentList $ModalFile, $ActionFile -ScriptBlock {
      param($mf, $af)
      $deadline = (Get-Date).AddSeconds(120)
      while ((Get-Date) -lt $deadline) {
        if (Test-Path $mf) { Set-Content -LiteralPath $af -Value 'OK' -Encoding UTF8; break }
        Start-Sleep -Milliseconds 200
      }
    }

    if ($Trace) { Write-Host "Loading VBA Sync addin: $AddinPath" -ForegroundColor Cyan }
    $addin = $xl.Workbooks.Open($AddinPath)
    $wb.Activate()

    if ($Trace) { Write-Host "Calling VBA Sync.xlam!modSync.ImportProject" -ForegroundColor Cyan }
    $xl.Run("'VBA Sync.xlam'!modSync.ImportProject", $null)

    try { Stop-Job $autoOkJob -ErrorAction SilentlyContinue; Remove-Job $autoOkJob -Force -ErrorAction SilentlyContinue } catch {}

    if ($Trace) { Write-Host "Saving workbook..." -ForegroundColor Cyan }
    $wb.Save()
    Write-Host "OK: $resolved synced + saved" -ForegroundColor Green
  }
  finally {
    if ($watcher) { try { Stop-DialogWatcher $watcher } catch {} }
    if ($wb)    { try { $wb.Close($false) }    catch {} }
    if ($addin) { try { $addin.Close($false) } catch {} }
    if ($xl)    {
      try { $xl.Quit() } catch {}
      try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null } catch {}
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
  }
}

# When invoked as a script (not dot-sourced), run the main function.
if ($MyInvocation.InvocationName -ne '.') {
  $params = @{ WorkbookPath = $WorkbookPath; AddinPath = $AddinPath; ModalFile = $ModalFile; ActionFile = $ActionFile }
  if ($VbaPassword) { $params['VbaPassword'] = $VbaPassword }
  if ($Trace)       { $params['Trace']       = $true }
  Sync-VbaToWorkbook @params
}
