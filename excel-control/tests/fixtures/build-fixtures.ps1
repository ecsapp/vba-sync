# Generate the fixture .xlsm files used by excel-control tests.
#
# Reproducible: deleting all fixtures and re-running this script produces
# functionally equivalent workbooks. Run when fixture VBA or form layout
# needs to change.
#
# Prereqs:
#   - Excel installed (uses COM)
#   - File -> Options -> Trust Center -> Trust access to the VBA project
#     object model (vba-sync needs this too)
#
# Usage:
#   pwsh .\excel-control\tests\fixtures\build-fixtures.ps1
#
# password-locked.xlsm needs a one-time manual lock step after the script
# runs; see this directory's README. VBA exposes no COM API for setting
# the project password, and Win32/UIA/SendKeys automation of the
# Tools > VBAProject Properties dialog proved unreliable from a
# non-interactive PowerShell session in testing.

[CmdletBinding()]
param(
    [string]$OutDir = $PSScriptRoot
)

$ErrorActionPreference = 'Stop'

# vbext_ComponentType
$VBEXT_CT_STDMODULE = 1
$VBEXT_CT_MSFORM    = 3
# xlFileFormat
$XL_OPEN_XML_WORKBOOK_MACRO_ENABLED = 52

function New-Excel {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $xl.AskToUpdateLinks = $false
    return $xl
}

function Save-Macro-Workbook($wb, [string]$Path) {
    if (Test-Path $Path) { Remove-Item -LiteralPath $Path -Force }
    $wb.SaveAs($Path, $XL_OPEN_XML_WORKBOOK_MACRO_ENABLED)
}

function Add-StdModule($wb, [string]$Name, [string]$Code) {
    $comp = $wb.VBProject.VBComponents.Add($VBEXT_CT_STDMODULE)
    $comp.Name = $Name
    $comp.CodeModule.AddFromString($Code)
}

# ---------- fixtures ----------

function Build-Empty([string]$Path) {
    $xl = New-Excel
    try {
        $wb = $xl.Workbooks.Add()
        Save-Macro-Workbook $wb $Path
        $wb.Close($false)
    } finally { $xl.Quit() | Out-Null }
    Write-Host "  $Path" -ForegroundColor Green
}

function Build-MsgBox([string]$Path) {
    $xl = New-Excel
    try {
        $wb = $xl.Workbooks.Add()
        $code = @'
Option Explicit

Public Sub PopMsgBox()
    MsgBox "Hello", vbOKCancel, "msgbox fixture"
End Sub
'@
        Add-StdModule $wb 'modTest' $code
        Save-Macro-Workbook $wb $Path
        $wb.Close($false)
    } finally { $xl.Quit() | Out-Null }
    Write-Host "  $Path" -ForegroundColor Green
}

function Build-RuntimeError([string]$Path) {
    $xl = New-Excel
    try {
        $wb = $xl.Workbooks.Add()
        $code = @'
Option Explicit

' Raises Err 9 (Subscript out of range) by indexing past the array bound.
Public Sub TriggerError()
    Dim a(1) As String
    a(99) = "x"
End Sub
'@
        Add-StdModule $wb 'modTest' $code
        Save-Macro-Workbook $wb $Path
        $wb.Close($false)
    } finally { $xl.Quit() | Out-Null }
    Write-Host "  $Path" -ForegroundColor Green
}

function Build-UserForm([string]$Path) {
    $xl = New-Excel
    try {
        $wb = $xl.Workbooks.Add()
        $proj = $wb.VBProject

        $form = $proj.VBComponents.Add($VBEXT_CT_MSFORM)
        $form.Name = 'frmLogin'
        $designer = $form.Designer
        $form.Properties.Item('Caption').Value = 'Login'
        $form.Properties.Item('Width').Value  = 240
        $form.Properties.Item('Height').Value = 160

        $lblU = $designer.Controls.Add('Forms.Label.1', 'lblUsername')
        $lblU.Left = 12; $lblU.Top = 12; $lblU.Width = 60; $lblU.Caption = 'Username:'

        $txtU = $designer.Controls.Add('Forms.TextBox.1', 'Username')
        $txtU.Left = 80; $txtU.Top = 10; $txtU.Width = 130; $txtU.Height = 18

        $lblP = $designer.Controls.Add('Forms.Label.1', 'lblPassword')
        $lblP.Left = 12; $lblP.Top = 40; $lblP.Width = 60; $lblP.Caption = 'Password:'

        $txtP = $designer.Controls.Add('Forms.TextBox.1', 'Password')
        $txtP.Left = 80; $txtP.Top = 38; $txtP.Width = 130; $txtP.Height = 18
        $txtP.PasswordChar = '*'

        $chk = $designer.Controls.Add('Forms.CheckBox.1', 'RememberMe')
        $chk.Left = 80; $chk.Top = 66; $chk.Width = 130; $chk.Caption = 'Remember me'

        $btnOK = $designer.Controls.Add('Forms.CommandButton.1', 'cmdOK')
        $btnOK.Left = 50; $btnOK.Top = 100; $btnOK.Width = 70; $btnOK.Height = 24
        $btnOK.Caption = 'OK'
        $btnOK.Default = $true

        $btnCancel = $designer.Controls.Add('Forms.CommandButton.1', 'cmdCancel')
        $btnCancel.Left = 130; $btnCancel.Top = 100; $btnCancel.Width = 70; $btnCancel.Height = 24
        $btnCancel.Caption = 'Cancel'
        $btnCancel.Cancel = $true

        # Wire OK/Cancel click handlers in the form's own code module.
        $formCode = @'
Option Explicit

Private Sub cmdOK_Click()
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub
'@
        $form.CodeModule.AddFromString($formCode)

        $stdCode = @'
Option Explicit

Public Sub ShowLoginForm()
    frmLogin.Show
End Sub
'@
        Add-StdModule $wb 'modTest' $stdCode

        Save-Macro-Workbook $wb $Path
        $wb.Close($false)
    } finally { $xl.Quit() | Out-Null }
    Write-Host "  $Path" -ForegroundColor Green
}

function Build-PasswordLockedSource([string]$Path) {
    # Builds the workbook content; the lock itself must be applied
    # manually one time via Excel UI. See README in this directory.
    $xl = New-Excel
    try {
        $wb = $xl.Workbooks.Add()
        $code = @'
Option Explicit

Public Sub PopMsgBox()
    MsgBox "Hello", vbOKCancel, "password-locked fixture"
End Sub
'@
        Add-StdModule $wb 'modTest' $code
        Save-Macro-Workbook $wb $Path
        $wb.Close($false)
    } finally { $xl.Quit() | Out-Null }
    Write-Host "  $Path (UNLOCKED — manual lock step required, see README)" -ForegroundColor Yellow
}

function Build-Obscured([string]$SrcPath, [string]$Path) {
    if (-not (Test-Path $SrcPath)) { throw "Source msgbox fixture missing: $SrcPath" }
    Copy-Item -LiteralPath $SrcPath -Destination $Path -Force
    Write-Host "  $Path (copy of msgbox.xlsm)" -ForegroundColor Green
}

# ---------- run ----------

Write-Host "Building fixtures into $OutDir" -ForegroundColor Cyan
# Kill orphan Excel processes from prior failed runs to avoid file locks.
Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Milliseconds 300

$paths = @{
    Empty    = Join-Path $OutDir 'empty.xlsm'
    MsgBox   = Join-Path $OutDir 'msgbox.xlsm'
    Runtime  = Join-Path $OutDir 'runtime-error.xlsm'
    Form     = Join-Path $OutDir 'userform.xlsm'
    Locked   = Join-Path $OutDir 'password-locked.xlsm'
    Obscured = Join-Path $OutDir 'obscured.xlsm'
}

Build-Empty               $paths.Empty
Build-MsgBox              $paths.MsgBox
Build-RuntimeError        $paths.Runtime
Build-UserForm            $paths.Form
Build-PasswordLockedSource $paths.Locked
Build-Obscured            $paths.MsgBox $paths.Obscured

# Final sweep — any orphan Excel left over from automation
Start-Sleep -Milliseconds 400
Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "Done. Fixtures in $OutDir" -ForegroundColor Cyan
Get-ChildItem $OutDir -Filter *.xlsm | Format-Table Name, Length, LastWriteTime -AutoSize

Write-Host ""
Write-Host "If you rebuilt password-locked.xlsm, you must manually re-apply" -ForegroundColor Yellow
Write-Host "the lock (see README in this directory)." -ForegroundColor Yellow
