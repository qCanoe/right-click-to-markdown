<#!
  Installs or removes a "Convert to Markdown" MarkItDown entry on the Windows
  right-click menu for all file types (current user, no admin).

  Usage:
    .\\install-markitdown-context-menu.ps1              # install
    .\\install-markitdown-context-menu.ps1 -Uninstall  # remove

  Requires: markitdown-convert.ps1 + markitdown-convert-launcher.vbs in the same folder.
  At runtime: python and/or markitdown available on PATH (see markitdown-convert.ps1).
#>
param(
    [switch]$Uninstall
)

$ErrorActionPreference = "Stop"

$verbKey = "Registry::HKEY_CURRENT_USER\Software\Classes\*\shell\MarkItDown"
$commandKey = "$verbKey\command"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ps1Path = Join-Path $scriptDir "markitdown-convert.ps1"
$vbsPath = Join-Path $scriptDir "markitdown-convert-launcher.vbs"
if (-not (Test-Path -LiteralPath $ps1Path)) {
    throw "Missing companion script: $ps1Path"
}
if (-not (Test-Path -LiteralPath $vbsPath)) {
    throw "Missing companion script: $vbsPath"
}

$ps1Path = (Resolve-Path -LiteralPath $ps1Path).Path
$vbsPath = (Resolve-Path -LiteralPath $vbsPath).Path

if ($Uninstall) {
    if (Test-Path -LiteralPath $verbKey) {
        Remove-Item -LiteralPath $verbKey -Recurse
        Write-Host "Removed MarkItDown context menu entry."
    } else {
        Write-Host "MarkItDown context menu entry was not installed."
    }
    exit 0
}

New-Item -Path $verbKey -Force | Out-Null
# ASCII-only label avoids mojibake (Explorer expects UTF-16 LE in registry; mixed encoding in .ps1 can garble CJK).
New-ItemProperty -Path $verbKey -Name "(default)" -PropertyType String -Value "Convert to Markdown (MarkItDown)" -Force | Out-Null

# Explorer shell verbs should pass the selected file as "%1".
# wscript.exe runs the launcher without a console; the VBS uses WshShell.Run style 0 (hidden) for PowerShell.
$commandValue = "wscript.exe `"$vbsPath`" `"%1`""
New-Item -Path $commandKey -Force | Out-Null
New-ItemProperty -Path $commandKey -Name "(default)" -PropertyType String -Value $commandValue -Force | Out-Null

Write-Host "Installed MarkItDown context menu for current user."
Write-Host "  Handler: $commandValue"
Write-Host ""
Write-Host "Restart Explorer or log off/on if the menu does not appear immediately."
Write-Host "Uninstall: .\install-markitdown-context-menu.ps1 -Uninstall"
