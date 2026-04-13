param(
    [string]$Path
)

$ErrorActionPreference = "Stop"

function Show-ErrorMessage {
    param([string]$Message)

    Add-Type -AssemblyName System.Windows.Forms | Out-Null
    [void][System.Windows.Forms.MessageBox]::Show(
        $Message,
        "MarkItDown",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
}

function Invoke-ConsoleToolNoWindow {
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [Parameter(Mandatory = $true)][string]$Arguments
    )

    $pinfo = New-Object System.Diagnostics.ProcessStartInfo
    $pinfo.FileName = $FilePath
    $pinfo.Arguments = $Arguments
    $pinfo.UseShellExecute = $false
    $pinfo.CreateNoWindow = $true
    $pinfo.RedirectStandardOutput = $true
    $pinfo.RedirectStandardError = $true

    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = $pinfo
    [void]$p.Start()
    $null = $p.StandardOutput.ReadToEnd()
    $stderr = $p.StandardError.ReadToEnd()
    $p.WaitForExit()

    return @{
        ExitCode = $p.ExitCode
        StdErr   = $stderr
    }
}

if (-not $Path) {
    Show-ErrorMessage "No input file was passed to the context-menu script."
    exit 1
}

if ($Path.IndexOf('"', [System.StringComparison]::Ordinal) -ge 0) {
    Show-ErrorMessage "Invalid path: contains a double-quote character."
    exit 1
}

$pythonCmd = Get-Command python -ErrorAction SilentlyContinue | Select-Object -First 1
$markitdownCmd = Get-Command markitdown -ErrorAction SilentlyContinue | Select-Object -First 1

if (-not $pythonCmd -and -not $markitdownCmd) {
    Show-ErrorMessage "Could not find `"python`" or `"markitdown`".`n`nInstall MarkItDown first (for example: pip install `"markitdown[all]`")."
    exit 1
}

try {
    $inputPath = (Resolve-Path -LiteralPath $Path).Path
}
catch {
    Show-ErrorMessage "Input file not found:`n$Path"
    exit 1
}

$outputDir = [System.IO.Path]::GetDirectoryName($inputPath)
$outputBase = [System.IO.Path]::GetFileNameWithoutExtension($inputPath)
$finalOutput = Join-Path $outputDir ($outputBase + ".md")
$tempOutput = Join-Path $outputDir ($outputBase + ".tmp-markitdown-" + [guid]::NewGuid().ToString("N") + ".md")

$converted = $false
$lastStdErr = ""

if ($pythonCmd) {
    $argsLine = "-m markitdown `"$inputPath`" -o `"$tempOutput`""
    $r = Invoke-ConsoleToolNoWindow -FilePath $pythonCmd.Source -Arguments $argsLine
    $lastStdErr = $r.StdErr
    if ($r.ExitCode -eq 0 -and (Test-Path -LiteralPath $tempOutput)) {
        $converted = $true
    }
}

if (-not $converted -and $markitdownCmd) {
    $argsLine = "`"$inputPath`" -o `"$tempOutput`""
    $r = Invoke-ConsoleToolNoWindow -FilePath $markitdownCmd.Source -Arguments $argsLine
    $lastStdErr = $r.StdErr
    if ($r.ExitCode -eq 0 -and (Test-Path -LiteralPath $tempOutput)) {
        $converted = $true
    }
}

if (-not $converted) {
    Remove-Item -LiteralPath $tempOutput -ErrorAction SilentlyContinue
    $detail = ""
    if (-not [string]::IsNullOrWhiteSpace($lastStdErr)) {
        $snippet = $lastStdErr.Trim()
        if ($snippet.Length -gt 900) { $snippet = $snippet.Substring(0, 897) + "..." }
        $detail = "`n`n" + $snippet
    }
    Show-ErrorMessage ("Conversion failed for:`n$inputPath" + $detail)
    exit 1
}

try {
    Move-Item -LiteralPath $tempOutput -Destination $finalOutput -Force
}
catch {
    Remove-Item -LiteralPath $tempOutput -ErrorAction SilentlyContinue
    Show-ErrorMessage ("Could not write output file:`n$finalOutput`n`n" + $_.Exception.Message)
    exit 1
}
