#requires -Version 5.1
Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

$basePath = if ([string]::IsNullOrWhiteSpace($PSScriptRoot)) { (Get-Location).Path } else { $PSScriptRoot }
$sourcePath = Join-Path $basePath "RespaldoCorreosLauncher.cs"
$exePath = Join-Path $basePath "RespaldoCorreos.exe"

if (-not (Test-Path -LiteralPath $sourcePath)) {
    throw "No se encontro $sourcePath"
}

if (Test-Path -LiteralPath $exePath) {
    Remove-Item -LiteralPath $exePath -Force
}

Add-Type `
    -TypeDefinition (Get-Content -LiteralPath $sourcePath -Raw) `
    -ReferencedAssemblies @("System.Windows.Forms.dll") `
    -OutputAssembly $exePath `
    -OutputType WindowsApplication

Write-Host "EXE creado: $exePath"
