param(
    [string]$WorkbookPath = (Join-Path $PSScriptRoot "..\\dev\\BetterArray.xlsm"),
    [string]$ModuleFilter = "",
    [string]$TestNamePattern = "*"
)

$ErrorActionPreference = "Stop"

Write-Host "[devCheck] Rebuilding dev workbook..."
$createWorkbookArgs = @(
    "-ExecutionPolicy", "Bypass",
    "-File", (Join-Path $PSScriptRoot "createDevWorkbook.ps1")
)
& powershell @createWorkbookArgs
if ($LASTEXITCODE -ne 0) {
    throw "[devCheck] createDevWorkbook failed with exit code $LASTEXITCODE."
}

Write-Host "[devCheck] Running tests..."
$runTestsArgs = @(
    "-ExecutionPolicy", "Bypass",
    "-File", (Join-Path $PSScriptRoot "runTests.ps1"),
    "-WorkbookPath", $WorkbookPath,
    "-TestNamePattern", $TestNamePattern
)
if (-not [string]::IsNullOrWhiteSpace($ModuleFilter)) {
    $runTestsArgs += @("-ModuleFilter", $ModuleFilter)
}
& powershell @runTestsArgs

if ($LASTEXITCODE -ne 0) {
    throw "[devCheck] runTests failed with exit code $LASTEXITCODE."
}

Write-Host "[devCheck] Completed successfully."
