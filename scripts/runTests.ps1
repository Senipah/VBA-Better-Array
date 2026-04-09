param(
    [string]$WorkbookPath = (Join-Path $PSScriptRoot "..\\dev\\BetterArray.xlsm"),
    [string]$ModuleFilter = "",
    [string]$TestNamePattern = "*"
)

$resolvedWorkbookPath = (Resolve-Path -Path $WorkbookPath).Path
$excel = $null
$workbook = $null
$exitCode = 0

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    # msoAutomationSecurityLow
    $excel.AutomationSecurity = 1

    $workbook = $excel.Workbooks.Open($resolvedWorkbookPath)
    $macroName = "'$($workbook.Name)'!TestRunner.RunAllTests_Report"

    Write-Host "Running VBA tests from $resolvedWorkbookPath"
    $report = [string]$excel.Run($macroName, $ModuleFilter, $TestNamePattern)
    Write-Host $report
    if ($report -match "Failed:\s+0\b") {
        Write-Host "All tests passed."
    }
    else {
        throw "Test failures detected."
    }
}
catch {
    $exitCode = 1
    Write-Error "Test run failed: $($_.Exception.Message)"
}
finally {
    if ($workbook -ne $null) {
        try { $workbook.Close($false) } catch {}
    }

    if ($excel -ne $null) {
        try { $excel.Quit() } catch {}
    }

    if ($workbook -ne $null) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    }
    if ($excel -ne $null) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

exit $exitCode
